"""
Lógica pura de descarga de carpetas SharePoint.

No tiene dependencias de FastAPI: puede ser importado directamente por otros
servicios o consumido a través del router en app/api/descarga.py.

AUTENTICACIÓN ANÓNIMA:
  SharePoint devuelve cookies de sesión al seguir la URL compartida (resolver()).
  Esas cookies deben viajar en TODAS las peticiones posteriores del mismo trabajo.
  Por eso se crea UNA sola sesión por llamada y se pasa explícitamente a cada función.

ESTRATEGIA DE DESCARGA SELECTIVA:
  La función principal del flujo de validación es descargar_carpeta_doc():
  en lugar de bajar toda la carpeta raíz (que puede incluir fotos, visitas, etc.),
  lista solo el primer nivel, localiza la subcarpeta 00_DOCUMENTACION en SharePoint
  y descarga únicamente esa subcarpeta. Esto evita descargar archivos irrelevantes.
"""

import logging
import re
import time
import urllib.parse
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from app.config import (
    CHUNK_SIZE,
    DOWNLOADS_DIR,
    MAX_WORKERS_DESCARGA,
    MAX_WORKERS_LISTADO,
)
from app.core.normalizacion import normalizar

logger = logging.getLogger(__name__)


# ── Creación de sesión ────────────────────────────────────────────────────────

def _nueva_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    adapter = HTTPAdapter(
        pool_connections=4,
        pool_maxsize=MAX_WORKERS_DESCARGA + MAX_WORKERS_LISTADO + 4,
        max_retries=Retry(
            total=3,
            backoff_factor=1.0,
            status_forcelist={500, 502, 503, 504},
            allowed_methods={"GET"},
        ),
    )
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


# ── Resolución de URL ─────────────────────────────────────────────────────────

def resolver(sess: requests.Session, url: str) -> tuple[str, str, str]:
    """
    Sigue redirecciones de la URL compartida de SharePoint.
    Las cookies de sesión anónima quedan guardadas en `sess`.
    Retorna: (host, web_raiz, ruta_relativa)
    """
    r = sess.get(url, allow_redirects=True, timeout=60)
    r.raise_for_status()

    final = urllib.parse.urlparse(r.url)
    host  = final.netloc
    q     = urllib.parse.parse_qs(final.query)

    rel = q["id"][0] if "id" in q else None
    if not rel:
        m   = re.search(r"id=(%2[fF]personal[^&\"']+)", r.text)
        rel = urllib.parse.unquote(m.group(1)) if m else None
    if not rel:
        raise RuntimeError("No pude extraer la ruta de la carpeta desde la URL.")

    rel = urllib.parse.unquote(rel)
    web = "/".join(rel.split("/")[:3])
    return host, web, rel


# ── Listado de carpetas ───────────────────────────────────────────────────────

def _listar_carpeta(
    sess: requests.Session, host: str, web: str, rel: str
) -> tuple[list[str], list[str]]:
    """Retorna (subcarpetas, archivos) como listas de ServerRelativeUrl."""
    rq  = urllib.parse.quote(rel, safe="/")
    api = (
        f"https://{host}{web}/_api/web/GetFolderByServerRelativeUrl('{rq}')"
        f"?$expand=Folders,Files"
        f"&$select=Folders/ServerRelativeUrl,Files/ServerRelativeUrl"
    )
    r = sess.get(api, headers={"Accept": "application/json;odata=verbose"}, timeout=60)

    if r.status_code != 200:
        raise RuntimeError(f"API REST {r.status_code}: {r.text[:300]}")

    d        = r.json()["d"]
    carpetas = [f["ServerRelativeUrl"] for f in d["Folders"]["results"]]
    archivos = [f["ServerRelativeUrl"] for f in d["Files"]["results"]]
    return carpetas, archivos


def _buscar_carpeta_doc_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """
    Busca recursivamente en SharePoint (sin descargar nada) la subcarpeta
    cuyo nombre normalizado contenga '00' Y 'documentacion'.
    Retorna la ServerRelativeUrl de esa carpeta, o None si no existe.
    """
    pendientes = [rel_raiz]
    while pendientes:
        actual    = pendientes.pop(0)
        try:
            carpetas, _ = _listar_carpeta(sess, host, web, actual)
        except Exception as exc:
            logger.warning("Error listando '%s': %s", actual, exc)
            continue

        for rel_sub in carpetas:
            nombre = rel_sub.rstrip("/").split("/")[-1]
            n      = normalizar(nombre)
            if "00" in n and "documentacion" in n:
                return rel_sub
            pendientes.append(rel_sub)

    return None


def _listar_arbol(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> list[str]:
    """BFS paralelo: retorna ServerRelativeUrl de todos los archivos bajo rel_raiz."""
    todos_archivos: list[str] = []
    nivel_actual = [rel_raiz]

    with ThreadPoolExecutor(max_workers=MAX_WORKERS_LISTADO) as pool:
        while nivel_actual:
            futuros = {
                pool.submit(_listar_carpeta, sess, host, web, rel): rel
                for rel in nivel_actual
            }
            nivel_actual = []
            for fut in as_completed(futuros):
                rel = futuros[fut]
                try:
                    carpetas, archivos = fut.result()
                    todos_archivos.extend(archivos)
                    nivel_actual.extend(carpetas)
                except Exception as exc:
                    logger.warning("Error listando '%s': %s", rel, exc)

    return todos_archivos


# ── Descarga de archivos ──────────────────────────────────────────────────────

def _srel_a_path(srel: str, rel_raiz: str, base_local: Path) -> Path:
    relativa = srel[len(rel_raiz):].lstrip("/")
    return base_local / Path(*relativa.split("/")) if relativa else base_local


def _bajar_archivo(
    sess: requests.Session, host: str, srel: str, destino: Path
) -> int:
    if destino.exists() and destino.stat().st_size > 0:
        return destino.stat().st_size

    nombre = srel.split("/")[-1]
    url    = f"https://{host}{urllib.parse.quote(srel, safe='/')}"
    destino.parent.mkdir(parents=True, exist_ok=True)

    for intento in range(3):
        try:
            logger.debug("  → %s", nombre)
            with sess.get(url, stream=True, timeout=(30, 60)) as r:
                r.raise_for_status()
                with open(destino, "wb") as fh:
                    for chunk in r.iter_content(CHUNK_SIZE):
                        if chunk:
                            fh.write(chunk)
            size = destino.stat().st_size
            logger.debug("  ✓ %s (%.1f KB)", nombre, size / 1024)
            return size
        except Exception as exc:
            logger.warning("Reintento %d/3 — %s: %s", intento + 1, nombre, exc)
            if destino.exists():
                destino.unlink(missing_ok=True)
            time.sleep(1 << intento)

    logger.error("Fallo definitivo: %s", nombre)
    return -1


def _descargar_arbol(
    sess: requests.Session, host: str, web: str,
    rel_raiz: str, base_local: Path
) -> dict:
    """Lista y descarga en paralelo todos los archivos bajo rel_raiz."""
    todos_archivos = _listar_arbol(sess, host, web, rel_raiz)
    total = len(todos_archivos)
    logger.info("%d archivos encontrados. Iniciando descarga paralela.", total)

    exitosos = 0
    fallidos  = 0

    def _tarea(srel: str) -> int:
        return _bajar_archivo(sess, host, srel, _srel_a_path(srel, rel_raiz, base_local))

    with ThreadPoolExecutor(max_workers=MAX_WORKERS_DESCARGA) as pool:
        futuros = {pool.submit(_tarea, srel): srel for srel in todos_archivos}
        for fut in as_completed(futuros):
            srel = futuros[fut]
            try:
                size = fut.result()
                if size >= 0:
                    exitosos += 1
                else:
                    fallidos += 1
            except Exception as exc:
                fallidos += 1
                logger.error("Error en %s: %s", srel, exc)

            completados = exitosos + fallidos
            if completados % 10 == 0 or completados == total:
                logger.info(
                    "  Descarga: %d/%d archivos  |  %d fallidos",
                    completados, total, fallidos,
                )

    return {"total": total, "exitosos": exitosos, "fallidos": fallidos}


# ── Puntos de entrada públicos ────────────────────────────────────────────────

def descargar_carpeta_doc(url: str, dest_base: Path | None = None) -> tuple[Path | None, dict]:
    """
    Función principal del flujo de validación.

    1. Resuelve la URL y obtiene cookies de sesión.
    2. Busca la subcarpeta 00_DOCUMENTACION en SharePoint (sin descargar nada más).
    3. Descarga SOLO esa subcarpeta.

    Esto evita descargar carpetas irrelevantes (fotos de visita, caracterización, etc.)
    que pueden ser pesadas y no aportan nada a la validación documental.

    Returns:
        (ruta_local_doc, stats) — ruta_local_doc es None si no se encontró la carpeta.
    """
    dest_base  = dest_base or DOWNLOADS_DIR / str(uuid.uuid4())
    sess       = _nueva_sesion()

    host, web, rel_raiz = resolver(sess, url)

    logger.info("Buscando 00_DOCUMENTACION en SharePoint: %s", rel_raiz)
    rel_doc = _buscar_carpeta_doc_en_sharepoint(sess, host, web, rel_raiz)

    if rel_doc is None:
        logger.warning("No se encontró 00_DOCUMENTACION en: %s", rel_raiz)
        return None, {"total": 0, "exitosos": 0, "fallidos": 0}

    logger.info("Carpeta doc encontrada: %s", rel_doc)
    nombre_doc = urllib.parse.unquote(rel_doc.rstrip("/").split("/")[-1])
    ruta_local = dest_base / nombre_doc
    ruta_local.mkdir(parents=True, exist_ok=True)

    stats = _descargar_arbol(sess, host, web, rel_doc, ruta_local)
    return ruta_local, stats


def descargar_carpeta(url: str, dest_base: Path | None = None) -> tuple[Path, dict]:
    """
    Descarga la carpeta raíz completa (para uso externo / endpoint /descargar).
    El flujo de validación usa descargar_carpeta_doc() en su lugar.
    """
    dest_base  = dest_base or DOWNLOADS_DIR / str(uuid.uuid4())
    sess       = _nueva_sesion()

    host, web, rel = resolver(sess, url)

    nombre     = urllib.parse.unquote(rel.rstrip("/").split("/")[-1])
    ruta_local = dest_base / nombre
    ruta_local.mkdir(parents=True, exist_ok=True)

    logger.info("Listando árbol completo: %s", rel)
    stats = _descargar_arbol(sess, host, web, rel, ruta_local)
    return ruta_local, stats
