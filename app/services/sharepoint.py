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

import socket
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from app.config import (
    CHUNK_SIZE,
    DOWNLOADS_DIR,
    MAX_WORKERS_DESCARGA,
    MAX_WORKERS_LISTADO,
    TIMEOUT_DESCARGA,
)
from app.core.normalizacion import normalizar

logger = logging.getLogger(__name__)


# ── Creación de sesión ────────────────────────────────────────────────────────

# Timeout global de socket: evita que iter_content se quede pegado entre chunks
# aunque requests no lo detecte. Aplica a todas las conexiones del proceso.
socket.setdefaulttimeout(TIMEOUT_DESCARGA)


def _nueva_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    adapter = HTTPAdapter(
        pool_connections=4,
        pool_maxsize=MAX_WORKERS_DESCARGA + MAX_WORKERS_LISTADO + 4,
        max_retries=Retry(
            total=3,
            backoff_factor=2.0,
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
    r = sess.get(url, allow_redirects=True, timeout=TIMEOUT_DESCARGA)
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
    r = sess.get(api, headers={"Accept": "application/json;odata=verbose"}, timeout=TIMEOUT_DESCARGA)

    if r.status_code != 200:
        raise RuntimeError(f"API REST {r.status_code}: {r.text[:300]}")

    d        = r.json()["d"]
    carpetas = [f["ServerRelativeUrl"] for f in d["Folders"]["results"]]
    archivos = [f["ServerRelativeUrl"] for f in d["Files"]["results"]]
    return carpetas, archivos


def _buscar_subcarpeta(
    sess: requests.Session, host: str, web: str, rel_raiz: str,
    criterio,   # callable(nombre_normalizado: str) -> bool
) -> str | None:
    """
    BFS en SharePoint (sin descargar) buscando una subcarpeta que cumpla el criterio.
    Retorna la ServerRelativeUrl de la primera coincidencia, o None.
    """
    pendientes = [rel_raiz]
    while pendientes:
        actual = pendientes.pop(0)
        try:
            carpetas, _ = _listar_carpeta(sess, host, web, actual)
        except Exception as exc:
            logger.warning("Error listando '%s': %s", actual, exc)
            continue
        for rel_sub in carpetas:
            nombre = normalizar(rel_sub.rstrip("/").split("/")[-1])
            if criterio(nombre):
                return rel_sub
            pendientes.append(rel_sub)
    return None


def _buscar_carpeta_doc_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """Localiza la subcarpeta 00_DOCUMENTACION (nombre contiene '00' y 'documentacion')."""
    return _buscar_subcarpeta(
        sess, host, web, rel_raiz,
        lambda n: "00" in n and "documentacion" in n,
    )


def _buscar_carpeta_visita_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """Localiza la subcarpeta 01_VISITA (nombre contiene '01' y 'visita' o 'caracterizacion')."""
    return _buscar_subcarpeta(
        sess, host, web, rel_raiz,
        lambda n: "01" in n and ("visita" in n or "caracterizacion" in n),
    )

def _buscar_carpeta_visita2_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """Localiza la subcarpeta 02_VISITA_2_DIAGNOSTICO."""
    return _buscar_subcarpeta(
        sess, host, web,  rel_raiz,
        lambda n: "02" in n and ("visita" in n or "diagnostico" in n),
    )


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

    # Timeout por chunk: si no llega ningún byte en CHUNK_STALL_TIMEOUT segundos
    # durante la descarga, se considera bloqueo y se reintenta.
    CHUNK_STALL_TIMEOUT = 60

    for intento in range(3):
        try:
            logger.debug("  → %s", nombre)
            with sess.get(url, stream=True, timeout=(30, TIMEOUT_DESCARGA)) as r:
                r.raise_for_status()
                with open(destino, "wb") as fh:
                    deadline = time.monotonic() + CHUNK_STALL_TIMEOUT
                    for chunk in r.iter_content(CHUNK_SIZE):
                        if chunk:
                            fh.write(chunk)
                            deadline = time.monotonic() + CHUNK_STALL_TIMEOUT
                        elif time.monotonic() > deadline:
                            raise TimeoutError(
                                f"Sin datos por más de {CHUNK_STALL_TIMEOUT}s (posible bloqueo)"
                            )
            size = destino.stat().st_size
            logger.debug("  ✓ %s (%.1f KB)", nombre, size / 1024)
            return size
        except Exception as exc:
            logger.warning("Reintento %d/3 — %s: %s", intento + 1, nombre, exc)
            if destino.exists():
                destino.unlink(missing_ok=True)
            time.sleep(2 ** intento)  # 1s, 2s, 4s

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
    Valida la carpeta 00_DOCUMENTACION SIN descargar archivos.

    Solo lista los nombres de archivo en SharePoint para detectar
    CEDULA / RUT / COMERCIO / TENENCIA por nombre — no se descarga
    ningún archivo, evitando timeouts con imágenes pesadas.

    Escribe un archivo .txt con los nombres encontrados para que
    _listar_archivos y _detectar_documentos puedan trabajar normalmente.

    Returns:
        (ruta_local_doc, stats) — ruta_local_doc es None si no se encontró la carpeta.
    """
    dest_base = dest_base or DOWNLOADS_DIR / str(uuid.uuid4())
    sess      = _nueva_sesion()

    host, web, rel_raiz = resolver(sess, url)

    logger.info("Buscando 00_DOCUMENTACION en SharePoint: %s", rel_raiz)
    rel_doc = _buscar_carpeta_doc_en_sharepoint(sess, host, web, rel_raiz)

    if rel_doc is None:
        logger.warning("No se encontró 00_DOCUMENTACION en: %s", rel_raiz)
        return None, {"total": 0, "exitosos": 0, "fallidos": 0}

    logger.info("Carpeta doc encontrada: %s", rel_doc)

    # Listar archivos sin descargar (solo nombres)
    try:
        _, archivos_srel = _listar_carpeta(sess, host, web, rel_doc)
    except Exception as exc:
        logger.warning("Error listando 00_DOCUMENTACION '%s': %s", rel_doc, exc)
        return None, {"total": 0, "exitosos": 0, "fallidos": 0}

    # Crear carpeta local con archivos "fantasma" (tamaño 0) que solo sirven
    # para que _listar_archivos detecte los nombres
    nombre_doc = urllib.parse.unquote(rel_doc.rstrip("/").split("/")[-1])
    ruta_local = dest_base / nombre_doc
    ruta_local.mkdir(parents=True, exist_ok=True)

    total = len(archivos_srel)
    for srel in archivos_srel:
        nombre_archivo = urllib.parse.unquote(srel.rstrip("/").split("/")[-1])
        (ruta_local / nombre_archivo).touch()   # archivo vacío — solo el nombre importa

    logger.info(
        "  %d archivos listados en 00 (sin descarga): %s",
        total, [urllib.parse.unquote(s.split("/")[-1]) for s in archivos_srel],
    )
    return ruta_local, {"total": total, "exitosos": total, "fallidos": 0}


def descargar_visita_selectiva(
    url: str, dest_base: Path
) -> dict:
    """
    Procesa la carpeta 01_VISITA_1_CARACTERIZACION de forma selectiva:
      1. Localiza la subcarpeta 01_VISITA en SharePoint.
      2. Lista su contenido (un solo nivel + recursivo para archivos).
      3. Cuenta archivos de tipo imagen/video SIN descargarlos.
      4. Descarga SOLO los 3 documentos obligatorios (compromiso, visita, tratamiento).

    Retorna un dict con:
      encontrada       → bool
      conteo_media     → int (cantidad de fotos/videos)
      docs_rutas       → {keyword: Path | None}  rutas locales de los 3 docs
      stats            → {total_descargados, exitosos, fallidos}
    """
    from app.core.normalizacion import normalizar as _norm
    from app.core.reglas import EXTENSIONES_MEDIA, PALABRAS_CLAVE_VISITA

    resultado = {
        "encontrada":   False,
        "conteo_media": 0,
        "docs_rutas":   {k: None for k in PALABRAS_CLAVE_VISITA},
        "stats":        {"total": 0, "exitosos": 0, "fallidos": 0},
    }

    sess            = _nueva_sesion()
    host, web, rel  = resolver(sess, url)

    rel_visita = _buscar_carpeta_visita_en_sharepoint(sess, host, web, rel)
    if rel_visita is None:
        return resultado

    resultado["encontrada"] = True

    # Listar todos los archivos de la carpeta (recursivo BFS)
    todos_archivos = _listar_arbol(sess, host, web, rel_visita)

    conteo_media   = 0
    a_descargar    = {}  # {keyword: srel}

    for srel in todos_archivos:
        nombre = srel.rstrip("/").split("/")[-1]
        ext    = Path(nombre).suffix.lower()
        stem   = _norm(Path(nombre).stem)

        # Contar media (sin descargar)
        if ext in EXTENSIONES_MEDIA:
            conteo_media += 1
            continue

        # Identificar los 3 documentos obligatorios
        for keyword, palabras in PALABRAS_CLAVE_VISITA.items():
            if keyword not in a_descargar and any(p in stem for p in palabras):
                a_descargar[keyword] = srel
                break

    resultado["conteo_media"] = conteo_media

    if not a_descargar:
        return resultado

    # Descargar solo los documentos identificados
    dest_base.mkdir(parents=True, exist_ok=True)
    exitosos = 0
    fallidos  = 0

    for keyword, srel in a_descargar.items():
        nombre_archivo = srel.rstrip("/").split("/")[-1]
        destino        = dest_base / nombre_archivo
        size           = _bajar_archivo(sess, host, srel, destino)
        if size >= 0:
            resultado["docs_rutas"][keyword] = destino
            exitosos += 1
        else:
            fallidos += 1

    resultado["stats"] = {
        "total":     len(a_descargar),
        "exitosos":  exitosos,
        "fallidos":  fallidos,
    }
    return resultado

def descargar_visita2_selectiva(
    url: str,
    dest_base: Path,
) -> dict:
    """
    Procesa la carpeta 02_VISITA_2_DIAGNOSTICO de forma selectiva:

      1. Localiza la subcarpeta 02_VISITA_2_DIAGNOSTICO.
      2. Lista todos sus archivos.
      3. Descarga únicamente:
           - ACTA_VISITA_2
           - DIAGNOSTICO
           - PLAN_NEGOCIO

    Retorna:
      encontrada  → bool
      docs_rutas  → rutas locales de los documentos encontrados
      stats       → estadísticas de descarga
    """

    from app.core.normalizacion import normalizar as _norm
    from app.core.visita2 import PALABRAS_CLAVE_VISITA2

    resultado = {
        "encontrada": False,
        "docs_rutas": {k: None for k in PALABRAS_CLAVE_VISITA2},
        "stats": {"total": 0, "exitosos": 0, "fallidos": 0},
    }

    sess = _nueva_sesion()
    host, web, rel = resolver(sess, url)

    rel_visita2 = _buscar_carpeta_visita2_en_sharepoint(
        sess, host, web, rel,
    )

    if rel_visita2 is None:
        return resultado

    resultado["encontrada"] = True

    todos_archivos = _listar_arbol(
        sess, host, web, rel_visita2,
    )

    a_descargar = {}

    for srel in todos_archivos:
        nombre = srel.rstrip("/").split("/")[-1]
        stem = _norm(Path(nombre).stem)

        for keyword, palabras in PALABRAS_CLAVE_VISITA2.items():
            if keyword not in a_descargar and any(
                palabra in stem
                for palabra in palabras
            ):
                a_descargar[keyword] = srel
                break

    if not a_descargar:
        return resultado

    dest_base.mkdir(parents=True, exist_ok=True)

    exitosos = 0
    fallidos = 0

    for keyword, srel in a_descargar.items():
        nombre_archivo = srel.rstrip("/").split("/")[-1]
        destino = dest_base / nombre_archivo

        size = _bajar_archivo(
            sess, host, srel, destino,
        )

        if size >= 0:
            resultado["docs_rutas"][keyword] = destino
            exitosos += 1
        else:
            fallidos += 1

    resultado["stats"] = {
        "total": len(a_descargar),
        "exitosos": exitosos,
        "fallidos": fallidos,
    }

    return resultado

def _buscar_carpeta_03_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """Localiza la subcarpeta que empieza por '03' (capacitación/formación)."""
    return _buscar_subcarpeta(
        sess, host, web, rel_raiz,
        lambda n: n.startswith("03"),
    )


def _buscar_carpeta_04_en_sharepoint(
    sess: requests.Session, host: str, web: str, rel_raiz: str
) -> str | None:
    """Localiza la subcarpeta que empieza por '04' (capitalización)."""
    return _buscar_subcarpeta(
        sess, host, web, rel_raiz,
        lambda n: n.startswith("04"),
    )


def _es_carpeta_modulo(nombre: str) -> bool:
    """True si el nombre normalizado corresponde a una carpeta de módulo."""
    return bool(re.search(r'mod(?:ulo)?[_\s\-]*\d', nombre))


def _es_txrx(stem_norm: str) -> bool:
    """
    True si el stem normalizado corresponde a un archivo de lista de asistencia.
    Patrones reales observados: T3_R5_..., T1_R4_..., TX_RX_..., RX_TX_...
    Regla: empieza por t{dígito(s)}_r{dígito(s)} o r{dígito(s)}_t{dígito(s)},
    o contiene literalmente 'tx' y 'rx'.
    """
    return bool(
        re.match(r't\d+[_\-]r\d+', stem_norm)
        or re.match(r'r\d+[_\-]t\d+', stem_norm)
        or ("tx" in stem_norm and "rx" in stem_norm)
    )


def descargar_carpeta_03(url: str, dest_base: Path, id_unico: str) -> dict:
    """
    Procesa la carpeta 03_* (capacitación) de forma selectiva:

    Raíz:
      - Descarga archivos con "encuesta" / "encuestas" en el nombre.
      - Descarga archivo con "grupal" en el nombre.
      - Descarga archivo con ID_<id_unico> en el nombre.

    Subcarpetas de módulo (nombre contiene "modulo" o "mod" + dígito):
      - Descarga el archivo TX_RX (o RX_TX).
      - Cuenta los demás archivos como evidencia fotográfica (sin descargarlos).

    Retorna:
      encontrada  → bool
      root        → {encuestas: [Path], grupal: Path|None, individual: Path|None}
      modulos     → [{nombre, txrx: Path|None, conteo_evidencia: int}]
    """
    resultado: dict = {
        "encontrada": False,
        "root": {"encuestas": [], "grupal": None, "individual": None},
        "modulos": [],
    }

    sess           = _nueva_sesion()
    host, web, rel = resolver(sess, url)

    rel_03 = _buscar_carpeta_03_en_sharepoint(sess, host, web, rel)
    if rel_03 is None:
        logger.warning("No se encontró carpeta 03 en: %s", rel)
        return resultado

    resultado["encontrada"] = True
    logger.info("Carpeta 03 encontrada: %s", rel_03)

    subcarpetas, archivos_raiz = _listar_carpeta(sess, host, web, rel_03)
    dest_base.mkdir(parents=True, exist_ok=True)

    # Dígitos del id_unico para detectar el informe individual
    id_digits = "".join(c for c in id_unico if c.isdigit())

    # ── Archivos en la raíz de 03 ─────────────────────────────────────────────
    for srel in archivos_raiz:
        nombre    = srel.rstrip("/").split("/")[-1]
        stem_norm = normalizar(Path(nombre).stem)

        es_encuesta   = "encuesta" in stem_norm or "encuestas" in stem_norm
        es_grupal     = "grupal" in stem_norm
        es_individual = (
            (id_digits and id_digits in stem_norm)
            or (id_unico and normalizar(id_unico) in stem_norm)
            or re.search(r'\bid[_\s]?\d', stem_norm) is not None
            or "formacion" in stem_norm
            or "cumplimiento" in stem_norm
        )

        if not (es_encuesta or es_grupal or es_individual):
            continue

        destino = dest_base / nombre
        size    = _bajar_archivo(sess, host, srel, destino)
        if size < 0:
            continue

        if es_encuesta:
            resultado["root"]["encuestas"].append(destino)
        if es_grupal and resultado["root"]["grupal"] is None:
            resultado["root"]["grupal"] = destino
        if es_individual and resultado["root"]["individual"] is None:
            resultado["root"]["individual"] = destino

    # ── Subcarpetas de módulo ─────────────────────────────────────────────────
    for rel_sub in subcarpetas:
        nombre_sub = rel_sub.rstrip("/").split("/")[-1]
        if not _es_carpeta_modulo(normalizar(nombre_sub)):
            logger.debug("Subcarpeta ignorada (no es módulo): %s", nombre_sub)
            continue

        _, archivos_mod = _listar_carpeta(sess, host, web, rel_sub)

        dest_mod = dest_base / nombre_sub
        dest_mod.mkdir(parents=True, exist_ok=True)

        txrx_path      = None
        conteo_evidencia = 0

        for srel_arch in archivos_mod:
            nombre_arch = srel_arch.rstrip("/").split("/")[-1]
            stem_norm   = normalizar(Path(nombre_arch).stem)

            if _es_txrx(stem_norm):
                destino = dest_mod / nombre_arch
                size    = _bajar_archivo(sess, host, srel_arch, destino)
                if size >= 0:
                    txrx_path = destino
            else:
                conteo_evidencia += 1

        resultado["modulos"].append({
            "nombre":           nombre_sub,
            "txrx":             txrx_path,
            "conteo_evidencia": conteo_evidencia,
        })
        logger.info(
            "  Módulo '%s': txrx=%s, evidencia=%d",
            nombre_sub, txrx_path is not None, conteo_evidencia,
        )

    return resultado


def descargar_carpeta_04(url: str, dest_base: Path, id_unico: str) -> dict:
    """
    Procesa la carpeta 04_* (capitalización):

    1. Localiza la carpeta 04 en la raíz del beneficiario.
    2. Dentro de 04, busca la primera subcarpeta cuyo nombre empieza por '01'.
    3. Descarga todos los archivos de esa subcarpeta 01_* en dest_base.

    Retorna:
      encontrada        → bool — se encontró la carpeta 04
      sub01_encontrada  → bool — se encontró la subcarpeta 01 dentro de 04
      archivos          → list[Path] — archivos descargados
      conteo            → int — total de archivos en la subcarpeta (descargados)
    """
    resultado: dict = {
        "encontrada":       False,
        "sub01_encontrada": False,
        "archivos":         [],
        "conteo":           0,
    }

    sess           = _nueva_sesion()
    host, web, rel = resolver(sess, url)

    rel_04 = _buscar_carpeta_04_en_sharepoint(sess, host, web, rel)
    if rel_04 is None:
        logger.warning("No se encontró carpeta 04 en: %s", rel)
        return resultado

    resultado["encontrada"] = True
    logger.info("Carpeta 04 encontrada: %s", rel_04)

    # Buscar subcarpeta que empieza con "01" dentro de 04
    try:
        subcarpetas_04, _ = _listar_carpeta(sess, host, web, rel_04)
    except Exception as exc:
        logger.warning("Error listando 04 '%s': %s", rel_04, exc)
        return resultado

    rel_sub01 = None
    for rel_sub in subcarpetas_04:
        nombre_sub = normalizar(rel_sub.rstrip("/").split("/")[-1])
        if nombre_sub.startswith("01"):
            rel_sub01 = rel_sub
            break

    if rel_sub01 is None:
        logger.warning("No se encontró subcarpeta 01_* dentro de 04: %s", rel_04)
        return resultado

    resultado["sub01_encontrada"] = True
    nombre_sub01 = urllib.parse.unquote(rel_sub01.rstrip("/").split("/")[-1])
    logger.info("Subcarpeta 01 encontrada dentro de 04: %s", nombre_sub01)

    # Descargar todos los archivos de la subcarpeta 01_*
    try:
        _, archivos_sub01 = _listar_carpeta(sess, host, web, rel_sub01)
    except Exception as exc:
        logger.warning("Error listando subcarpeta '%s': %s", rel_sub01, exc)
        return resultado

    dest_sub = dest_base / nombre_sub01
    dest_sub.mkdir(parents=True, exist_ok=True)

    descargados: list[Path] = []
    for srel in archivos_sub01:
        nombre_arch = urllib.parse.unquote(srel.rstrip("/").split("/")[-1])
        destino     = dest_sub / nombre_arch
        size        = _bajar_archivo(sess, host, srel, destino)
        if size >= 0:
            descargados.append(destino)
            logger.info("    04/01 descargado: %s (%d bytes)", nombre_arch, size)
        else:
            logger.warning("    04/01 fallo descarga: %s", nombre_arch)

    resultado["archivos"] = descargados
    resultado["conteo"]   = len(descargados)
    logger.info(
        "  04/%s: %d/%d archivos descargados",
        nombre_sub01, len(descargados), len(archivos_sub01),
    )
    return resultado


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
