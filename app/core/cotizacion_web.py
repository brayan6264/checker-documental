"""
Búsqueda web de precios de referencia para cotizaciones seleccionadas.

Flujo por producto:
  1. buscar_links_openai()   — gpt-4o-mini + web_search_preview → 9 URLs candidatas.
  2. tomar_screenshots()     — por ronda:
       a. Playwright visita todas las URLs pendientes → guarda capturas en disco.
       b. Se arma un PDF con las capturas de la ronda.
       c. gpt-4o-mini analiza el PDF completo → identifica qué páginas son válidas
          y extrae el precio de cada una.
       d. Si hay ≥ 3 válidas → listo. Si no → buscar N URLs más y repetir.
  3. generar_excel_cotizaciones() — Excel con capturas válidas + análisis de precio.
"""

import base64
import hashlib
import io
import json
import logging
import re
from datetime import date
from pathlib import Path
from typing import NamedTuple
from urllib.parse import urlparse

_RE_HOJA_INVALIDOS = re.compile(r'[\\/?*\[\]:]')

logger = logging.getLogger(__name__)

# Un solo modelo para todo: búsqueda web + análisis de PDF
_MODELO = "gpt-4o-mini"

# URLs que Playwright visita en la primera ronda
_URLS_INICIALES = 9

# Máximo de rondas de búsqueda (inicial + reemplazos)
_MAX_RONDAS = 3


# ── Tipos ─────────────────────────────────────────────────────────────────────

class LinkProducto(NamedTuple):
    url:           str
    precio_texto:  str
    precio_numero: int | None


class ResultadoScreenshot(NamedTuple):
    ruta:          Path | None
    precio_texto:  str | None
    precio_numero: int | None
    sin_stock:     bool = False


class ProductoLinks(NamedTuple):
    item:        str
    descripcion: str | None
    links:       list[LinkProducto]


# Dominios bloqueados en código — el modelo los ignora a veces en el prompt
_DOMINIOS_BLOQUEADOS = {
    "mercadolibre.com.co",
    "mercadolibre.com",
    "mercadolibre.com.mx",
    "mercadolibre.com.ar",
    "mercadolibre.cl",
    "mercadolibre.com.pe",
    "mercadolibre.com.ve",
    "meli.com.co",
}


def _url_bloqueada(url: str) -> bool:
    """Retorna True si la URL pertenece a un dominio bloqueado."""
    dominio = urlparse(url).netloc.lower().removeprefix("www.")
    return any(dominio == b or dominio.endswith("." + b) for b in _DOMINIOS_BLOQUEADOS)


_RE_PRECIO_COP = re.compile(
    r'(?:COP\s*|cop\s*|\$\s*)?(\d{1,3}(?:[.,]\d{3})+)(?:\s*(?:COP|cop))?'
)

_PRECIO_MIN_COP = 5_000
_PRECIO_MAX_COP = 50_000_000


# ── Helpers generales ─────────────────────────────────────────────────────────

def _normalizar_url(url: str) -> str:
    if "tienda.exito.com" in url:
        url = url.replace("tienda.exito.com", "www.exito.com")
    return url


def _es_precio_razonable(num: int) -> bool:
    return _PRECIO_MIN_COP <= num <= _PRECIO_MAX_COP


def _parsear_precio_cop(texto: str) -> int | None:
    limpio = re.sub(r'[^\d.,]', '', texto.strip())
    if re.fullmatch(r'\d{1,3}(\.\d{3})+', limpio):
        return int(limpio.replace('.', ''))
    if re.fullmatch(r'\d{1,3}(,\d{3})+', limpio):
        return int(limpio.replace(',', ''))
    if re.fullmatch(r'\d+', limpio) and len(limpio) >= 4:
        return int(limpio)
    return None


def _formatear_cop(valor: int) -> str:
    return f"$ {valor:,}".replace(",", ".")


def _extraer_json(text: str) -> dict | list | None:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    for pat in [
        r"```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```",
        r"(\{.*\}|\[.*\])",
    ]:
        m = re.search(pat, text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass
    return None


# ── PASO 1 — Búsqueda de URLs con gpt-4o-mini + web_search_preview ───────────

def buscar_links_openai(
    seleccionadas: list,
    api_key: str,
) -> tuple[bool, list[ProductoLinks], str]:
    """
    Llama a gpt-4o-mini con web_search_preview para obtener 9 URLs candidatas
    por producto.  Retorna (ok, productos, resumen).
    """
    try:
        from openai import OpenAI
    except ImportError:
        return False, [], "FALTA — openai no instalado"

    if not api_key:
        return False, [], "FALTA — OPENAI_API_KEY no configurada"

    client  = OpenAI(api_key=api_key)
    hoy_str = date.today().strftime("%d de %B de %Y")

    productos_out: list[ProductoLinks] = []
    errores: list[str] = []

    for sel in seleccionadas:
        desc = sel.descripcion or sel.item

        prompt = f"""Hoy es {hoy_str}. Necesito comprar este producto en Colombia HOY.

PRODUCTO: {desc}

Busca en internet 9 páginas de tiendas colombianas donde se venda este producto.
Necesito variedad: mezcla grandes superficies, tiendas especializadas, distribuidores, etc.

REGLAS obligatorias para cada URL:
1. URL directa a la ficha del producto (no categorías, no buscadores, no home).
2. Cada URL de un DOMINIO DISTINTO.
3. Precio visible en pesos colombianos (COP).
4. Producto disponible (no agotado).
5. PROHIBIDO ESTRICTAMENTE: MercadoLibre y TODAS sus variantes (mercadolibre.com.co, mercadolibre.com, meli.com.co, mercadolibre.com.mx, etc.). Cualquier resultado de MercadoLibre será descartado automáticamente. No incluyas ningún link de MercadoLibre bajo ninguna circunstancia.

Tiendas sugeridas (no exclusivas): alkosto.com, homecenter.com.co, exito.com,
falabella.com.co, ktronix.com, panamericana.com.co, jumbo.co, makro.com.co,
y cualquier tienda especializada del sector.

Responde ÚNICAMENTE con este JSON (sin texto extra, sin ```):
{{
  "links": [
    {{"url": "https://...", "precio": "$ 89.900", "precio_numero": 89900, "tienda": "tienda.com"}},
    ... (9 objetos en total)
  ]
}}
precio_numero: entero COP sin puntos ni símbolos."""

        try:
            response = client.responses.create(
                model=_MODELO,
                tools=[{"type": "web_search_preview"}],
                input=[{"role": "user", "content": prompt}],
            )
            text = response.output_text
        except Exception as exc:
            logger.error("OpenAI búsqueda error '%s': %s", sel.item, exc)
            errores.append(str(exc))
            continue

        data = _extraer_json(text)
        if not data:
            logger.error("JSON inválido para '%s': %s", sel.item, text[:400])
            errores.append(f"JSON inválido para {sel.item}")
            continue

        links: list[LinkProducto] = []
        dominios_vistos: set[str] = set()
        for lk in (data.get("links") or [])[:_URLS_INICIALES]:
            url = str(lk.get("url", "")).strip()
            if not url:
                continue
            if _url_bloqueada(url):
                logger.warning("  URL bloqueada (dominio excluido): %s", url)
                continue
            dominio = urlparse(url).netloc.lower().removeprefix("www.")
            if dominio in dominios_vistos:
                continue
            dominios_vistos.add(dominio)

            precio_num = lk.get("precio_numero")
            if isinstance(precio_num, float):
                precio_num = int(precio_num)
            if precio_num and not _es_precio_razonable(precio_num):
                precio_num = None

            links.append(LinkProducto(
                url=url,
                precio_texto=str(lk.get("precio", "N/A")),
                precio_numero=precio_num,
            ))

        if links:
            productos_out.append(ProductoLinks(item=sel.item, descripcion=desc, links=links))
            logger.info("  '%s': %d URL(s) candidatas", sel.item, len(links))
        else:
            logger.warning("  '%s': sin enlaces en la respuesta", sel.item)
            errores.append(f"Sin enlaces para {sel.item}")

    if not productos_out:
        return False, [], "gpt-4o-mini no retornó enlaces para ningún producto"

    total   = sum(len(p.links) for p in productos_out)
    resumen = f"OK — {len(productos_out)} producto(s), {total} URL(s) candidatas"
    if errores:
        resumen += f" | errores: {'; '.join(errores)}"
    return True, productos_out, resumen


# ── PASO 2a — Playwright: captura + precio DOM ────────────────────────────────

_JS_JSONLD = """
() => {
    const buscar = obj => {
        if (!obj) return null;
        if (Array.isArray(obj)) {
            for (const item of obj) { const r = buscar(item); if (r) return r; }
            return null;
        }
        if (obj['@graph']) { const r = buscar(obj['@graph']); if (r) return r; }
        const t = obj['@type'];
        const esProducto = t === 'Product' || (Array.isArray(t) && t.includes('Product'));
        if (esProducto) {
            let offers = obj.offers;
            if (Array.isArray(offers)) offers = offers[0];
            if (offers) {
                return {
                    price:        offers.price        ?? offers.lowPrice ?? null,
                    currency:     offers.priceCurrency ?? null,
                };
            }
        }
        return null;
    };
    for (const s of document.querySelectorAll('script[type="application/ld+json"]')) {
        try { const r = buscar(JSON.parse(s.textContent)); if (r) return r; } catch(e) {}
    }
    return null;
}
"""

_JS_PRECIO_PRINCIPAL = """
() => {
    const EXCLUIR_CARD = '[class*="productCard"],[class*="ProductCard"],[class*="shelf-item"],' +
                         '[class*="ShelfItem"],[class*="carousel"],[class*="Carousel"],' +
                         '[class*="related"],[class*="Related"],[class*="recommendation"],' +
                         '[class*="Recommendation"],[class*="priceOfferButton"]';
    const EXCLUIR_TACHADO = /dashed|Dashed|original|Original|old-price|before|tachado|list-price|listPrice|regularPrice/;
    const SELS = [
        '[itemprop="price"]','[data-price]',
        '[class*="sellingPrice"]:not([class*="list"])',
        '[class*="SellingPrice"]:not([class*="list"])',
        '[class*="container__price"]','[class*="ProductPrice_container__price"]',
        '[class*="selling-price"]','[class*="price-final"]',
        '[class*="price--sale"]','[class*="price--selling"]',
        '[class*="current-price"]','[class*="precio-actual"]',
        '[class*="effectivePrice"]','[class*="price-box"]',
        '[class*="product-prices__effective"]',
    ];
    const extraer = el => { const c = el.getAttribute('content'); return (c && /\\d/.test(c)) ? c : el.textContent.trim(); };
    const valido  = el => { if (el.closest(EXCLUIR_CARD)) return false; return !EXCLUIR_TACHADO.test(el.className||''); };
    const h1 = document.querySelector('h1');
    if (h1) {
        let cont = h1.parentElement;
        for (let i=0; i<12; i++) {
            if (!cont || cont===document.body) break;
            for (const s of SELS) {
                for (const el of Array.from(cont.querySelectorAll(s)).filter(valido)) {
                    const t = extraer(el); if (t && /\\d{3}/.test(t)) return t;
                }
            }
            cont = cont.parentElement;
        }
    }
    for (const s of SELS) {
        const el = Array.from(document.querySelectorAll(s)).find(valido);
        if (el) { const t = extraer(el); if (t && /\\d{3}/.test(t)) return t; }
    }
    return null;
}
"""


def _screenshot_una_url(
    page,
    url: str,
    dir_out: Path,
) -> tuple[Path | None, str | None, int | None]:
    """
    Navega a la URL, extrae precio con DOM y toma screenshot.
    Retorna (ruta_png, precio_texto, precio_numero).
    """
    fname = hashlib.md5(url.encode()).hexdigest()[:12] + ".png"
    ruta  = dir_out / fname
    try:
        page.goto(url, timeout=30_000, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)

        # Precio DOM (respaldo por si el PDF no lo lee bien)
        precio_txt, precio_num = _extraer_precio_dom(page)
        page.screenshot(path=str(ruta), full_page=False)
        return ruta, precio_txt, precio_num

    except Exception as exc:
        logger.warning("  Screenshot error %s: %s", url, exc)
        return None, None, None


def _extraer_precio_dom(page) -> tuple[str | None, int | None]:
    # 1. JSON-LD
    try:
        info = page.evaluate(_JS_JSONLD)
        if info:
            currency = (info.get("currency") or "").upper()
            if currency in ("COP", ""):
                raw = info.get("price")
                if raw is not None:
                    s = str(raw).strip()
                    num = _parsear_precio_cop(s)
                    if not num:
                        try:
                            num = int(float(s))
                        except (ValueError, OverflowError):
                            pass
                    if num and _es_precio_razonable(num):
                        return _formatear_cop(num), num
    except Exception:
        pass

    # 2. JS cerca del H1
    try:
        t = page.evaluate(_JS_PRECIO_PRINCIPAL)
        if t:
            num = _parsear_precio_cop(t)
            if not num:
                nums = [_parsear_precio_cop(m) for m in _RE_PRECIO_COP.findall(t)]
                nums = [n for n in nums if n and _es_precio_razonable(n)]
                num  = max(nums) if nums else None
            if num and _es_precio_razonable(num):
                return _formatear_cop(num), num
    except Exception:
        pass

    # 3. Escaneo de texto
    try:
        texto = page.evaluate("document.body.innerText") or ""
        candidatos = [
            _parsear_precio_cop(m.group(0))
            for m in _RE_PRECIO_COP.finditer(texto)
        ]
        candidatos = [n for n in candidatos if n and _es_precio_razonable(n)]
        if candidatos:
            from collections import Counter
            return _formatear_cop(Counter(candidatos).most_common(1)[0][0]), Counter(candidatos).most_common(1)[0][0]
    except Exception:
        pass

    return None, None


# ── PASO 2b — Armar PDF con capturas de la ronda ─────────────────────────────

def _crear_pdf_capturas(
    capturas: list[tuple[str, Path | None]],   # [(url, ruta_png), ...]
) -> bytes | None:
    """
    Crea un PDF (una captura por página) usando Pillow.
    Retorna los bytes del PDF o None si no hay imágenes.
    """
    try:
        from PIL import Image
    except ImportError:
        logger.error("Pillow no instalado (pip install Pillow)")
        return None

    pages: list = []
    for _url, ruta in capturas:
        if ruta and ruta.exists():
            try:
                img = Image.open(str(ruta)).convert("RGB")
                pages.append(img)
            except Exception as exc:
                logger.warning("  PDF: no se pudo cargar imagen %s: %s", ruta, exc)

    if not pages:
        return None

    buf = io.BytesIO()
    pages[0].save(buf, format="PDF", save_all=True, append_images=pages[1:])
    return buf.getvalue()


# ── PASO 2c — GPT analiza el PDF y decide cuáles capturas sirven ─────────────

def _analizar_pdf_con_gpt(
    pdf_bytes: bytes,
    urls: list[str],
    descripcion: str,
    api_key: str,
) -> list[dict]:
    """
    Envía el PDF a gpt-4o-mini vía Responses API.
    Cada página del PDF = una captura de pantalla (en el orden de `urls`).

    Retorna lista de dicts para las páginas válidas:
      [{url, precio_texto, precio_numero}, ...]
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    pdf_b64   = base64.b64encode(pdf_bytes).decode()
    lista_pag = "\n".join(f"  Página {i+1}: {u}" for i, u in enumerate(urls))

    prompt = f"""Este PDF contiene {len(urls)} capturas de pantalla de páginas web de tiendas colombianas.

Producto buscado: "{descripcion}"

Correspondencia páginas → URLs:
{lista_pag}

Para CADA página evalúa:
1. ¿Muestra una ficha de producto que corresponde a "{descripcion}" (o equivalente)?
2. ¿Cuál es el precio principal de venta en pesos colombianos (COP)?

En Colombia los precios usan PUNTOS como separadores de miles: $ 1.234.567
Busca el precio prominente junto al nombre del producto, no precios tachados ni de recomendados.

Devuelve SOLO este JSON (sin texto adicional):
{{
  "validos": [
    {{
      "pagina": 1,
      "url": "url exacta copiada de la lista de arriba",
      "precio_texto": "$ 1.234.567",
      "precio_numero": 1234567
    }}
  ],
  "hay_suficientes": true
}}

Reglas:
- Incluir en "validos" solo páginas que SÍ muestran el producto correcto.
- Si el precio no es legible en la imagen, igual incluye la página en "validos" con precio_texto y precio_numero en null.
- "hay_suficientes": true si hay 3 o más páginas válidas.
- precio_numero: entero sin puntos (1234567, no "1.234.567")."""

    try:
        response = client.responses.create(
            model=_MODELO,
            input=[{
                "role": "user",
                "content": [
                    {
                        "type": "input_file",
                        "filename": "capturas.pdf",
                        "file_data": f"data:application/pdf;base64,{pdf_b64}",
                    },
                    {
                        "type": "input_text",
                        "text": prompt,
                    },
                ],
            }],
        )
        texto = response.output_text
        logger.debug("GPT-PDF respuesta: %s", texto[:500])
        datos = _extraer_json(texto)

        if datos and isinstance(datos, dict):
            resultado = []
            for v in (datos.get("validos") or []):
                pagina = int(v.get("pagina", 0)) - 1   # 0-based
                if not (0 <= pagina < len(urls)):
                    continue
                # Usar la URL de la lista (más fiable que lo que devuelva el modelo)
                url_real   = urls[pagina]
                precio_num = v.get("precio_numero")
                if isinstance(precio_num, float):
                    precio_num = int(precio_num)
                if precio_num and not _es_precio_razonable(precio_num):
                    logger.warning("GPT-PDF: precio fuera de rango (%s) descartado", precio_num)
                    precio_num = None
                resultado.append({
                    "url":           url_real,
                    "precio_texto":  v.get("precio_texto"),
                    "precio_numero": precio_num,
                })
            logger.info(
                "GPT-PDF: %d/%d páginas válidas (hay_suficientes=%s)",
                len(resultado), len(urls), datos.get("hay_suficientes"),
            )
            return resultado

        logger.warning("GPT-PDF: respuesta inesperada: %s", texto[:300])

    except Exception as exc:
        logger.error("Error analizando PDF con GPT: %s", exc)

    return []


# ── PASO 2d — Búsqueda de URLs adicionales ───────────────────────────────────

def _buscar_links_reemplazo(
    descripcion: str,
    n: int,
    excluir: set[str],
    api_key: str,
) -> list[str]:
    """
    Pide a gpt-4o-mini + web_search_preview exactamente N URLs nuevas,
    excluyendo los dominios ya intentados.
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    excluir_dominios = sorted({
        urlparse(u).netloc.lower().removeprefix("www.")
        for u in excluir
    })

    prompt = (
        f"Necesito {n} URL(s) alternativa(s) de tiendas colombianas para:\n"
        f"PRODUCTO: {descripcion}\n\n"
        f"Dominios ya intentados (NO repetir): {', '.join(excluir_dominios) or 'ninguno'}\n\n"
        f"REGLAS:\n"
        f"1. URL directa a la ficha del producto.\n"
        f"2. Dominio DIFERENTE a los excluidos.\n"
        f"3. Precio en COP, producto en stock.\n"
        f"4. PROHIBIDO ESTRICTAMENTE: MercadoLibre y TODAS sus variantes "
        f"(mercadolibre.com.co, mercadolibre.com, meli.com.co, mercadolibre.com.mx, etc.). "
        f"Cualquier resultado de MercadoLibre será descartado automáticamente. "
        f"No incluyas ningún link de MercadoLibre bajo ninguna circunstancia.\n\n"
        f"Responde SOLO con JSON:\n"
        f'{{"links":[{{"url":"https://...","precio":"$ 89.900","precio_numero":89900}}]}}'
    )

    try:
        response = client.responses.create(
            model=_MODELO,
            tools=[{"type": "web_search_preview"}],
            input=[{"role": "user", "content": prompt}],
        )
        data = _extraer_json(response.output_text)
        if data and isinstance(data, dict):
            urls = []
            for lk in (data.get("links") or [])[:n]:
                url = str(lk.get("url", "")).strip()
                if not url or url in excluir:
                    continue
                if _url_bloqueada(url):
                    logger.warning("  Reemplazo bloqueado (dominio excluido): %s", url)
                    continue
                urls.append(url)
            return urls
    except Exception as exc:
        logger.error("Búsqueda reemplazo error: %s", exc)

    return []


# ── PASO 2 — Orquestador principal ───────────────────────────────────────────

def tomar_screenshots(
    productos: list[ProductoLinks],
    dir_out: Path,
    api_key: str = "",
) -> tuple[list[ProductoLinks], dict[str, ResultadoScreenshot]]:
    """
    Para cada producto:
      Ronda 1+:
        a. Playwright visita todas las URLs pendientes → capturas en disco.
        b. Se ensambla un PDF con las capturas de esta ronda.
        c. gpt-4o-mini analiza el PDF → identifica válidas + precios.
        d. Si < 3 válidas → buscar N URLs más → nueva ronda.
      Máx _MAX_RONDAS rondas en total.

    Retorna (productos_actualizados, {url: ResultadoScreenshot}).
    """
    from playwright.sync_api import sync_playwright

    dir_out.mkdir(parents=True, exist_ok=True)
    dir_pdfs = dir_out / "_pdfs_analisis"
    dir_pdfs.mkdir(exist_ok=True)

    res: dict[str, ResultadoScreenshot]  = {}
    productos_actualizados: list[ProductoLinks] = []

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled"],
            )
            ctx = browser.new_context(
                viewport={"width": 1280, "height": 900},
                locale="es-CO",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            ctx.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            page = ctx.new_page()
            page.route(
                "**/*",
                lambda route: route.abort()
                if route.request.resource_type in ("font", "media")
                else route.continue_(),
            )

            for producto in productos:
                desc            = producto.descripcion or producto.item
                urls_pendientes = [_normalizar_url(lk.url) for lk in producto.links if lk.url]
                urls_intentadas: set[str] = set(urls_pendientes)
                links_validos: list[LinkProducto] = []
                # precio DOM por URL (guardado durante la visita para usarlo como respaldo)
                precio_dom: dict[str, tuple[str | None, int | None]] = {}
                rondas_ok = 0

                for ronda in range(1, _MAX_RONDAS + 1):
                    if not urls_pendientes:
                        break

                    logger.info(
                        "  [%s] Ronda %d/%d — visitando %d URL(s) con Playwright",
                        producto.item, ronda, _MAX_RONDAS, len(urls_pendientes),
                    )

                    # ── a. Playwright: screenshots ────────────────────────────
                    capturas_ronda: list[tuple[str, Path | None]] = []
                    for url in urls_pendientes:
                        ruta, p_txt, p_num = _screenshot_una_url(page, url, dir_out)
                        capturas_ronda.append((url, ruta))
                        precio_dom[url] = (p_txt, p_num)
                        if ruta:
                            res[url] = ResultadoScreenshot(ruta, p_txt, p_num, False)
                        else:
                            res[url] = ResultadoScreenshot(None, None, None, False)

                    # ── b. Armar PDF ──────────────────────────────────────────
                    pdf_bytes = _crear_pdf_capturas(capturas_ronda)

                    if pdf_bytes:
                        pdf_path = dir_pdfs / (
                            hashlib.md5(producto.item.encode()).hexdigest()[:8]
                            + f"_r{ronda}.pdf"
                        )
                        pdf_path.write_bytes(pdf_bytes)
                        logger.info("  [%s] PDF generado: %s", producto.item, pdf_path.name)

                    # ── c. GPT analiza el PDF ─────────────────────────────────
                    urls_ronda = [url for url, _ in capturas_ronda]

                    if api_key and pdf_bytes:
                        validos_gpt = _analizar_pdf_con_gpt(pdf_bytes, urls_ronda, desc, api_key)
                    else:
                        # Sin IA: aceptar todas las que tienen captura
                        validos_gpt = [
                            {"url": url, "precio_texto": None, "precio_numero": None}
                            for url, ruta in capturas_ronda
                            if ruta
                        ]

                    # Registrar resultados de esta ronda
                    urls_invalidas = set(urls_ronda)
                    for v in validos_gpt:
                        url        = v["url"]
                        ruta       = next((r for u, r in capturas_ronda if u == url), None)
                        p_txt_dom, p_num_dom = precio_dom.get(url, (None, None))

                        # GPT tiene prioridad de precio (lee la imagen); DOM como respaldo
                        precio_num = v["precio_numero"] or p_num_dom
                        precio_txt = v["precio_texto"]  or p_txt_dom
                        if precio_num and not _es_precio_razonable(precio_num):
                            precio_num = None
                            precio_txt = None

                        res[url] = ResultadoScreenshot(
                            ruta=ruta,
                            precio_texto=precio_txt,
                            precio_numero=precio_num,
                            sin_stock=False,
                        )
                        links_validos.append(LinkProducto(
                            url=url,
                            precio_texto=precio_txt or "N/A",
                            precio_numero=precio_num,
                        ))
                        urls_invalidas.discard(url)
                        logger.info(
                            "  [%s] ✓ válido (%d/3): %s → %s",
                            producto.item, len(links_validos), url,
                            precio_txt or "sin precio",
                        )

                    for url in urls_invalidas:
                        logger.warning("  [%s] ✗ inválido: %s", producto.item, url)

                    rondas_ok += 1

                    # ── d. ¿Tenemos suficientes? ──────────────────────────────
                    needed = 3 - len(links_validos)
                    if needed <= 0:
                        logger.info("  [%s] 3/3 válidos — listo.", producto.item)
                        break

                    if ronda >= _MAX_RONDAS:
                        logger.warning(
                            "  [%s] Límite de rondas alcanzado con %d/3 válidos.",
                            producto.item, len(links_validos),
                        )
                        break

                    if not api_key:
                        break

                    logger.info("  [%s] Buscando %d URL(s) de reemplazo...", producto.item, needed)
                    nuevas = _buscar_links_reemplazo(desc, needed, urls_intentadas, api_key)
                    nuevas = [_normalizar_url(u) for u in nuevas if u not in urls_intentadas]
                    if not nuevas:
                        logger.warning("  [%s] Sin reemplazos disponibles.", producto.item)
                        break

                    urls_pendientes = nuevas
                    urls_intentadas.update(nuevas)

                logger.info(
                    "  [%s] Final: %d/3 válidos tras %d ronda(s).",
                    producto.item, len(links_validos), rondas_ok,
                )
                productos_actualizados.append(ProductoLinks(
                    item=producto.item,
                    descripcion=producto.descripcion,
                    links=links_validos[:3],
                ))

            browser.close()

    except Exception as exc:
        logger.error("Playwright error general: %s", exc)
        for producto in productos:
            if not any(p.item == producto.item for p in productos_actualizados):
                productos_actualizados.append(producto)

    return productos_actualizados, res


# ── Análisis de desviación de precio ─────────────────────────────────────────

_RE_UNIDADES = re.compile(
    r'[-–]\s*(\d+)\s*(?:unidades?|und\.?|uni\.?|uds?\.?|piezas?|pcs?\.?)',
    re.IGNORECASE,
)


def _extraer_unidades(nombre_item: str) -> int:
    """
    Extrae la cantidad de unidades del nombre del ítem.
    Ejemplos: "Hilos Macrame - 22 Unidades" → 22
              "Lámpara - 1 unidad"           → 1
              "Computador"                   → 1  (default)
    """
    m = _RE_UNIDADES.search(nombre_item or "")
    if m:
        try:
            return max(1, int(m.group(1)))
        except ValueError:
            pass
    return 1


def _mediana(valores: list[int]) -> int:
    """Mediana de enteros: más robusta que el promedio ante precios atípicos."""
    s = sorted(valores)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 == 1 else (s[mid - 1] + s[mid]) // 2


class AlertaDesviacion:
    __slots__ = ("mediana_unitaria_internet", "mediana_total_internet",
                 "precio_cotizacion", "unidades",
                 "desviacion_pct", "limite_pct", "supera_limite", "mensaje")

    def __init__(self, mediana_unitaria_internet, mediana_total_internet,
                 precio_cotizacion, unidades,
                 desviacion_pct, limite_pct, supera_limite, mensaje):
        self.mediana_unitaria_internet = mediana_unitaria_internet
        self.mediana_total_internet    = mediana_total_internet
        self.precio_cotizacion         = precio_cotizacion
        self.unidades                  = unidades
        self.desviacion_pct            = desviacion_pct
        self.limite_pct                = limite_pct
        self.supera_limite             = supera_limite
        self.mensaje                   = mensaje


def _analizar_desviacion(
    precio_cotizacion: int,
    precios_unitarios_internet: list[int],
    nombre_item: str = "",
) -> "AlertaDesviacion | None":
    """
    Compara el total cotizado con la mediana de internet ajustada por unidades.

    Se usa la mediana (no el promedio) para evitar que precios atípicos
    distorsionen la referencia de mercado.

    Lógica:
      - precio_cotizacion  = valor_total de la cotización (incluye N unidades)
      - precios_internet   = precios UNITARIOS encontrados en internet
      - unidades           = parseado del nombre del ítem ("22 Unidades")
      - mediana_total      = mediana_unitaria × unidades

    Reglas de tolerancia (sobre precio_cotizacion / unidades, es decir, por unidad):
      - precio_cotizacion/unidad < 500.000 COP → tolerancia 50 %
      - precio_cotizacion/unidad ≥ 500.000 COP → tolerancia 20 %

    Alerta cuando el precio cotizado supera la mediana de internet ajustada
    en más del % de tolerancia.
    """
    validos = [p for p in precios_unitarios_internet if p and p > 0]
    if not validos or not precio_cotizacion:
        return None

    unidades       = _extraer_unidades(nombre_item)
    mediana_unit   = _mediana(validos)
    mediana_total  = mediana_unit * unidades

    precio_unit_cotiz = precio_cotizacion / unidades
    limite            = 50.0 if precio_unit_cotiz < 500_000 else 20.0

    desviacion = (precio_cotizacion - mediana_total) / mediana_total * 100
    supera     = desviacion > limite

    mtotal_fmt = _formatear_cop(mediana_total)
    munit_fmt  = _formatear_cop(mediana_unit)
    pcotiz_fmt = _formatear_cop(precio_cotizacion)

    unid_str = f"{unidades} und." if unidades > 1 else "1 und."

    if supera:
        exceso  = desviacion - limite
        mensaje = (
            f"⚠ ALERTA: la cotización ({pcotiz_fmt} / {unid_str}) supera la mediana "
            f"de internet en {desviacion:.1f}% — "
            f"mediana internet: {munit_fmt}/und. × {unidades} = {mtotal_fmt} — "
            f"límite: {limite:.0f}%, exceso: {exceso:.1f}%"
        )
    else:
        mensaje = (
            f"✓ OK: cotización ({pcotiz_fmt} / {unid_str}) dentro del rango permitido — "
            f"mediana internet: {munit_fmt}/und. × {unidades} = {mtotal_fmt} — "
            f"desviación: {desviacion:+.1f}%, límite: {limite:.0f}%"
        )

    return AlertaDesviacion(
        mediana_unitaria_internet=mediana_unit,
        mediana_total_internet=mediana_total,
        precio_cotizacion=precio_cotizacion,
        unidades=unidades,
        desviacion_pct=desviacion,
        limite_pct=limite,
        supera_limite=supera,
        mensaje=mensaje,
    )


# ── PASO 3 — Excel ────────────────────────────────────────────────────────────

def generar_excel_cotizaciones(
    id_unico: str,
    productos: list[ProductoLinks],
    screenshots: dict[str, ResultadoScreenshot],
    ruta_salida: Path,
    seleccionadas: list | None = None,
) -> list[str]:
    """
    Genera el Excel de referencia de precios web (una hoja por producto + índice).
    Retorna lista de mensajes de alerta de desviación de precio disparados
    (uno por producto que supere el límite), para incluir en el checklist.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.drawing.image import Image as XLImage

    precio_cotizacion_por_item: dict[str, int] = {}
    if seleccionadas:
        for sel in seleccionadas:
            if sel.valor_total:
                precio_cotizacion_por_item[sel.item] = sel.valor_total

    alertas_desviacion: list[str] = []

    hoy = date.today().strftime("%d/%m/%Y")

    def _sheet_name(name: str) -> str:
        return _RE_HOJA_INVALIDOS.sub("_", name)[:31]

    _C_DARK  = "1F5C2E"
    _C_MED   = "2E7D32"
    _C_ODD   = "E8F5E9"
    _C_EVEN  = "F5F5F5"
    _C_WHITE = "FFFFFF"

    def _fill(hex6):
        return PatternFill("solid", start_color=hex6, end_color=hex6)

    def _font(bold=False, size=11, color="000000", name="Arial"):
        return Font(bold=bold, size=size, color=color, name=name)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Índice ───────────────────────────────────────────────────────────
    ws_i = wb.create_sheet("Índice")
    ws_i.merge_cells("A1:E1")
    ws_i["A1"].value     = f"Cotizaciones de referencia web — {id_unico}"
    ws_i["A1"].font      = _font(bold=True, size=14, color=_C_WHITE)
    ws_i["A1"].fill      = _fill(_C_DARK)
    ws_i["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_i.row_dimensions[1].height = 32

    ws_i["A2"].value = f"Generado: {hoy}  |  Modelo: {_MODELO}"
    ws_i["A2"].font  = _font(size=10, color="555555")
    ws_i.row_dimensions[2].height = 16

    for col, hdr in enumerate(["Producto", "Descripción", "Hoja"], 1):
        c = ws_i.cell(row=4, column=col, value=hdr)
        c.font  = _font(bold=True, color=_C_WHITE)
        c.fill  = _fill(_C_MED)
        c.alignment = Alignment(horizontal="center")
    ws_i.row_dimensions[4].height = 18

    for r, prod in enumerate(productos, 5):
        ws_i.cell(row=r, column=1, value=prod.item)
        ws_i.cell(row=r, column=2, value=prod.descripcion or "")
        ws_i.cell(row=r, column=3, value=_sheet_name(prod.item))
    ws_i.column_dimensions["A"].width = 35
    ws_i.column_dimensions["B"].width = 60
    ws_i.column_dimensions["C"].width = 35

    # ── Una hoja por producto ────────────────────────────────────────────
    IMG_ROWS = 24
    IMG_H_PX = 460
    IMG_W_PX = 900

    for prod in productos:
        ws = wb.create_sheet(title=_sheet_name(prod.item))

        ws.merge_cells("A1:E1")
        ws["A1"].value     = f"COTIZACIONES DE REFERENCIA — {prod.item.upper()}"
        ws["A1"].font      = _font(bold=True, size=13, color=_C_WHITE)
        ws["A1"].fill      = _fill(_C_DARK)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:C2")
        ws["A2"].value = f"Descripción: {prod.descripcion or 'N/A'}"
        ws["A2"].font  = _font(size=10)
        ws.merge_cells("D2:E2")
        ws["D2"].value     = f"Fecha: {hoy}  |  ID: {id_unico}"
        ws["D2"].font      = _font(size=10, color="555555")
        ws["D2"].alignment = Alignment(horizontal="right")
        ws.row_dimensions[2].height = 16

        for col, hdr in enumerate(["#", "URL", "Precio actual", "Captura de pantalla", "Fuente precio"], 1):
            c = ws.cell(row=3, column=col, value=hdr)
            c.font      = _font(bold=True, color=_C_WHITE)
            c.fill      = _fill(_C_MED)
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 20

        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 70
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["D"].width = 120
        ws.column_dimensions["E"].width = 22

        current_row = 4
        for i, link in enumerate(prod.links, 1):
            bg = _C_ODD if i % 2 == 1 else _C_EVEN

            lbl           = ws.cell(row=current_row, column=1, value=f"Op. {i}")
            lbl.font      = _font(bold=True)
            lbl.fill      = _fill(bg)
            lbl.alignment = Alignment(horizontal="center", vertical="center")

            url_cell           = ws.cell(row=current_row, column=2, value=link.url)
            url_cell.hyperlink = link.url
            url_cell.font      = Font(name="Arial", size=10, color="0563C1", underline="single")
            url_cell.fill      = _fill(bg)

            ss = screenshots.get(link.url, ResultadoScreenshot(None, None, None, False))
            if ss.precio_texto:
                precio_mostrar = ss.precio_texto
                precio_fuente  = "precio actual página"
            else:
                precio_mostrar = link.precio_texto or "No disponible"
                precio_fuente  = "precio referencia IA"

            precio_cell           = ws.cell(row=current_row, column=3, value=precio_mostrar)
            precio_cell.font      = _font(bold=True, size=11)
            precio_cell.fill      = _fill(bg)
            precio_cell.alignment = Alignment(horizontal="center", vertical="center")

            fuente_cell       = ws.cell(row=current_row, column=5, value=precio_fuente)
            fuente_cell.font  = _font(size=8, color="777777")
            fuente_cell.fill  = _fill(bg)
            fuente_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[current_row].height = 22
            current_row += 1

            img_row = current_row
            for r in range(IMG_ROWS):
                ws.row_dimensions[current_row + r].height = 20

            img_path = ss.ruta
            if img_path and img_path.exists():
                try:
                    xl_img        = XLImage(str(img_path))
                    xl_img.width  = IMG_W_PX
                    xl_img.height = IMG_H_PX
                    ws.add_image(xl_img, f"D{img_row}")
                except Exception as exc:
                    logger.warning("No se pudo insertar imagen en Excel: %s", exc)
                    ws.cell(row=img_row, column=4, value="[Captura no disponible]").font = _font(color="999999")
            else:
                ws.cell(row=img_row, column=4, value="[Captura no disponible]").font = _font(color="999999")

            current_row += IMG_ROWS + 1

        # ── Análisis de desviación ────────────────────────────────────────
        # Los precios de internet son UNITARIOS; el valor_total de la cotización
        # incluye N unidades (parseadas del nombre del ítem).
        precios_unitarios_internet = [
            (screenshots.get(lk.url) or ResultadoScreenshot(None, None, None, False)).precio_numero
            or lk.precio_numero
            for lk in prod.links
        ]
        precio_cotiz = precio_cotizacion_por_item.get(prod.item)
        unidades     = _extraer_unidades(prod.item)

        current_row += 1

        ws.merge_cells(f"A{current_row}:E{current_row}")
        h = ws.cell(row=current_row, column=1, value="ANÁLISIS DE PRECIO")
        h.font      = _font(bold=True, size=11, color=_C_WHITE)
        h.fill      = _fill(_C_DARK)
        h.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        # Fila: total cotizado
        ws.cell(row=current_row, column=1, value="Total cotización elegida:").font = _font(bold=True, size=10)
        ws.cell(row=current_row, column=2,
                value=_formatear_cop(precio_cotiz) if precio_cotiz else "No disponible"
                ).font = _font(bold=True, size=10)
        unid_label = f"({unidades} unidad{'es' if unidades != 1 else ''})"
        ws.cell(row=current_row, column=3, value=unid_label).font = _font(size=9, color="555555")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Filas: precios unitarios de internet + total ajustado
        for idx, (lk, p_unit) in enumerate(zip(prod.links, precios_unitarios_internet), 1):
            ws.cell(row=current_row, column=1,
                    value=f"  Precio unitario internet #{idx}:").font = _font(size=10, color="444444")
            if p_unit:
                ws.cell(row=current_row, column=2,
                        value=_formatear_cop(p_unit)).font = _font(size=10)
                if unidades > 1:
                    ws.cell(row=current_row, column=3,
                            value=f"× {unidades} = {_formatear_cop(p_unit * unidades)}"
                            ).font = _font(size=9, color="555555")
            else:
                ws.cell(row=current_row, column=2, value="No extraído").font = _font(size=10, color="999999")
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        # Fila: mediana unitaria + mediana total ajustada
        precios_validos = [p for p in precios_unitarios_internet if p]
        ws.cell(row=current_row, column=1,
                value="Mediana unitaria internet:").font = _font(bold=True, size=10)
        if precios_validos:
            med_unit  = _mediana(precios_validos)
            med_total = med_unit * unidades
            ws.cell(row=current_row, column=2,
                    value=_formatear_cop(med_unit)).font = _font(bold=True, size=10)
            if unidades > 1:
                ws.cell(row=current_row, column=3,
                        value=f"× {unidades} = {_formatear_cop(med_total)} (total)"
                        ).font = _font(size=9, color="555555")
        else:
            ws.cell(row=current_row, column=2, value="Sin datos").font = _font(size=10, color="999999")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Fila: resultado del análisis
        alerta = (
            _analizar_desviacion(precio_cotiz, precios_unitarios_internet, prod.item)
            if precio_cotiz else None
        )
        ws.merge_cells(f"A{current_row}:E{current_row}")
        if alerta:
            color_fondo = "FFCCCC" if alerta.supera_limite else "CCFFCC"
            color_texto = "CC0000" if alerta.supera_limite else "1A5C1A"
            rc = ws.cell(row=current_row, column=1, value=alerta.mensaje)
            rc.font      = Font(name="Arial", size=10, bold=True, color=color_texto)
            rc.fill      = _fill(color_fondo)
            rc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[current_row].height = 36
            if alerta.supera_limite:
                alertas_desviacion.append(
                    f"04-WEB [{prod.item}]: {alerta.mensaje}"
                )
        else:
            nc = ws.cell(row=current_row, column=1,
                         value="Sin datos suficientes para el análisis de desviación.")
            nc.font = _font(size=10, color="999999")
            ws.row_dimensions[current_row].height = 18

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(ruta_salida))
    logger.info("Excel cotizaciones web guardado: %s", ruta_salida)
    return alertas_desviacion
