"""
Búsqueda web de precios de referencia para cotizaciones seleccionadas.

Flujo por producto:
  1. buscar_links_openai()        — GPT-4.1 + web_search_preview → 3 URLs.
  2. tomar_screenshots()          — por cada ronda:
       a. Playwright visita URLs → screenshot + extracción de precio.
       b. GPT-4o-mini recibe las capturas → valida producto + extrae precio.
       c. URLs inválidas → GPT-4.1 busca reemplazos → vuelve a (a).
       (máx 3 rondas; se detiene al tener 3 resultados válidos)
  3. generar_excel_cotizaciones() — Excel con una hoja por producto.
"""

import base64
import hashlib
import json
import logging
import re
from datetime import date
from pathlib import Path
from typing import NamedTuple
from urllib.parse import urlparse

_RE_HOJA_INVALIDOS = re.compile(r'[\\/?*\[\]:]')

logger = logging.getLogger(__name__)

# GPT-4.1 + web_search_preview para búsqueda de links
_MODELO_BUSQUEDA = "gpt-4.1"
# GPT-4o-mini para verificación visual de screenshots (barato y preciso con imágenes)
_MODELO_VISION   = "gpt-4o-mini"


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
    links:       list[LinkProducto]   # hasta 3 candidatos válidos


# Regex para precios colombianos: $ 89.900 / 1.234.567 / COP 89.900
_RE_PRECIO_COP = re.compile(
    r'(?:COP\s*|cop\s*|\$\s*)?(\d{1,3}(?:[.,]\d{3})+)(?:\s*(?:COP|cop))?'
)

_PRECIO_MIN_COP = 5_000
_PRECIO_MAX_COP = 50_000_000


# ── Normalización de URLs ─────────────────────────────────────────────────────

def _normalizar_url(url: str) -> str:
    """Corrige subdominios conocidos que causan problemas en Playwright."""
    # tienda.exito.com muestra popup bloqueante; www.exito.com funciona bien
    if "tienda.exito.com" in url:
        url = url.replace("tienda.exito.com", "www.exito.com")
    return url


# ── OpenAI — búsqueda inicial de links ───────────────────────────────────────

def buscar_links_openai(
    seleccionadas: list,
    api_key: str,
) -> tuple[bool, list[ProductoLinks], str]:
    """
    Llama a GPT-4.1 con web_search_preview para buscar 3 URLs de tiendas
    colombianas por cada producto.  Retorna (ok, productos, resumen).
    """
    try:
        from openai import OpenAI
    except ImportError:
        return False, [], "FALTA — openai no instalado (pip install openai)"

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

Tu tarea: usa la búsqueda web para encontrar 3 URLs de páginas de producto en tiendas colombianas donde este artículo esté disponible para compra en este momento.

REGLAS (una URL que no cumpla TODAS será descartada):
1. URL directa a la ficha del producto — NO categorías, NO buscadores, NO la home de la tienda.
2. Stock disponible — verifica en la página que NO aparezca: "sin stock", "agotado", "no disponible". Si aparece, busca otra tienda.
3. Precio en PESOS COLOMBIANOS (COP). Excluye tiendas con precios en USD/EUR.
4. Cada URL de un DOMINIO DISTINTO. Prohibido repetir dominio.
5. Sin MercadoLibre (mercadolibre.com.co y variantes).
6. El precio que reportes debe ser el precio de venta actual visible hoy en la página.

Tiendas colombianas donde buscar (no exclusivas): alkosto.com, homecenter.com.co, exito.com, falabella.com.co, ktronix.com, panamericana.com.co, jumbo.co, makro.com.co, y tiendas especializadas del sector.

Responde ÚNICAMENTE con este JSON válido (sin texto extra, sin ```, sin explicaciones):
{{
  "links": [
    {{"url": "https://...", "precio": "$ 89.900", "precio_numero": 89900, "tienda": "alkosto.com"}},
    {{"url": "https://...", "precio": "$ 95.000", "precio_numero": 95000, "tienda": "homecenter.com.co"}},
    {{"url": "https://...", "precio": "$ 92.500", "precio_numero": 92500, "tienda": "exito.com"}}
  ]
}}

precio_numero: entero COP sin puntos ni símbolos."""

        try:
            response = client.responses.create(
                model=_MODELO_BUSQUEDA,
                tools=[{"type": "web_search_preview"}],
                input=[{"role": "user", "content": prompt}],
            )
            text = response.output_text
            logger.debug("[%s] gpt-4.1 respuesta (600 chars): %s", sel.item, text[:600])
        except Exception as exc:
            logger.error("OpenAI API error para '%s': %s", sel.item, exc)
            errores.append(str(exc))
            continue

        data = _extraer_json(text)
        if data is None:
            logger.error("JSON inválido para '%s'. Respuesta: %s", sel.item, text[:600])
            errores.append(f"JSON inválido para {sel.item}")
            continue

        links: list[LinkProducto] = []
        dominios_vistos: set[str] = set()
        for link_data in (data.get("links") or [])[:3]:
            url = str(link_data.get("url", "")).strip()
            if not url:
                continue
            dominio = urlparse(url).netloc.lower().removeprefix("www.")
            if dominio in dominios_vistos:
                continue
            dominios_vistos.add(dominio)

            precio_num = link_data.get("precio_numero")
            if isinstance(precio_num, float):
                precio_num = int(precio_num)
            if precio_num and not _es_precio_razonable(precio_num):
                precio_num = None

            links.append(LinkProducto(
                url=url,
                precio_texto=str(link_data.get("precio", "N/A")),
                precio_numero=precio_num,
            ))

        if links:
            productos_out.append(ProductoLinks(item=sel.item, descripcion=desc, links=links))
            logger.info("  Producto '%s': %d URL(s) para verificar", sel.item, len(links))
        else:
            logger.warning("  Producto '%s': sin enlaces válidos en la respuesta", sel.item)
            errores.append(f"Sin enlaces para {sel.item}")

    if not productos_out:
        return False, [], "GPT-4.1 no retornó enlaces para ningún producto"

    total   = sum(len(p.links) for p in productos_out)
    resumen = f"OK — {len(productos_out)} producto(s), {total} URL(s) iniciales"
    if errores:
        resumen += f" (errores: {'; '.join(errores)})"
    return True, productos_out, resumen


def _extraer_json(text: str) -> dict | list | None:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    m = re.search(r"```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    m = re.search(r"(\{.*\}|\[.*\])", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return None


# ── Playwright — screenshot + extracción de precio ───────────────────────────

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
                    availability: offers.availability  ?? null,
                };
            }
        }
        return null;
    };
    for (const s of document.querySelectorAll('script[type="application/ld+json"]')) {
        try {
            const r = buscar(JSON.parse(s.textContent));
            if (r) return r;
        } catch(e) {}
    }
    return null;
}
"""

_JS_PRECIO_PRINCIPAL = """
() => {
    const EXCLUIR_CARD = '[class*="productCard"],[class*="ProductCard"],' +
                         '[class*="shelf-item"],[class*="ShelfItem"],' +
                         '[class*="carousel"],[class*="Carousel"],' +
                         '[class*="related"],[class*="Related"],' +
                         '[class*="recommendation"],[class*="Recommendation"],' +
                         '[class*="priceOfferButton"]';

    const EXCLUIR_TACHADO = /dashed|Dashed|original|Original|old-price|before|tachado|list-price|listPrice|regularPrice/;

    const PRECIO_SELS = [
        '[itemprop="price"]',
        '[data-price]',
        '[class*="sellingPrice"]:not([class*="list"])',
        '[class*="SellingPrice"]:not([class*="list"])',
        '[class*="container__price"]',
        '[class*="ProductPrice_container__price"]',
        '[class*="selling-price"]',
        '[class*="price-final"]',
        '[class*="price--sale"]',
        '[class*="price--selling"]',
        '[class*="current-price"]',
        '[class*="precio-actual"]',
        '[class*="price-effective"]',
        '[class*="effectivePrice"]',
        '[class*="price-box"]',
        '[class*="product-prices__effective"]',
    ];

    const extraerTexto = el => {
        const content = el.getAttribute('content');
        if (content && /\\d/.test(content)) return content;
        return el.textContent.trim();
    };

    const esValido = el => {
        if (el.closest(EXCLUIR_CARD)) return false;
        const cls = el.className || '';
        if (EXCLUIR_TACHADO.test(cls)) return false;
        return true;
    };

    const h1 = document.querySelector('h1');
    if (h1) {
        let container = h1.parentElement;
        for (let i = 0; i < 12; i++) {
            if (!container || container === document.body) break;
            for (const sel of PRECIO_SELS) {
                const els = Array.from(container.querySelectorAll(sel)).filter(esValido);
                for (const el of els) {
                    const t = extraerTexto(el);
                    if (t && /\\d{3}/.test(t)) return t;
                }
            }
            container = container.parentElement;
        }
    }

    for (const sel of PRECIO_SELS) {
        const el = Array.from(document.querySelectorAll(sel)).find(esValido);
        if (el) {
            const t = extraerTexto(el);
            if (t && /\\d{3}/.test(t)) return t;
        }
    }
    return null;
}
"""


def _screenshot_una_url(
    page,
    url: str,
    dir_out: Path,
) -> tuple[Path | None, str | None, str | None, int | None]:
    """
    Navega a la URL, extrae precio con Playwright y toma screenshot.
    Retorna (ruta_png, base64_png, precio_texto, precio_numero).
    """
    fname = hashlib.md5(url.encode()).hexdigest()[:12] + ".png"
    ruta  = dir_out / fname
    try:
        page.goto(url, timeout=30_000, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)

        precio_txt, precio_num = _extraer_precio_pagina(page)
        page.screenshot(path=str(ruta), full_page=False)
        b64 = base64.b64encode(ruta.read_bytes()).decode()
        return ruta, b64, precio_txt, precio_num

    except Exception as exc:
        logger.warning("  Screenshot error %s: %s", url, exc)
        return None, None, None, None


# ── OpenAI — verificación visual de screenshots ───────────────────────────────

def _verificar_con_vision(
    descripcion: str,
    capturas: list[tuple[str, str]],   # [(url, base64_png), ...]
    api_key: str,
) -> list[dict]:
    """
    Envía las capturas a GPT-4o (visión) para verificar si cada página
    muestra el producto correcto y extraer el precio visible.

    - "valido": true si la página es una ficha de producto que corresponde al
      producto buscado (aunque no se vea el precio).
    - Una página de error, categoría, buscador, o producto completamente
      diferente es inválida.
    - El precio es un campo aparte: puede ser null aunque la página sea válida.

    Retorna lista de dicts: {url, valido, precio_texto, precio_numero, motivo}.
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    n = len(capturas)
    prompt_texto = f"""Eres un asistente de verificación de páginas de e-commerce colombianas.

Producto buscado: "{descripcion}"

Analiza las {n} imagen(es) de capturas de pantalla de tiendas en línea y para CADA UNA devuelve un objeto JSON con:

1. "valido" (boolean): true si la captura muestra una página de ficha de producto que corresponde a "{descripcion}" o un producto muy similar/equivalente. false si es una página de error, de categoría, de búsqueda, un producto completamente diferente, o una página en blanco.

2. "precio_texto" (string o null): el precio PRINCIPAL de venta visible en la imagen, en formato colombiano con puntos como separadores de miles. Ejemplos válidos: "$ 1.234.567", "$ 89.900", "1.796.826". Busca el precio más prominente en la página (generalmente junto al nombre del producto). IMPORTANTE: en Colombia los puntos separan miles (1.234.567 = un millón doscientos...). Si no hay precio visible, devuelve null.

3. "precio_numero" (integer o null): el mismo precio como número entero sin puntos ni símbolos. Ejemplo: si el precio es "$ 1.234.567" devuelve 1234567. Si no hay precio, devuelve null.

4. "motivo" (string o null): si "valido" es false, explica brevemente por qué. Si es true, devuelve null.

Responde ÚNICAMENTE con un array JSON de {n} objeto(s) en el mismo orden que las imágenes, sin texto adicional ni explicaciones fuera del JSON.

Ejemplo de respuesta para 2 imágenes:
[
  {{"valido": true, "precio_texto": "$ 1.234.567", "precio_numero": 1234567, "motivo": null}},
  {{"valido": false, "precio_texto": null, "precio_numero": null, "motivo": "página de error 404"}}
]"""

    content: list[dict] = [{"type": "text", "text": prompt_texto}]
    for _url, b64 in capturas:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"},
        })

    try:
        resp = client.chat.completions.create(
            model=_MODELO_VISION,
            messages=[{"role": "user", "content": content}],
            max_tokens=800,
        )
        texto = resp.choices[0].message.content or ""
        logger.debug("Vision respuesta: %s", texto[:500])
        datos = _extraer_json(texto)

        if isinstance(datos, list) and len(datos) == n:
            resultado = []
            for i, d in enumerate(datos):
                precio_num = d.get("precio_numero")
                if isinstance(precio_num, float):
                    precio_num = int(precio_num)
                if precio_num and not _es_precio_razonable(precio_num):
                    logger.warning(
                        "Vision: precio fuera de rango (%s) descartado para %s",
                        precio_num, capturas[i][0],
                    )
                    precio_num = None
                resultado.append({
                    "url":           capturas[i][0],
                    "valido":        bool(d.get("valido")),
                    "precio_texto":  d.get("precio_texto"),
                    "precio_numero": precio_num,
                    "motivo":        d.get("motivo") or "",
                })
            return resultado
        else:
            logger.warning(
                "Vision: array de longitud inesperada (esperado %d, respuesta: %s). "
                "Marcando todos como válidos.",
                n, texto[:200],
            )
    except Exception as exc:
        logger.error("Vision API error: %s", exc)

    # Fallback: asumir válidos para no bloquear el flujo
    return [
        {"url": url, "valido": True, "precio_texto": None, "precio_numero": None, "motivo": ""}
        for url, _ in capturas
    ]


# ── OpenAI — búsqueda de links de reemplazo ──────────────────────────────────

def _buscar_links_reemplazo(
    descripcion: str,
    n: int,
    excluir: set[str],
    api_key: str,
) -> list[str]:
    """
    Llama a GPT-4.1 + web_search_preview para obtener N URLs adicionales,
    excluyendo dominios ya intentados.  Retorna lista de URLs.
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    excluir_dominios = sorted({
        urlparse(u).netloc.lower().removeprefix("www.")
        for u in excluir
    })

    prompt = (
        f"Necesito {n} URL(s) alternativa(s) de tiendas colombianas para comprar HOY:\n"
        f"PRODUCTO: {descripcion}\n\n"
        f"Los siguientes dominios ya fueron verificados y NO funcionaron: "
        f"{', '.join(excluir_dominios) or 'ninguno'}.\n"
        f"Busca en OTROS dominios completamente diferentes.\n\n"
        f"REGLAS:\n"
        f"1. URL directa a la ficha del producto (no categorías ni buscadores)\n"
        f"2. Producto en stock en Colombia\n"
        f"3. Precio en pesos colombianos (COP)\n"
        f"4. Dominio diferente a los excluidos arriba\n"
        f"5. Sin MercadoLibre\n\n"
        f"Responde SOLO con JSON:\n"
        f'{{"links": [{{"url": "https://...", "precio": "$ 89.900", "precio_numero": 89900}}]}}'
    )

    try:
        response = client.responses.create(
            model=_MODELO_BUSQUEDA,
            tools=[{"type": "web_search_preview"}],
            input=[{"role": "user", "content": prompt}],
        )
        data = _extraer_json(response.output_text)
        if data and isinstance(data, dict):
            urls = []
            for lk in (data.get("links") or [])[:n]:
                url = str(lk.get("url", "")).strip()
                if url and url not in excluir:
                    urls.append(url)
            return urls
    except Exception as exc:
        logger.error("Búsqueda reemplazo error: %s", exc)

    return []


# ── Orquestador principal ─────────────────────────────────────────────────────

def tomar_screenshots(
    productos: list[ProductoLinks],
    dir_out: Path,
    api_key: str = "",
) -> tuple[list[ProductoLinks], dict[str, ResultadoScreenshot]]:
    """
    Para cada producto ejecuta el ciclo de verificación:
      Ronda 1-3:
        a. Playwright visita URLs pendientes → screenshot + precio DOM.
        b. GPT-4o-mini recibe capturas → valida si es el producto correcto.
        c. Los válidos se acumulan; por cada inválido se pide un reemplazo.
        d. Si se llega a 3 válidos o no quedan rondas, se para.

    Retorna (productos_actualizados, {url: ResultadoScreenshot}).
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout  # noqa: F401

    dir_out.mkdir(parents=True, exist_ok=True)

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

                # Límite de seguridad: máximo 15 URLs intentadas por producto para
                # garantizar siempre 3 válidos sin iterar indefinidamente.
                MAX_URLS_TOTAL  = 15
                total_intentadas = len(urls_pendientes)
                ronda            = 0

                while len(links_validos) < 3 and urls_pendientes:
                    ronda += 1
                    logger.info(
                        "  [%s] Ronda %d — %d URL(s) pendientes (%d válidos hasta ahora)",
                        producto.item, ronda, len(urls_pendientes), len(links_validos),
                    )

                    # ── Paso A: Playwright screenshot + precio DOM ────────────
                    datos_ronda: dict[str, tuple] = {}
                    for url in urls_pendientes:
                        ruta, b64, precio_txt, precio_num = _screenshot_una_url(page, url, dir_out)
                        datos_ronda[url] = (ruta, b64, precio_txt, precio_num)

                    # ── Paso B: Vision AI — validación + extracción precio ────
                    capturas_con_img = [
                        (url, b64)
                        for url, (_, b64, _, _) in datos_ronda.items()
                        if b64
                    ]

                    if api_key and capturas_con_img:
                        verificaciones = _verificar_con_vision(desc, capturas_con_img, api_key)
                        verif_por_url  = {v["url"]: v for v in verificaciones}
                    else:
                        verif_por_url = {}

                    # ── Paso C: clasificar válidos / inválidos ────────────────
                    for url, (ruta, b64, precio_txt_pw, precio_num_pw) in datos_ronda.items():
                        verif = verif_por_url.get(url)

                        if verif:
                            if verif["valido"]:
                                # Vision lee la misma imagen que va al Excel → es la fuente de verdad.
                                # Playwright como respaldo si vision no extrajo precio.
                                precio_num = verif["precio_numero"] or precio_num_pw
                                precio_txt = verif["precio_texto"]  or precio_txt_pw
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
                                logger.info(
                                    "  [%s] ✓ válido (%d/3): %s → %s",
                                    producto.item, len(links_validos), url,
                                    precio_txt or "precio no extraído",
                                )
                            else:
                                res[url] = ResultadoScreenshot(
                                    ruta=None, precio_texto=None, precio_numero=None, sin_stock=False
                                )
                                logger.warning(
                                    "  [%s] ✗ inválido (%s): %s",
                                    producto.item, verif.get("motivo", ""), url,
                                )
                        else:
                            # Sin verificación AI — aceptar todo con precios de Playwright
                            res[url] = ResultadoScreenshot(
                                ruta=ruta,
                                precio_texto=precio_txt_pw,
                                precio_numero=precio_num_pw,
                                sin_stock=False,
                            )
                            links_validos.append(LinkProducto(
                                url=url,
                                precio_texto=precio_txt_pw or "N/A",
                                precio_numero=precio_num_pw,
                            ))

                    # ── Paso D: ¿Necesitamos más? ────────────────────────────
                    needed = 3 - len(links_validos)
                    if needed <= 0:
                        logger.info("  [%s] 3/3 válidos — listo.", producto.item)
                        break

                    if not api_key:
                        logger.warning("  [%s] Sin API key — no se pueden pedir reemplazos.", producto.item)
                        break

                    if total_intentadas >= MAX_URLS_TOTAL:
                        logger.warning(
                            "  [%s] Límite de %d URLs alcanzado con solo %d/3 válidos.",
                            producto.item, MAX_URLS_TOTAL, len(links_validos),
                        )
                        break

                    # Pedir URLs de reemplazo
                    logger.info("  [%s] Necesito %d reemplazo(s)...", producto.item, needed)
                    nuevas = _buscar_links_reemplazo(desc, needed, urls_intentadas, api_key)
                    nuevas = [_normalizar_url(u) for u in nuevas if u not in urls_intentadas]

                    if not nuevas:
                        logger.warning("  [%s] OpenAI no devolvió URLs de reemplazo.", producto.item)
                        break

                    urls_pendientes = nuevas
                    urls_intentadas.update(nuevas)
                    total_intentadas += len(nuevas)

                logger.info(
                    "  [%s] Resultado final: %d/3 válidos (intentadas: %d URL(s)).",
                    producto.item, len(links_validos), total_intentadas,
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


# ── Helpers de precio ─────────────────────────────────────────────────────────

def _es_precio_razonable(num: int) -> bool:
    return _PRECIO_MIN_COP <= num <= _PRECIO_MAX_COP


def _precio_desde_jsonld(valor) -> int | None:
    if valor is None:
        return None
    s   = str(valor).strip()
    num = _parsear_precio_cop(s)
    if num and _es_precio_razonable(num):
        return num
    try:
        num = int(float(s))
        if _es_precio_razonable(num):
            return num
    except (ValueError, OverflowError):
        pass
    return None


def _extraer_precio_pagina(page) -> tuple[str | None, int | None]:
    """
    Extrae precio de venta del producto principal.
    1. JSON-LD  2. JS cerca del H1  3. Escaneo de texto completo
    """
    # 1. JSON-LD
    try:
        info = page.evaluate(_JS_JSONLD)
        if info:
            currency = (info.get("currency") or "").upper()
            if currency in ("COP", ""):
                num = _precio_desde_jsonld(info.get("price"))
                if num:
                    return _formatear_cop(num), num
    except Exception:
        pass

    # 2. JS cerca del H1
    try:
        texto_precio = page.evaluate(_JS_PRECIO_PRINCIPAL)
        if texto_precio:
            num = _parsear_precio_cop(texto_precio)
            if not num:
                matches = _RE_PRECIO_COP.findall(texto_precio)
                nums = [_parsear_precio_cop(m) for m in matches if _parsear_precio_cop(m)]
                nums = [n for n in nums if _es_precio_razonable(n)]
                if nums:
                    num = max(nums)
            if num and _es_precio_razonable(num):
                return _formatear_cop(num), num
    except Exception:
        pass

    # 3. Escaneo de texto completo
    try:
        texto_pagina = page.evaluate("document.body.innerText") or ""
        candidatos: list[int] = []
        for m in _RE_PRECIO_COP.finditer(texto_pagina):
            num = _parsear_precio_cop(m.group(0))
            if num and _es_precio_razonable(num):
                candidatos.append(num)
        if candidatos:
            from collections import Counter
            mas_comun = Counter(candidatos).most_common(1)[0][0]
            return _formatear_cop(mas_comun), mas_comun
    except Exception:
        pass

    return None, None


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


# ── Análisis de desviación de precio ─────────────────────────────────────────

class AlertaDesviacion:
    """Resultado del análisis de precio cotización vs. promedio internet."""
    __slots__ = ("promedio_internet", "precio_cotizacion", "desviacion_pct",
                 "limite_pct", "supera_limite", "mensaje")

    def __init__(
        self,
        promedio_internet: int,
        precio_cotizacion: int,
        desviacion_pct: float,
        limite_pct: float,
        supera_limite: bool,
        mensaje: str,
    ):
        self.promedio_internet  = promedio_internet
        self.precio_cotizacion  = precio_cotizacion
        self.desviacion_pct     = desviacion_pct
        self.limite_pct         = limite_pct
        self.supera_limite      = supera_limite
        self.mensaje            = mensaje


def _analizar_desviacion(
    precio_cotizacion: int,
    precios_internet: list[int],
) -> AlertaDesviacion | None:
    """
    Compara el precio de la cotización seleccionada contra el promedio de los
    precios encontrados en internet.

    Reglas:
      - Precio cotización < 500.000 COP → límite de desviación: 50 %
      - Precio cotización ≥ 500.000 COP → límite de desviación: 20 %

    La desviación se mide como:
        (promedio_internet - precio_cotizacion) / precio_cotizacion × 100

    Si promedio_internet > precio_cotizacion × (1 + límite/100) → alerta.
    Retorna None si no hay precios de internet disponibles.
    """
    precios_validos = [p for p in precios_internet if p and p > 0]
    if not precios_validos or not precio_cotizacion:
        return None

    promedio = int(sum(precios_validos) / len(precios_validos))
    desviacion_pct = (promedio - precio_cotizacion) / precio_cotizacion * 100
    limite_pct = 50.0 if precio_cotizacion < 500_000 else 20.0
    supera = desviacion_pct > limite_pct

    if supera:
        exceso = desviacion_pct - limite_pct
        mensaje = (
            f"⚠ ALERTA: el promedio de internet ($ {promedio:,}".replace(",", ".") +
            f") supera el precio cotizado en {desviacion_pct:.1f}% "
            f"(límite permitido: {limite_pct:.0f}%, exceso: {exceso:.1f}%)"
        )
    else:
        mensaje = (
            f"✓ OK: el promedio de internet ($ {promedio:,}".replace(",", ".") +
            f") está dentro del rango permitido "
            f"(desviación: {desviacion_pct:+.1f}%, límite: ±{limite_pct:.0f}%)"
        )

    return AlertaDesviacion(
        promedio_internet=promedio,
        precio_cotizacion=precio_cotizacion,
        desviacion_pct=desviacion_pct,
        limite_pct=limite_pct,
        supera_limite=supera,
        mensaje=mensaje,
    )


# ── Excel ─────────────────────────────────────────────────────────────────────

def generar_excel_cotizaciones(
    id_unico: str,
    productos: list[ProductoLinks],
    screenshots: dict[str, ResultadoScreenshot],
    ruta_salida: Path,
    seleccionadas: list | None = None,   # list[CotizacionSeleccionada]
) -> None:
    """Genera el Excel de referencia de precios web (una hoja por producto + índice)."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.drawing.image import Image as XLImage

    # Índice rápido para buscar precio de cotización por nombre de item
    precio_cotizacion_por_item: dict[str, int] = {}
    if seleccionadas:
        for sel in seleccionadas:
            if sel.valor_total:
                precio_cotizacion_por_item[sel.item] = sel.valor_total

    hoy = date.today().strftime("%d/%m/%Y")

    def _sheet_name(name: str) -> str:
        return _RE_HOJA_INVALIDOS.sub("_", name)[:31]

    _C_DARK  = "1F5C2E"
    _C_MED   = "2E7D32"
    _C_ODD   = "E8F5E9"
    _C_EVEN  = "F5F5F5"
    _C_WHITE = "FFFFFF"

    def _fill(hex6: str) -> PatternFill:
        return PatternFill("solid", start_color=hex6, end_color=hex6)

    def _font(bold=False, size=11, color="000000", name="Arial") -> Font:
        return Font(bold=bold, size=size, color=color, name=name)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Hoja índice ──────────────────────────────────────────────────────
    ws_i = wb.create_sheet("Índice")
    ws_i.merge_cells("A1:E1")
    ws_i["A1"].value     = f"Cotizaciones de referencia web — {id_unico}"
    ws_i["A1"].font      = _font(bold=True, size=14, color=_C_WHITE)
    ws_i["A1"].fill      = _fill(_C_DARK)
    ws_i["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_i.row_dimensions[1].height = 32

    ws_i["A2"].value = f"Generado: {hoy}  |  Modelo búsqueda: {_MODELO_BUSQUEDA}  |  Verificación: {_MODELO_VISION}"
    ws_i["A2"].font  = _font(size=10, color="555555")
    ws_i.row_dimensions[2].height = 16

    for col, hdr in enumerate(["Producto", "Descripción", "Hoja"], 1):
        c = ws_i.cell(row=4, column=col, value=hdr)
        c.font      = _font(bold=True, color=_C_WHITE)
        c.fill      = _fill(_C_MED)
        c.alignment = Alignment(horizontal="center")
    ws_i.row_dimensions[4].height = 18

    for r, prod in enumerate(productos, 5):
        ws_i.cell(row=r, column=1, value=prod.item)
        ws_i.cell(row=r, column=2, value=prod.descripcion or "")
        ws_i.cell(row=r, column=3, value=_sheet_name(prod.item))
    ws_i.column_dimensions["A"].width = 35
    ws_i.column_dimensions["B"].width = 60
    ws_i.column_dimensions["C"].width = 35

    # ── Una hoja por producto ─────────────────────────────────────────────
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
        ws["D2"].value     = f"Fecha consulta: {hoy}  |  ID: {id_unico}"
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
            if ss.sin_stock:
                precio_mostrar = "SIN STOCK"
                precio_fuente  = "verificado en página"
            elif ss.precio_texto:
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

            img_path = ss.ruta if not ss.sin_stock else None
            if img_path and img_path.exists():
                try:
                    xl_img        = XLImage(str(img_path))
                    xl_img.width  = IMG_W_PX
                    xl_img.height = IMG_H_PX
                    ws.add_image(xl_img, f"D{img_row}")
                except Exception as exc:
                    logger.warning("No se pudo insertar imagen en Excel: %s", exc)
                    ws.cell(row=img_row, column=4, value="[Captura no disponible]").font = _font(color="999999")
            elif ss.sin_stock:
                c = ws.cell(row=img_row, column=4, value="⚠ PRODUCTO SIN STOCK EN ESTA TIENDA")
                c.font = Font(name="Arial", size=12, bold=True, color="CC0000")
            else:
                ws.cell(row=img_row, column=4, value="[Captura no disponible]").font = _font(color="999999")

            current_row += IMG_ROWS + 1

        # ── Bloque de análisis de desviación ─────────────────────────────
        precios_internet = [
            (screenshots.get(lk.url) or ResultadoScreenshot(None, None, None, False)).precio_numero
            or lk.precio_numero
            for lk in prod.links
        ]
        precio_cotiz = precio_cotizacion_por_item.get(prod.item)

        current_row += 1  # fila de separación

        # Encabezado del bloque
        ws.merge_cells(f"A{current_row}:E{current_row}")
        hdr_cell           = ws.cell(row=current_row, column=1, value="ANÁLISIS DE PRECIO")
        hdr_cell.font      = _font(bold=True, size=11, color=_C_WHITE)
        hdr_cell.fill      = _fill(_C_DARK)
        hdr_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        # Fila: precio de la cotización seleccionada
        ws.cell(row=current_row, column=1, value="Precio cotización elegida:").font = _font(bold=True, size=10)
        val_cotiz = _formatear_cop(precio_cotiz) if precio_cotiz else "No disponible"
        c = ws.cell(row=current_row, column=2, value=val_cotiz)
        c.font      = _font(bold=True, size=10)
        c.alignment = Alignment(horizontal="left")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Filas: precios individuales de internet
        for idx, (lk, p_int) in enumerate(zip(prod.links, precios_internet), 1):
            ws.cell(row=current_row, column=1, value=f"  Precio internet #{idx}:").font = _font(size=10, color="444444")
            val_txt = _formatear_cop(p_int) if p_int else "No extraído"
            ws.cell(row=current_row, column=2, value=val_txt).font = _font(size=10)
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        # Fila: promedio internet
        precios_con_valor = [p for p in precios_internet if p]
        if precios_con_valor:
            promedio_int = int(sum(precios_con_valor) / len(precios_con_valor))
            ws.cell(row=current_row, column=1, value="Promedio internet:").font = _font(bold=True, size=10)
            ws.cell(row=current_row, column=2, value=_formatear_cop(promedio_int)).font = _font(bold=True, size=10)
        else:
            ws.cell(row=current_row, column=1, value="Promedio internet:").font = _font(bold=True, size=10)
            ws.cell(row=current_row, column=2, value="Sin datos").font = _font(size=10, color="999999")
        ws.row_dimensions[current_row].height = 18
        current_row += 1

        # Fila: resultado del análisis
        alerta = _analizar_desviacion(precio_cotiz, precios_internet) if precio_cotiz else None
        if alerta:
            color_fondo = "FFCCCC" if alerta.supera_limite else "CCFFCC"
            color_texto = "CC0000" if alerta.supera_limite else "1A5C1A"
            ws.merge_cells(f"A{current_row}:E{current_row}")
            result_cell           = ws.cell(row=current_row, column=1, value=alerta.mensaje)
            result_cell.font      = Font(name="Arial", size=10, bold=True, color=color_texto)
            result_cell.fill      = _fill(color_fondo)
            result_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[current_row].height = 30
        else:
            ws.merge_cells(f"A{current_row}:E{current_row}")
            nc = ws.cell(row=current_row, column=1, value="Sin datos suficientes para el análisis de desviación.")
            nc.font = _font(size=10, color="999999")
            ws.row_dimensions[current_row].height = 18

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(ruta_salida))
    logger.info("Excel cotizaciones web guardado: %s", ruta_salida)
