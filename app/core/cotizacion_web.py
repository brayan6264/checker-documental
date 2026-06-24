"""
Búsqueda web de precios de referencia para cotizaciones seleccionadas.

Flujo por producto:
  1. buscar_links_openai()   — web_search_preview → URLs candidatas reales (no inventadas).
  2. tomar_screenshots()     — por ronda:
       a. Playwright visita las URLs → screenshot + precio DOM + base64.
       b. La visión recibe las capturas como IMÁGENES → valida producto + extrae precio.
       c. Si faltan válidas → busca URLs de reemplazo y repite.
       d. Se detiene al alcanzar el mínimo o agotar las rondas.
  3. generar_excel_cotizaciones() — Excel con capturas válidas + análisis de precio.
"""

import asyncio
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

# Modelos diferenciados por tarea (calidad donde importa, ahorro donde no):
#  - BÚSQUEDA: gpt-4o devuelve URLs reales de tiendas mucho mejor que mini.
#  - ANÁLISIS (visión): gpt-4o detecta capturas vacías/erróneas y lee precios con
#    mucha más fiabilidad que mini → menos falsos válidos y menos falsos inválidos.
#  - TÉRMINOS: tarea simple de texto, mini es suficiente y barato.
_MODELO_BUSQUEDA = "gpt-4o"
_MODELO_ANALISIS = "gpt-4o-mini"   # visión: mini es mucho más barato y suficiente para validar capturas
_MODELO_TERMINOS = "gpt-4o-mini"

# URLs que Playwright visita en la primera ronda (optimización de costo/tiempo)
_URLS_INICIALES = 5

# URLs por ronda de reemplazo (máximo)
_URLS_POR_RONDA = 5

# Máximo de rondas de búsqueda (inicial + reemplazos).
_MAX_RONDAS = 3

# Mínimo obligatorio de capturas válidas por producto.
_MIN_VALIDOS = 3

# ── Navegación Playwright (optimización de tiempo) ───────────────────────────
_NAV_TIMEOUT_MS  = 12_000   # antes 30s; los dominios muertos fallan rápido
_NAV_WAIT_MS     = 1_000    # antes 2.5s; espera de render más corta
_CONCURRENCIA_PW = 5        # URLs visitadas en paralelo por ronda

# Dominios que en la práctica siempre dan timeout/bloqueo (anti-bot) y solo
# desperdician tiempo. Se descartan antes de intentar capturarlos.
_DOMINIOS_LENTOS = {
    "jumbo.co", "makro.com.co", "sodimac.com.co", "linio.com.co",
}


def _url_lenta(url: str) -> bool:
    dominio = urlparse(url).netloc.lower().removeprefix("www.")
    return any(dominio == b or dominio.endswith("." + b) for b in _DOMINIOS_LENTOS)

# Tiendas reconocidas colombianas con inventario amplio y páginas estables.
# La ronda 1 prioriza EXCLUSIVAMENTE estas tiendas para maximizar capturas exitosas.
_TIENDAS_RECONOCIDAS = [
    "alkosto.com", "homecenter.com.co", "exito.com", "falabella.com.co",
    "ktronix.com", "panamericana.com.co", "jumbo.co", "makro.com.co",
    "linio.com.co", "sodimac.com.co", "tuvendedor.com.co", "olimpica.com.co",
    "flamingo.com.co", "cencosud.com.co", "lacomer.com.co", "pricesmart.com.co",
    "craftmaster.com.co", "ferreterias-abc.com", "construmart.com.co",
]

# Tiendas reconocidas para productos agropecuarios / seres vivos
_TIENDAS_AGRO = [
    "agroinsumos.com.co", "almacenagrario.com.co", "agropecuariamundo.com",
    "lahacienda.com.co", "cultivar.com.co", "bioagro.com.co",
    "agropecuariacolombia.com", "viverosonline.com.co",
]


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


# Palabras clave que indican que el producto es un ser vivo (planta, semilla, animal, etc.)
# Cuando la descripción los contiene, la búsqueda debe restringirse al departamento del proyecto.
_RE_SERES_VIVOS = re.compile(
    r'\b(?:'
    r'planta[s]?|semilla[s]?|arbusto[s]?|árbol|arboles|árboles|palma[s]?|cactus|suculenta[s]?'
    r'|flor(?:es)?|orquídea[s]?|helecho[s]?|hierba[s]?|yerba[s]?|follaje'
    r'|abono[s]?|compost|sustrato[s]?|tierra\s+(?:fértil|para\s+(?:siembra|cultivo))'
    r'|almácigo[s]?|esquejes?|bulbo[s]?|tubérculo[s]?|rizoma[s]?'
    r'|animal(?:es)?|mascota[s]?|aves?|gallina[s]?|pollo[s]?|cerdo[s]?|conejo[s]?'
    r'|bovino[s]?|ovino[s]?|caprino[s]?|pez|peces|alevino[s]?'
    r'|cultivo[s]?|siembra|cosecha|vivero'
    r')\b',
    re.IGNORECASE,
)


def _es_producto_ser_vivo(descripcion: str) -> bool:
    """Retorna True si la descripción corresponde a plantas, semillas u otros seres vivos."""
    return bool(_RE_SERES_VIVOS.search(descripcion or ""))


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


# ── PASO 0 — Normalización del término de búsqueda con IA ────────────────────

def _derivar_terminos_busqueda(seleccionadas: list, api_key: str) -> dict[str, str]:
    """
    Para CADA cotización deriva un término de búsqueda comercial limpio.

    Solución general: el texto de entrada puede ser cualquier cosa (un nombre
    corto, un párrafo técnico, instrucciones de uso, un prospecto médico,
    medidas, etc.). La IA lo reduce al producto comprable concreto: nombre
    comercial + características clave que identifican el ítem en una tienda.

    Retorna {item: termino_busqueda}. Si algo falla, cae a un recorte simple.
    """
    # Fallback genérico sin IA: primer fragmento antes de un punto/salto.
    def _fallback(sel) -> str:
        base = (sel.item or sel.descripcion or "").strip()
        return re.split(r'[.\n]', base)[0].strip()[:120] or base[:120]

    if not api_key:
        return {sel.item: _fallback(sel) for sel in seleccionadas}

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
    except Exception:
        return {sel.item: _fallback(sel) for sel in seleccionadas}

    entradas = [
        {
            "id":          i,
            "nombre":      (sel.item or "")[:300],
            "descripcion": (sel.descripcion or "")[:600],
        }
        for i, sel in enumerate(seleccionadas)
    ]

    import json as _json
    prompt = (
        "Eres un experto en compras. Para cada ítem de la lista, genera el TÉRMINO DE BÚSQUEDA "
        "más efectivo para encontrarlo a la venta en tiendas online colombianas.\n\n"
        "REGLAS para el término:\n"
        "- Debe ser el nombre comercial del producto físico comprable + sus características "
        "distintivas clave (capacidad, potencia, tamaño, material, modelo si aplica).\n"
        "- Máximo 12 palabras. Conciso, como lo buscaría un comprador.\n"
        "- IGNORA texto que no sirve para buscar: instrucciones de uso, modos de empleo, "
        "prospectos, dosis, descripciones de para qué sirve, condiciones de entrega, garantías.\n"
        "- Si el texto es un prospecto o instrucción (ej. de un medicamento o insumo), extrae "
        "el NOMBRE del producto en sí, no su modo de uso.\n"
        "- No inventes marcas que no aparezcan en el texto.\n\n"
        f"Ítems:\n{_json.dumps(entradas, ensure_ascii=False, indent=2)}\n\n"
        "Responde SOLO con JSON:\n"
        '{"terminos": [{"id": 0, "busqueda": "..."}]}'
    )

    try:
        response = client.responses.create(
            model=_MODELO_TERMINOS,
            input=[{"role": "user", "content": prompt}],
        )
        data = _extraer_json(response.output_text) or {}
        por_id = {
            int(t["id"]): str(t.get("busqueda", "")).strip()
            for t in (data.get("terminos") or [])
            if "id" in t
        }
    except Exception as exc:
        logger.warning("  Normalización de términos falló: %s — usando fallback", exc)
        por_id = {}

    resultado: dict[str, str] = {}
    for i, sel in enumerate(seleccionadas):
        termino = por_id.get(i, "").strip()
        resultado[sel.item] = termino or _fallback(sel)
        if termino:
            logger.info("  Término búsqueda '%s' → '%s'", (sel.item or "")[:40], termino)
    return resultado


# ── PASO 1 — Búsqueda de URLs con gpt-4o-mini + web_search_preview ───────────

def buscar_links_openai(
    seleccionadas: list,
    api_key: str,
    departamento: str = "",
) -> tuple[bool, list[ProductoLinks], str]:
    """
    Llama a gpt-4o-mini con web_search_preview para obtener 9 URLs candidatas
    por producto.  Retorna (ok, productos, resumen).

    Si `departamento` está informado y el producto es un ser vivo (planta, semilla,
    animal, etc.), las URLs deben ser de vendedores ubicados en ese departamento,
    ya que los seres vivos son difíciles de transportar.
    """
    try:
        from openai import OpenAI
    except ImportError:
        return False, [], "FALTA — openai no instalado"

    if not api_key:
        return False, [], "FALTA — OPENAI_API_KEY no configurada"

    client  = OpenAI(api_key=api_key)
    hoy_str = date.today().strftime("%d de %B de %Y")

    # Derivar términos de búsqueda limpios para TODOS los productos (1 sola llamada).
    terminos = _derivar_terminos_busqueda(seleccionadas, api_key)

    productos_out: list[ProductoLinks] = []
    errores: list[str] = []

    for sel in seleccionadas:
        # Término de búsqueda comercial normalizado por IA (general para cualquier producto).
        desc = terminos.get(sel.item) or (sel.item or "")[:120]

        es_vivo  = _es_producto_ser_vivo(desc)
        geo_regla = ""
        geo_nota  = ""
        if es_vivo and departamento:
            geo_regla = (
                f"6. RESTRICCIÓN GEOGRÁFICA OBLIGATORIA: este producto es un ser vivo "
                f"(planta, semilla, animal, etc.) y DEBE comprarse localmente. "
                f"TODAS las URLs deben ser de vendedores, viveros, granjas o distribuidores "
                f"ubicados en el departamento de {departamento} (Colombia). "
                f"No incluyas tiendas nacionales de e-commerce que no tengan presencia "
                f"o envío garantizado en {departamento}."
            )
            geo_nota = (
                f"\nBUSCA ESPECÍFICAMENTE en {departamento}: viveros, productores agropecuarios, "
                f"agrotiendas, cooperativas o distribuidores locales que vendan este producto."
            )
            logger.info(
                "  '%s': producto ser vivo detectado — búsqueda restringida a %s",
                sel.item, departamento,
            )

        prompt = f"""Hoy es {hoy_str}. Necesito comprar este producto en Colombia.

PRODUCTO A BUSCAR: {desc}{geo_nota}

USA LA HERRAMIENTA DE BÚSQUEDA WEB para encontrar páginas de tiendas colombianas donde se venda este producto.
Devuelve hasta 9 URLs de fichas de producto que REALMENTE hayas encontrado en los resultados de búsqueda.

═══════════════════════════════════════════════════════
⛔ REGLA #1 — LA MÁS IMPORTANTE: NO INVENTES URLs
═══════════════════════════════════════════════════════
- Cada URL que devuelvas DEBE ser una que apareció literalmente en los resultados de tu búsqueda web.
- PROHIBIDO construir, adivinar o deducir URLs a partir del nombre de una tienda.
  Ejemplo de lo que NUNCA debes hacer: inventar "tienda.com/" + nombre-del-producto.
- PROHIBIDO inventar IDs de producto, slugs o códigos (nada de "/product/1234567/", "/p/n0x1y2", etc.).
- Si solo encuentras 3 URLs reales, devuelve 3. NUNCA rellenes con URLs inventadas para llegar a un número.
- Es MUCHO mejor devolver pocas URLs verdaderas que muchas inventadas. Una URL inventada es un error grave.
- Antes de incluir una URL, verifica mentalmente: "¿esta URL exacta apareció en mi búsqueda?" Si no, descártala.

REGLAS adicionales para cada URL:
1. Debe llevar a la ficha de UN producto (no a categorías, búsquedas ni home de la tienda).
2. Preferible cada URL de un dominio distinto, pero pueden repetirse si son productos distintos reales.
3. PROHIBIDO MercadoLibre y todas sus variantes (mercadolibre.com.co, meli.com.co, etc.).
{geo_regla}

Responde ÚNICAMENTE con este JSON (sin texto extra, sin ```):
{{
  "links": [
    {{"url": "https://...(url real encontrada en la búsqueda)...", "precio": "$ ...", "precio_numero": null, "tienda": "dominio.com"}}
  ]
}}
Incluye solo las URLs reales que hallaste (pueden ser menos de 9). precio_numero: déjalo en null."""

        try:
            response = client.responses.create(
                model=_MODELO_BUSQUEDA,
                tools=[{"type": "web_search_preview"}],
                tool_choice={"type": "web_search_preview"},  # OBLIGA a ejecutar la búsqueda web
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
        # Itera TODOS los links devueltos (puede ser más de 9) y filtra bloqueados/duplicados.
        # El corte a _URLS_INICIALES se hace DESPUÉS del filtro para garantizar 9 válidos.
        for lk in (data.get("links") or []):
            if len(links) >= _URLS_INICIALES:
                break
            url = str(lk.get("url", "")).strip()
            if not url:
                continue
            if _url_bloqueada(url):
                logger.warning("  URL bloqueada (dominio excluido): %s", url)
                continue
            if _url_lenta(url):
                logger.info("  URL descartada (dominio lento/anti-bot): %s", url)
                continue
            dominio = urlparse(url).netloc.lower().removeprefix("www.")
            if dominio in dominios_vistos:
                continue
            dominios_vistos.add(dominio)

            precio_num  = lk.get("precio_numero")
            precio_txt  = str(lk.get("precio", "") or "")

            # Descartar precios en USD o moneda extranjera
            if any(s in precio_txt.upper() for s in ("USD", "US$", "EUR", "€", "MXN")):
                precio_num = None
                precio_txt = "N/A"

            if isinstance(precio_num, float):
                precio_num = int(precio_num)
            if precio_num is not None and not _es_precio_razonable(precio_num):
                precio_num = None

            links.append(LinkProducto(
                url=url,
                precio_texto=precio_txt or "N/A",
                precio_numero=precio_num,
            ))

        if links:
            productos_out.append(ProductoLinks(item=sel.item, descripcion=desc, links=links))
            logger.info("  '%s': %d/%d URL(s) candidatas obtenidas", sel.item, len(links), _URLS_INICIALES)
            if len(links) < _URLS_INICIALES:
                logger.warning("  '%s': solo %d links (objetivo: %d) — posibles dominios bloqueados o respuesta incompleta", sel.item, len(links), _URLS_INICIALES)
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
# ── PASO 2a — Captura CONCURRENTE de capturas (async, optimización de tiempo) ─

async def _extraer_precio_dom_async(page) -> tuple[str | None, int | None]:
    """Versión async de _extraer_precio_dom (JSON-LD → H1 → escaneo de texto)."""
    try:
        info = await page.evaluate(_JS_JSONLD)
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

    try:
        t = await page.evaluate(_JS_PRECIO_PRINCIPAL)
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

    try:
        texto = await page.evaluate("document.body.innerText") or ""
        candidatos = [
            _parsear_precio_cop(m.group(0))
            for m in _RE_PRECIO_COP.finditer(texto)
        ]
        candidatos = [n for n in candidatos if n and _es_precio_razonable(n)]
        if candidatos:
            from collections import Counter
            mc = Counter(candidatos).most_common(1)[0][0]
            return _formatear_cop(mc), mc
    except Exception:
        pass

    return None, None


def _capturar_concurrente(urls: list[str], dir_out: Path) -> dict:
    """
    Captura varias URLs EN PARALELO (hasta _CONCURRENCIA_PW a la vez).
    Retorna {url: (ruta_png, base64_png, precio_texto, precio_numero)}.
    Reemplaza el bucle serial: corta el tiempo de cada ronda ~N veces.
    """
    if not urls:
        return {}
    try:
        return asyncio.run(_capturar_async(list(urls), dir_out))
    except Exception as exc:
        logger.error("Captura concurrente falló: %s", exc)
        return {u: (None, None, None, None) for u in urls}


async def _capturar_async(urls: list[str], dir_out: Path) -> dict:
    from playwright.async_api import async_playwright

    resultados: dict = {}
    sem = asyncio.Semaphore(_CONCURRENCIA_PW)

    async def _ruta_handler(route):
        # Silenciar errores cuando la página/contexto ya se cerró: las peticiones
        # "en vuelo" al cerrar la página lanzan TargetClosedError; es inofensivo.
        try:
            if route.request.resource_type in ("font", "media"):
                await route.abort()
            else:
                await route.continue_()
        except Exception:
            pass

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        ctx = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="es-CO",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        await ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        await ctx.route("**/*", _ruta_handler)

        async def _una(url: str):
            async with sem:
                page = await ctx.new_page()
                try:
                    await page.goto(url, timeout=_NAV_TIMEOUT_MS, wait_until="domcontentloaded")
                    await page.wait_for_timeout(_NAV_WAIT_MS)
                    p_txt, p_num = await _extraer_precio_dom_async(page)
                    ruta = dir_out / (hashlib.md5(url.encode()).hexdigest()[:12] + ".png")
                    await page.screenshot(path=str(ruta), full_page=False)
                    b64 = base64.b64encode(ruta.read_bytes()).decode()
                    resultados[url] = (ruta, b64, p_txt, p_num)
                except Exception as exc:
                    logger.warning("  Screenshot error %s: %s", url, exc)
                    resultados[url] = (None, None, None, None)
                finally:
                    try:
                        await page.close()
                    except Exception:
                        pass

        await asyncio.gather(*[_una(u) for u in urls])
        # Quitar el handler de rutas antes de cerrar para evitar callbacks en vuelo
        try:
            await ctx.unroute_all(behavior="ignoreErrors")
        except Exception:
            pass
        await browser.close()

    return resultados


# ── PASO 2b/2c — Verificación visual de capturas (IMÁGENES, no PDF) ───────────

def _verificar_con_vision(
    descripcion: str,
    capturas: list[tuple[str, str]],   # [(url, base64_png), ...]
    api_key: str,
) -> list[dict]:
    """
    Envía las capturas como IMÁGENES a la visión para verificar cada página y
    extraer el precio. Enviar imágenes individuales (no un PDF) es más fiable:
    evita las negativas del modelo ("no puedo analizar PDFs") y mejora la lectura.

    Criterio de validez (NO ser rígido — productos similares SÍ valen):
      - válido: ficha de producto igual, similar, equivalente o del mismo sector.
      - inválido SOLO si: página en blanco / error 404 / CAPTCHA / listado-categoría /
        producto de categoría completamente distinta.

    Retorna lista de dicts: {url, valido, precio_texto, precio_numero, motivo}.
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    n = len(capturas)
    prompt_texto = f"""Eres un verificador de páginas de e-commerce colombianas.

Producto buscado: "{descripcion}"

Analiza las {n} captura(s) de pantalla y para CADA UNA devuelve un objeto JSON con:

1. "valido" (boolean):
   - true si la captura muestra una FICHA de producto que corresponde a "{descripcion}"
     o a un producto SIMILAR, EQUIVALENTE o DEL MISMO SECTOR/CATEGORÍA,
     QUE ESTÉ DISPONIBLE PARA COMPRA Y CON PRECIO VISIBLE.
     NO tiene que ser idéntico: distinta marca, modelo, presentación o tamaño SÍ vale.
   - false si: página en blanco, error 404/"no encontrado", CAPTCHA, pantalla de
     carga, listado de categorías/búsqueda (sin ficha), o un producto de una categoría
     completamente distinta (ej: se busca herramienta y aparece ropa).
   - false TAMBIÉN si: el producto está AGOTADO o SIN STOCK ("producto sin stock",
     "agotado", "no disponible"), o la imagen es un PLACEHOLDER sin foto real del producto
     (ej: "Estamos preparando la imagen de este producto"), o NO se ve ningún precio.
     Estos casos NO sirven como cotización y deben marcarse false.

2. "precio_texto" (string o null): el precio PRINCIPAL de venta visible, formato colombiano
   con puntos de miles (ej: "$ 1.234.567", "$ 89.900"). En Colombia los puntos separan miles.
   Ignora precios tachados. Si NO hay precio visible, null (y la página NO es válida).

3. "precio_numero" (integer o null): el mismo precio como entero sin puntos ni símbolos
   (ej: "$ 1.234.567" → 1234567). El número REAL que ves; si no lo ves, null. NO copies ejemplos.

4. "motivo" (string o null): si "valido" es false, explica brevemente por qué; si true, null.

Responde ÚNICAMENTE con un array JSON de {n} objeto(s) en el MISMO ORDEN que las imágenes,
sin texto adicional. Ejemplo para 2 imágenes:
[
  {{"valido": true, "precio_texto": "$ 1.234.567", "precio_numero": 1234567, "motivo": null}},
  {{"valido": false, "precio_texto": null, "precio_numero": null, "motivo": "error 404"}}
]"""

    content: list[dict] = [{"type": "text", "text": prompt_texto}]
    for _url, b64 in capturas:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64}", "detail": "high"},
        })

    try:
        resp = client.chat.completions.create(
            model=_MODELO_ANALISIS,
            messages=[{"role": "user", "content": content}],
            max_tokens=1000,
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
                    logger.warning("Vision: precio fuera de rango (%s) descartado", precio_num)
                    precio_num = None
                resultado.append({
                    "url":           capturas[i][0],
                    "valido":        bool(d.get("valido")),
                    "precio_texto":  d.get("precio_texto"),
                    "precio_numero": precio_num,
                    "motivo":        d.get("motivo") or "",
                })
            return resultado

        logger.warning(
            "Vision: array de longitud inesperada (esperado %d). Respuesta: %s",
            n, texto[:200],
        )
    except Exception as exc:
        logger.error("Vision API error: %s", exc)

    # Fallback conservador: marcar inválidos (no inflar el Excel con capturas dudosas)
    return [
        {"url": url, "valido": False, "precio_texto": None, "precio_numero": None,
         "motivo": "verificación de visión no disponible"}
        for url, _ in capturas
    ]


# ── PASO 2d — Búsqueda de URLs adicionales ───────────────────────────────────

def _buscar_links_reemplazo(
    descripcion: str,
    n: int,
    excluir: set[str],
    api_key: str,
    departamento: str = "",
) -> list[str]:
    """
    Busca con web_search hasta N URLs nuevas reales, excluyendo dominios ya intentados.
    Devuelve solo URLs realmente encontradas (puede ser menos de N); nunca inventadas.
    Si `departamento` informado y el producto es ser vivo, restringe a ese departamento.
    """
    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    excluir_dominios = sorted({
        urlparse(u).netloc.lower().removeprefix("www.")
        for u in excluir
    })

    es_vivo   = _es_producto_ser_vivo(descripcion)
    geo_regla = ""
    geo_nota  = ""
    if es_vivo and departamento:
        geo_regla = (
            f"5. RESTRICCIÓN GEOGRÁFICA OBLIGATORIA: este producto es un ser vivo. "
            f"TODOS los links deben ser de vendedores ubicados en el departamento de "
            f"{departamento} (Colombia). Busca viveros, granjas, agrotiendas o "
            f"distribuidores locales en {departamento}.\n"
        )
        geo_nota = f" en el departamento de {departamento}"

    prompt = (
        f"USA LA HERRAMIENTA DE BÚSQUEDA WEB para encontrar tiendas colombianas{geo_nota} que vendan:\n"
        f"PRODUCTO: {descripcion}\n\n"
        f"CONTEXTO: Ya intenté los dominios de abajo y fallaron. Busca en OTRAS tiendas distintas.\n"
        f"Dominios ya intentados (NO repetir): {', '.join(excluir_dominios) or 'ninguno'}\n\n"
        f"⛔ REGLA MÁS IMPORTANTE — NO INVENTES URLs NI DOMINIOS:\n"
        f"- Cada URL debe haber aparecido LITERALMENTE en los resultados de tu búsqueda web.\n"
        f"- PROHIBIDO inventar nombres de tiendas o dominios (ej: 'agroferreteria.com.co', "
        f"'maquinariapesada.com.co' y similares inventados NO existen y rompen el proceso).\n"
        f"- PROHIBIDO construir URLs pegando el nombre del producto al dominio de una tienda.\n"
        f"- PROHIBIDO inventar IDs, slugs o códigos de producto.\n"
        f"- Si solo encuentras 1 o 2 URLs reales, devuelve solo esas. NUNCA rellenes con inventadas.\n"
        f"- Una sola URL inventada es un error grave. Prefiero pocas reales que muchas falsas.\n\n"
        f"REGLAS adicionales:\n"
        f"1. URL directa a la ficha de UN producto (no categorías, no home, no búsquedas).\n"
        f"2. Dominio diferente a los ya intentados.\n"
        f"3. PROHIBIDO MercadoLibre y todas sus variantes.\n"
        f"{geo_regla}"
        f"\nResponde SOLO con JSON (solo las URLs reales halladas, pueden ser menos de {n}):\n"
        f'{{"links":[{{"url":"https://...(url real de la búsqueda)...","precio":"$ ...","precio_numero":null}}]}}'
    )

    try:
        response = client.responses.create(
            model=_MODELO_BUSQUEDA,
            tools=[{"type": "web_search_preview"}],
            tool_choice={"type": "web_search_preview"},  # OBLIGA a ejecutar la búsqueda web
            input=[{"role": "user", "content": prompt}],
        )
        data = _extraer_json(response.output_text)
        if data and isinstance(data, dict):
            urls = []
            for lk in (data.get("links") or [])[:n]:
                url = str(lk.get("url", "")).strip()
                if not url or url in excluir:
                    continue
                if _url_bloqueada(url) or _url_lenta(url):
                    logger.info("  Reemplazo descartado (dominio excluido/lento): %s", url)
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
    departamento: str = "",
) -> tuple[list[ProductoLinks], dict[str, ResultadoScreenshot]]:
    """
    Para cada producto:
      Ronda 1+:
        a. Playwright visita todas las URLs pendientes → screenshot + precio DOM + base64.
        b. La visión recibe las capturas como IMÁGENES → valida producto + extrae precio.
        c. Los válidos se acumulan; si faltan, se piden URLs de reemplazo → nueva ronda.
        d. Se detiene al alcanzar el mínimo de capturas válidas o agotar las rondas.
      Máx _MAX_RONDAS rondas en total.

    Retorna (productos_actualizados, {url: ResultadoScreenshot}).
    """
    dir_out.mkdir(parents=True, exist_ok=True)

    res: dict[str, ResultadoScreenshot]  = {}
    productos_actualizados: list[ProductoLinks] = []

    try:
        for producto in productos:
            desc            = producto.descripcion or producto.item
            urls_pendientes = [_normalizar_url(lk.url) for lk in producto.links if lk.url]
            urls_intentadas: set[str] = set(urls_pendientes)
            links_validos: list[LinkProducto] = []
            rondas_ok = 0

            for ronda in range(1, _MAX_RONDAS + 1):
                if not urls_pendientes:
                    break

                logger.info(
                    "  [%s] Ronda %d/%d — visitando %d URL(s) en paralelo",
                    producto.item, ronda, _MAX_RONDAS, len(urls_pendientes),
                )

                # ── a. Playwright CONCURRENTE: screenshot + precio DOM + base64
                datos_ronda = _capturar_concurrente(urls_pendientes, dir_out)

                # ── b. Verificación visual por IMÁGENES (no PDF) ──────────
                capturas_img = [
                    (url, b64) for url, (_, b64, _, _) in datos_ronda.items() if b64
                ]
                if api_key and capturas_img:
                    verifs = _verificar_con_vision(desc, capturas_img, api_key)
                    verif_por_url = {v["url"]: v for v in verifs}
                else:
                    # Sin IA: aceptar todas las que tienen captura
                    verif_por_url = {
                        url: {"valido": True, "precio_texto": None, "precio_numero": None, "motivo": ""}
                        for url, (ruta, _, _, _) in datos_ronda.items() if ruta
                    }

                # ── c. Clasificar válidos / inválidos ─────────────────────
                for url, (ruta, _b64, p_txt_dom, p_num_dom) in datos_ronda.items():
                    verif = verif_por_url.get(url)

                    if not verif or not verif.get("valido"):
                        res[url] = ResultadoScreenshot(None, None, None, False)
                        motivo = (verif or {}).get("motivo", "captura no disponible")
                        logger.warning("  [%s] ✗ inválido (%s): %s", producto.item, motivo, url)
                        continue

                    # La visión lee la MISMA imagen que va al Excel → fuente de verdad.
                    # DOM como respaldo solo si la visión no extrajo precio.
                    precio_num = p_num_dom or verif.get("precio_numero")
                    precio_txt = p_txt_dom or verif.get("precio_texto")
                    if precio_num and not _es_precio_razonable(precio_num):
                        precio_num = None
                        precio_txt = None

                    # Regla obligatoria: una cotización SOLO cuenta si tiene precio real.
                    # Sin precio (captura vacía, placeholder, sin stock) NO sirve como
                    # referencia → no cuenta para las 3 y disparará completación manual.
                    if not precio_num:
                        res[url] = ResultadoScreenshot(None, None, None, False)
                        logger.warning(
                            "  [%s] ✗ sin precio (no cuenta): %s", producto.item, url
                        )
                        continue

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
                        "  [%s] ✓ válido (%d/%d): %s → %s",
                        producto.item, len(links_validos), _MIN_VALIDOS, url,
                        precio_txt,
                    )

                rondas_ok += 1

                # ── d. ¿Tenemos suficientes? ──────────────────────────────
                needed = _MIN_VALIDOS - len(links_validos)
                if needed <= 0:
                    logger.info("  [%s] %d/%d válidos — listo.", producto.item, len(links_validos), _MIN_VALIDOS)
                    break

                if ronda >= _MAX_RONDAS:
                    logger.warning(
                        "  [%s] Límite de rondas alcanzado con %d/%d válidos.",
                        producto.item, len(links_validos), _MIN_VALIDOS,
                    )
                    break

                if not api_key:
                    break

                # Pedir un buffer pequeño de reemplazos, capado a _URLS_POR_RONDA
                # para controlar costo/tiempo (cada URL = navegación + visión).
                n_pedir = min(_URLS_POR_RONDA, needed + 2)
                logger.info(
                    "  [%s] Faltan %d — buscando %d URL(s) de reemplazo...",
                    producto.item, needed, n_pedir,
                )
                nuevas = _buscar_links_reemplazo(desc, n_pedir, urls_intentadas, api_key, departamento)
                nuevas = [_normalizar_url(u) for u in nuevas if u not in urls_intentadas]
                if not nuevas:
                    logger.warning("  [%s] Sin reemplazos disponibles.", producto.item)
                    break

                urls_pendientes = nuevas
                urls_intentadas.update(nuevas)

            if len(links_validos) < _MIN_VALIDOS:
                logger.warning(
                    "  [%s] ⚠ Final: solo %d/%d válidos tras %d ronda(s) — NO cumple el mínimo.",
                    producto.item, len(links_validos), _MIN_VALIDOS, rondas_ok,
                )
            else:
                logger.info(
                    "  [%s] Final: %d/%d válidos tras %d ronda(s).",
                    producto.item, len(links_validos), _MIN_VALIDOS, rondas_ok,
                )
            productos_actualizados.append(ProductoLinks(
                item=producto.item,
                descripcion=producto.descripcion,
                links=links_validos[:_MIN_VALIDOS],
            ))

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
    descripcion_original_por_item: dict[str, str] = {}
    if seleccionadas:
        for sel in seleccionadas:
            if sel.valor_total:
                precio_cotizacion_por_item[sel.item] = sel.valor_total
            if sel.descripcion:
                descripcion_original_por_item[sel.item] = sel.descripcion

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

    ws_i["A2"].value = f"Generado el  {hoy}"
    ws_i["A2"].font  = _font(size=10, color="555555")
    ws_i.row_dimensions[2].height = 16

    for col, hdr in enumerate(["Producto", "Descripción", "Hoja"], 1):
        c = ws_i.cell(row=4, column=col, value=hdr)
        c.font  = _font(bold=True, color=_C_WHITE)
        c.fill  = _fill(_C_MED)
        c.alignment = Alignment(horizontal="center")
    ws_i.row_dimensions[4].height = 18

    for r, prod in enumerate(productos, 5):
        desc_display = descripcion_original_por_item.get(prod.item) or prod.descripcion or ""
        ws_i.cell(row=r, column=1, value=prod.item)
        ws_i.cell(row=r, column=2, value=desc_display)
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
        _desc_orig = descripcion_original_por_item.get(prod.item) or prod.descripcion or "N/A"
        ws["A2"].value = f"Descripción: {_desc_orig}"
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
            # El precio SOLO se acepta si proviene de la página real:
            #  - ss.precio_numero: leído del DOM (JSON-LD/HTML) o de la visión sobre el screenshot.
            # Nunca se usa el precio "adivinado" por el web search (no es un dato verificable).
            if ss.precio_numero:
                precio_mostrar = ss.precio_texto or _formatear_cop(ss.precio_numero)
                precio_fuente  = "precio leído de la página"
            else:
                precio_mostrar = "No disponible"
                precio_fuente  = "no se pudo leer de la página"

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
        # SOLO se usan precios leídos de la página real (DOM/visión), nunca el
        # precio adivinado por el web search.
        precios_unitarios_internet = [
            (screenshots.get(lk.url) or ResultadoScreenshot(None, None, None, False)).precio_numero
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
                    f"04_VALIDACION_PRECIO_MERCADO [{prod.item}]: {alerta.mensaje}"
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
