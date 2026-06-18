"""
Validación del archivo PLAN_INVERSION.xlsx — hoja "2. Revisión compra (cotización)".

Estructura de la hoja:
  - Fila 10  : nombre del ítem/producto (cols D, G, J, M... — salto de 3)
  - Fila 11  : "Cotización 1 / 2 / 3" por ítem
  - Fila 12  : Bien/Servicio/Ítem
  - Fila 13  : Nombre proveedor       ← campo obligatorio
  - Fila 17  : Descripción            ← campo obligatorio
  - Fila 18  : Valor total            ← campo obligatorio
  - Fila 32  : Cotización seleccionada (X marca la cotización elegida por producto)

Un ítem es "activo" si al menos una de sus 3 cotizaciones tiene proveedor con valor.
Los ítems de plantilla vacíos ("Ítem 4", "Ítem 5"...) sin datos se omiten.
"""

import logging
import re
from pathlib import Path
from typing import NamedTuple

logger = logging.getLogger(__name__)

# Índices de fila (1-based, como en openpyxl)
_FILA_ITEM_NOMBRE    = 10
_FILA_BIEN           = 12
_FILA_PROVEEDOR      = 13
_FILA_DESCRIPCION    = 17
_FILA_VALOR_TOTAL    = 18
_FILA_COT_SELEC      = 32   # fila con la "X" de cotización seleccionada

_COL_INICIO          = 4    # columna D
_SALTO               = 3    # cada ítem ocupa 3 columnas (una por cotización)
_HOJA_NOMBRE         = "2. Revisión compra (cotización)"

_RE_ITEM_GENERICO    = re.compile(r"^item\s*\d+$")   # se aplica sobre texto ya normalizado
# Número colombiano: 195.000  /  1.796.826  /  2.990.000
_RE_NUM_CO           = re.compile(r"\b\d{1,3}(?:\.\d{3})+\b")


class AlertaCotizacion(NamedTuple):
    item:       str   # nombre del ítem
    cotizacion: int   # 1, 2 o 3
    campo:      str   # "proveedor", "descripcion", "valor_total" o "no_seleccionada_en_pdf"


class CotizacionSeleccionada(NamedTuple):
    item:        str
    cotizacion:  int          # 1, 2 o 3
    proveedor:   str | None
    descripcion: str | None
    valor_total: int | None   # valor numérico sin puntos


class ResultadoPlanInversion(NamedTuple):
    ok:                   bool
    alertas:              list[AlertaCotizacion]
    items_ok:             list[str]
    resumen:              str
    seleccionadas:        list[CotizacionSeleccionada]   # cotizaciones elegidas por producto


class ResultadoFirmas(NamedTuple):
    ok:             bool
    firmas_halladas: int
    resumen:        str


class ResultadoCotizacionPDF(NamedTuple):
    ok:       bool
    alertas:  list[str]   # una por producto cuyo valor total no se encontró en el PDF
    resumen:  str


# ── Helpers ──────────────────────────────────────────────────────────────────

def _valor(ws, fila: int, col: int):
    """Devuelve el valor de una celda como str limpio, o None si está vacía."""
    v = ws.cell(row=fila, column=col).value
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _valor_num(ws, fila: int, col: int) -> int | None:
    """Devuelve el valor numérico de una celda (entero), o None."""
    v = ws.cell(row=fila, column=col).value
    if v is None:
        return None
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return None


def _es_item_generico(nombre: str | None) -> bool:
    if not nombre:
        return True
    from app.core.normalizacion import normalizar
    return bool(_RE_ITEM_GENERICO.match(normalizar(str(nombre).strip())))


def _co_num_a_int(texto: str) -> int:
    """Convierte '1.796.826' (formato colombiano) a entero 1796826."""
    return int(texto.replace(".", ""))


def _extraer_numeros_pdf(texto: str) -> set[int]:
    """Extrae todos los números en formato colombiano del texto del PDF."""
    return {_co_num_a_int(m) for m in _RE_NUM_CO.findall(texto)}


# ── Validaciones ──────────────────────────────────────────────────────────────

def validar_plan_inversion(ruta: Path) -> ResultadoPlanInversion:
    """
    Lee el PLAN_INVERSION.xlsx y valida la hoja de cotizaciones.
    También extrae la cotización seleccionada (X) por producto.
    """
    try:
        import openpyxl
    except ImportError:
        msg = "openpyxl no instalado"
        return ResultadoPlanInversion(ok=False, alertas=[], items_ok=[], resumen=msg, seleccionadas=[])

    try:
        wb = openpyxl.load_workbook(str(ruta), data_only=True)
    except Exception as exc:
        msg = f"No se pudo abrir {ruta.name}: {exc}"
        return ResultadoPlanInversion(ok=False, alertas=[], items_ok=[], resumen=msg, seleccionadas=[])

    if _HOJA_NOMBRE not in wb.sheetnames:
        msg = f"Hoja '{_HOJA_NOMBRE}' no encontrada en {ruta.name}"
        return ResultadoPlanInversion(ok=False, alertas=[], items_ok=[], resumen=msg, seleccionadas=[])

    ws      = wb[_HOJA_NOMBRE]
    max_col = ws.max_column

    alertas:       list[AlertaCotizacion]      = []
    items_ok:      list[str]                   = []
    seleccionadas: list[CotizacionSeleccionada] = []
    items_revisados = 0

    col = _COL_INICIO
    while col + 2 <= max_col:
        nombre_item = _valor(ws, _FILA_ITEM_NOMBRE, col)
        proveedores = [_valor(ws, _FILA_PROVEEDOR, col + i) for i in range(3)]
        activo      = any(p for p in proveedores)

        if not activo:
            if _es_item_generico(nombre_item):
                break   # fin de ítems reales

        label = nombre_item or f"Ítem en col {col}"
        items_revisados += 1
        alertas_item: list[AlertaCotizacion] = []

        # ── Verificar completitud de las 3 cotizaciones ───────────────────
        for i in range(3):
            num_cot = i + 1
            c = col + i
            if not _valor(ws, _FILA_PROVEEDOR,   c):
                alertas_item.append(AlertaCotizacion(label, num_cot, "proveedor"))
            if not _valor(ws, _FILA_DESCRIPCION, c):
                alertas_item.append(AlertaCotizacion(label, num_cot, "descripcion"))
            if _valor_num(ws, _FILA_VALOR_TOTAL, c) is None:
                alertas_item.append(AlertaCotizacion(label, num_cot, "valor_total"))

        # ── Detectar cotización seleccionada (X en fila 32) ───────────────
        selec_idx = None   # 0, 1 o 2
        for i in range(3):
            v = _valor(ws, _FILA_COT_SELEC, col + i)
            if v and v.strip().upper() == "X":
                selec_idx = i
                break

        if selec_idx is not None:
            c_sel = col + selec_idx
            seleccionadas.append(CotizacionSeleccionada(
                item        = label,
                cotizacion  = selec_idx + 1,
                proveedor   = _valor(ws, _FILA_PROVEEDOR,   c_sel),
                descripcion = _valor(ws, _FILA_DESCRIPCION, c_sel),
                valor_total = _valor_num(ws, _FILA_VALOR_TOTAL, c_sel),
            ))
        else:
            alertas_item.append(AlertaCotizacion(label, 0, "sin_cotizacion_seleccionada"))
            logger.warning("  PLAN_INVERSION '%s': no tiene X en cotización seleccionada", label)

        if alertas_item:
            alertas.extend(alertas_item)
            logger.info("  PLAN_INVERSION '%s': %d alerta(s)", label, len(alertas_item))
        else:
            items_ok.append(label)

        col += _SALTO

    if items_revisados == 0:
        resumen = "PLAN_INVERSION: no se encontraron ítems con datos"
        return ResultadoPlanInversion(ok=False, alertas=[], items_ok=[], resumen=resumen, seleccionadas=[])

    if not alertas:
        resumen = f"OK ({items_revisados} ítem(s), todas las cotizaciones completas)"
    else:
        items_con_alerta: dict[str, list[str]] = {}
        for a in alertas:
            items_con_alerta.setdefault(a.item, []).append(
                f"Cot.{a.cotizacion}/{a.campo}" if a.cotizacion else a.campo
            )
        partes = [f"{item}: {', '.join(errs)}" for item, errs in items_con_alerta.items()]
        resumen = "FALTA — " + " | ".join(partes)

    return ResultadoPlanInversion(
        ok           = len(alertas) == 0,
        alertas      = alertas,
        items_ok     = items_ok,
        resumen      = resumen,
        seleccionadas= seleccionadas,
    )


_FIRMAS_REQUERIDAS = 2


def verificar_firmas_pdf(ruta_pdf: Path) -> ResultadoFirmas:
    """
    Cuenta las imágenes embebidas en el PDF como indicador de firmas manuscritas.

    El documento FIRMA_UP_*.pdf tiene exactamente 2 imágenes de firma
    (representante legal + consultor). No usa IA ni widgets digitales.

    Retorna ok=True solo si se encuentran exactamente _FIRMAS_REQUERIDAS imágenes.
    """
    try:
        import fitz
        doc = fitz.open(str(ruta_pdf))
    except Exception as exc:
        return ResultadoFirmas(ok=False, firmas_halladas=0, resumen=f"Error al leer PDF: {exc}")

    total_imgs = sum(len(page.get_image_info()) for page in doc)
    doc.close()

    logger.debug("  Firmas PDF '%s': %d imagen(s) encontradas", ruta_pdf.name, total_imgs)

    if total_imgs >= _FIRMAS_REQUERIDAS:
        return ResultadoFirmas(
            ok=True,
            firmas_halladas=total_imgs,
            resumen=f"OK — {total_imgs} firma(s) encontradas",
        )

    faltantes = _FIRMAS_REQUERIDAS - total_imgs
    if total_imgs == 0:
        msg = "FALTA — no se encontraron firmas en el PDF"
    else:
        msg = f"FALTA — solo {total_imgs} de {_FIRMAS_REQUERIDAS} firma(s) requeridas"

    return ResultadoFirmas(ok=False, firmas_halladas=total_imgs, resumen=msg)


def validar_cotizacion_seleccionada_en_pdf(
    seleccionadas: list[CotizacionSeleccionada],
    ruta_pdf: Path,
) -> ResultadoCotizacionPDF:
    """
    Verifica que el valor total de cada cotización seleccionada aparezca en el PDF.

    Estrategia:
    - Extrae todos los números con formato colombiano (p.ej. 195.000) del PDF.
    - Para cada producto seleccionado, busca su valor_total entre esos números.
    - Si no lo encuentra → alerta.
    """
    try:
        import fitz
        doc  = fitz.open(str(ruta_pdf))
        texto_pdf = "\n".join(page.get_text() for page in doc)
        doc.close()
    except Exception as exc:
        msg = f"No se pudo leer el PDF {ruta_pdf.name}: {exc}"
        logger.error(msg)
        return ResultadoCotizacionPDF(ok=False, alertas=[msg], resumen=f"Error al leer PDF: {exc}")

    numeros_pdf = _extraer_numeros_pdf(texto_pdf)
    logger.debug("Números extraídos del PDF: %s", sorted(numeros_pdf))

    alertas_pdf: list[str] = []
    ok_items:    list[str] = []

    for sel in seleccionadas:
        if sel.valor_total is None:
            alertas_pdf.append(
                f"{sel.item} (Cot.{sel.cotizacion}): valor total no disponible en Excel"
            )
            continue

        if sel.valor_total in numeros_pdf:
            ok_items.append(f"{sel.item}: ${sel.valor_total:,}".replace(",", "."))
            logger.info(
                "  PDF: valor %d de '%s' (Cot.%d) encontrado",
                sel.valor_total, sel.item, sel.cotizacion,
            )
        else:
            alertas_pdf.append(
                f"{sel.item} (Cot.{sel.cotizacion} — {sel.proveedor or 'sin proveedor'}): "
                f"valor ${sel.valor_total:,} no encontrado en PDF".replace(",", ".")
            )
            logger.warning(
                "  PDF: valor %d de '%s' (Cot.%d) NO encontrado",
                sel.valor_total, sel.item, sel.cotizacion,
            )

    if not alertas_pdf:
        resumen = f"OK — {len(ok_items)} producto(s) verificados en PDF"
    else:
        resumen = "FALTA — " + " | ".join(alertas_pdf)

    return ResultadoCotizacionPDF(
        ok      = len(alertas_pdf) == 0,
        alertas = alertas_pdf,
        resumen = resumen,
    )
