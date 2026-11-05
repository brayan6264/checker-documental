"""
Analizadores de documentos de la carpeta
02_VISITA_2_DIAGNOSTICO.

- Plan de Negocio
- Diagnóstico
"""

import logging
import re
from pathlib import Path
from typing import Dict, List
from app.core.normalizacion import normalizar

import openpyxl

from app.config import (
    IA_HABILITADO,
    IA_MAX_PAGINAS_VISITA,
    OPENAI_API_KEY,
)

from app.ia.analizador_visita import (
    ResultadoAnalisis,
    _SISTEMA_BASE,
    _llamar_gpt,
    _obtener_imagenes_b64,
    _parsear_json,
)

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────────────────────────
# PLAN DE NEGOCIO
# ──────────────────────────────────────────────────────────────────────────────

# Dónde está la respuesta de cada etiqueta en el formato.
_A_LA_DERECHA = "derecha"   # misma fila; la etiqueta ocupa hasta 2 columnas
_DEBAJO       = "debajo"    # fila siguiente, misma columna
_LISTA_FODA   = "lista"     # filas numeradas debajo: "1.1" + respuesta
_TABLA        = "tabla"     # filas debajo, misma columna, hasta la sección siguiente

_COLUMNAS_ETIQUETA = 2
_FILAS_LISTA_FODA  = 4
_FILAS_TABLA       = 8
_RE_NUMERACION     = re.compile(r"^\d+(\.\d+)*\.?$")
_RE_FIN_TABLA      = re.compile(r"\d+\.\s*procesos asociativos")
_RE_INSTRUCCION    = re.compile(r"agregar las filas")

# (campo, etiqueta al inicio de la celda, ubicación de la respuesta).
# Las etiquetas se anclan al inicio para no confundirlas con respuestas de
# texto libre que mencionan las mismas palabras. "decripcion", "amenzas" y
# "pertence" son errores de digitación presentes en versiones reales del formato.
CAMPOS_PLAN_NEGOCIO = [
    ("unidad productiva",             r"nombre de la unidad productiva",             _A_LA_DERECHA),
    ("representante",                 r"nombre del representante",                   _A_LA_DERECHA),
    ("celular",                       r"celular",                                    _A_LA_DERECHA),
    ("actividad economica principal", r"actividad economica principal",              _A_LA_DERECHA),
    ("necesidad o problema",          r"necesidad o problema",                       _A_LA_DERECHA),
    ("principales clientes",          r"des?cripcion de los principales clientes",   _A_LA_DERECHA),
    ("otro tipo de clientes",         r"a que otro tipo de clientes",                _A_LA_DERECHA),
    ("fortalezas",                    r"\d+\.\s*fortalezas",                         _LISTA_FODA),
    ("oportunidades",                 r"\d+\.\s*oportunidades",                      _LISTA_FODA),
    ("debilidades",                   r"\d+\.\s*debilidades",                        _LISTA_FODA),
    ("amenazas",                      r"\d+\.\s*(amenazas|amenzas)",                 _LISTA_FODA),
    ("plan de inversion",             r"objetivo del plan de (inversion|accion)",    _DEBAJO),
    ("estrategias administrativas",   r"estrategias? administrativas?",              _A_LA_DERECHA),
    ("fortalecimiento productivo",    r"estrategias? de fortalecimiento productivo", _A_LA_DERECHA),
    ("asistencia tecnica",            r"temas de asistencia tecnica",                _TABLA),
    ("modalidad educativa",           r"modalidad educativa",                        _TABLA),
    ("organizacion y/o asociacion",   r"pert\w+ (a )?algun tipo de organizacion",    _A_LA_DERECHA),
    ("asociacion, red o plataforma",  r"conoce alguna asociacion, red o plataforma", _A_LA_DERECHA),
]

_RE_ETIQUETA_FODA = re.compile(r"\d+\.\s*(fortalezas|oportunidades|debilidades|amenazas|amenzas)")


def _texto_celda(ws, fila: int, columna: int) -> str:
    """Valor de la celda como texto sin espacios extremos ('' si está vacía)."""
    valor = ws.cell(row=fila, column=columna).value
    return "" if valor is None else str(valor).strip()


def _texto_etiqueta(valor) -> str:
    """Normaliza una posible etiqueta: sin tildes, signos iniciales ni dobles espacios."""
    return " ".join(normalizar(str(valor)).lstrip("¿?¡! ").split())


def _buscar_etiqueta(ws, patron: str):
    """Retorna la primera celda cuyo texto inicia con la etiqueta, o None."""
    regex = re.compile(patron)

    for fila in ws.iter_rows():
        for celda in fila:
            if celda.value is not None and regex.match(_texto_etiqueta(celda.value)):
                return celda

    return None


def _lista_foda_tiene_respuesta(ws, fila: int, columna: int) -> bool:
    """True si alguna fila numerada bajo el título del FODA tiene respuesta."""
    for r in range(fila + 1, fila + 1 + _FILAS_LISTA_FODA):
        if _RE_ETIQUETA_FODA.match(_texto_etiqueta(_texto_celda(ws, r, columna))):
            return False

        for c in (columna, columna + 1):
            texto = _texto_celda(ws, r, c)
            if texto and not _RE_NUMERACION.match(texto):
                return True

    return False


def _tabla_tiene_respuesta(ws, fila: int, columna: int) -> bool:
    """True si la columna tiene una respuesta real antes de la sección siguiente."""
    for r in range(fila + 1, fila + 1 + _FILAS_TABLA):
        if any(
            _RE_FIN_TABLA.match(_texto_etiqueta(_texto_celda(ws, r, c)))
            for c in range(1, columna + 1)
        ):
            return False

        texto = _texto_etiqueta(_texto_celda(ws, r, columna))
        if texto and not _RE_INSTRUCCION.match(texto):
            return True

    return False


def campo_diligenciado(ws, patron: str, ubicacion: str) -> bool:
    """True si la etiqueta existe en la hoja y su propia respuesta tiene contenido."""
    celda = _buscar_etiqueta(ws, patron)

    if celda is None:
        return False

    if ubicacion == _LISTA_FODA:
        return _lista_foda_tiene_respuesta(ws, celda.row, celda.column)

    if ubicacion == _TABLA:
        return _tabla_tiene_respuesta(ws, celda.row, celda.column)

    if ubicacion == _DEBAJO:
        return bool(_texto_celda(ws, celda.row + 1, celda.column))

    return any(
        _texto_celda(ws, celda.row, celda.column + desplazamiento)
        for desplazamiento in range(1, _COLUMNAS_ETIQUETA + 1)
    )


def _hoja_plan_negocio(wb):
    """Retorna la hoja con más etiquetas del formato (no siempre es la primera)."""
    return max(
        wb.worksheets,
        key=lambda ws: sum(
            _buscar_etiqueta(ws, patron) is not None
            for _, patron, _ in CAMPOS_PLAN_NEGOCIO
        ),
    )


def analizar_plan_negocio(ruta_excel: str) -> Dict[str, object]:
    """
    Valida que los campos mínimos del plan de negocio estén diligenciados.

    Si el archivo no se puede abrir retorna "error" con el motivo; en ese caso
    "faltantes" queda vacío porque el contenido no fue revisado.
    """

    try:
        wb = openpyxl.load_workbook(
            ruta_excel,
            data_only=True,
        )

        ws = _hoja_plan_negocio(wb)

        faltantes: List[str] = [
            campo
            for campo, patron, ubicacion in CAMPOS_PLAN_NEGOCIO
            if not campo_diligenciado(ws, patron, ubicacion)
        ]

        wb.close()

        return {
            "completo": len(faltantes) == 0,
            "faltantes": faltantes,
        }

    except Exception as exc:
        logger.warning("Plan de negocio ilegible '%s': %s", ruta_excel, exc)

        return {
            "completo": False,
            "faltantes": [],
            "error": str(exc),
        }


# ──────────────────────────────────────────────────────────────────────────────
# DIAGNÓSTICO
# ──────────────────────────────────────────────────────────────────────────────

def analizar_diagnostico(
    ruta_pdf: str | Path,
) -> ResultadoAnalisis:
    """
    Verifica que todas las secciones principales del
    Diagnóstico DPS estén diligenciadas.
    """

    if not IA_HABILITADO or not OPENAI_API_KEY:
        return ResultadoAnalisis(ok=True, alerta="")

    ruta = Path(ruta_pdf)

    imagenes = _obtener_imagenes_b64(
        ruta,
        IA_MAX_PAGINAS_VISITA,
    )

    if not imagenes:
        return ResultadoAnalisis(
            ok=False,
            alerta="No se pudo leer el archivo para análisis IA",
        )

    prompt_usuario = """Analiza este formato DIAGNÓSTICO DE UNIDADES PRODUCTIVAS DPS escaneado.

Verifica los siguientes puntos:

1. ¿Todos los campos OBLIGATORIOS del formulario tienen contenido?
    - Un campo obligatorio está vacío SOLO si está completamente en blanco.

2. Para preguntas con opciones SI / NO:
   - Debe existir al menos una opción marcada.

3. Para preguntas que incluyan observaciones, descripción,
   comentarios o justificaciones:
   - Debe existir contenido escrito.

4. Para espacios de firma:
   - Cualquier rúbrica, garabato o marca visible cuenta como firma.

5. Ignora campos que indiquen:
   - si aplica
   - cuando aplique
   - opcional
   - no aplica
   - o cualquier variante equivalente

Responde SOLO con este JSON:

{
  "campos_completos": true/false,
  "campos_vacios": [
    "lista de campos obligatorios completamente vacíos"
  ],
  "alerta": "descripción breve del problema o null"
}
"""

    try:

        data = _parsear_json(
            _llamar_gpt(
                imagenes,
                _SISTEMA_BASE,
                prompt_usuario,
            )
        )

    except Exception as exc:

        logger.error(
            "Error IA diagnóstico '%s': %s",
            ruta.name,
            exc,
        )

        return ResultadoAnalisis(
            ok=False,
            alerta=f"Error en análisis IA: {exc}",
        )

    if not data.get("campos_completos", True):

        campos_vacios = data.get(
            "campos_vacios",
            [],
        )

        if campos_vacios:

            return ResultadoAnalisis(
                ok=False,
                alerta=(
                    "Campos vacíos: "
                    + ", ".join(campos_vacios[:5])
                ),
            )

        return ResultadoAnalisis(
            ok=False,
            alerta=data.get("alerta")
            or "Diagnóstico incompleto",
        )

    return ResultadoAnalisis(ok=True)

