"""
Analizadores de documentos de la carpeta
02_VISITA_2_DIAGNOSTICO.

- Plan de Negocio
- Diagnóstico
"""

import logging
from pathlib import Path
from typing import Dict, List

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

CAMPOS_CONTROL = {
    "D11": "Nombre Unidad Productiva",
    "D12": "Nombre representante",
    "D15": "Celular",
    "D17": "Actividad económica principal",
    "D19": "Necesidad o problema que atiende",
    "D22": "Principales clientes",
    "D23": "Clientes a los que quiere llegar",
    "E26": "Acción de alianza",
    "C31": "Fortaleza",
    "E31": "Oportunidad",
    "C36": "Debilidad",
    "E36": "Amenaza",
    "B42": "Objetivo del plan",
    "D44": "Acción administrativa",
    "D46": "Acción productiva",
    "B55": "Tema asistencia técnica",
    "D55": "Modalidad educativa",
    "D60": "Pertenece a organización",
    "D64": "Interés en participar",
}


def analizar_plan_negocio(ruta_excel: str) -> Dict[str, object]:
    """
    Valida que los campos mínimos del plan de negocio estén diligenciados.
    """

    faltantes: List[str] = []

    try:
        wb = openpyxl.load_workbook(
            ruta_excel,
            data_only=True,
        )

        ws = wb.worksheets[0]

        for celda, descripcion in CAMPOS_CONTROL.items():
            valor = ws[celda].value

            if valor is None:
                faltantes.append(f"{celda} - {descripcion}")
                continue

            if isinstance(valor, str) and not valor.strip():
                faltantes.append(f"{celda} - {descripcion}")

        wb.close()

        return {
            "completo": len(faltantes) == 0,
            "faltantes": faltantes,
        }

    except Exception as exc:
        return {
            "completo": False,
            "faltantes": [f"Error al leer archivo: {exc}"],
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