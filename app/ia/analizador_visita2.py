"""
Analizadores de documentos de la carpeta
02_VISITA_2_DIAGNOSTICO.

- Plan de Negocio
- Diagnóstico
"""

import logging
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

CAMPOS_CONTROL = [
    "Nombre Unidad Productiva",
    "Nombre representante",
    "Celular",
    "Actividad económica principal",
    "Necesidad o problema que atiende",
    "Principales clientes",
    "Clientes a los que quiere llegar",
    "Acción de alianza",
    "Fortaleza",
    "Oportunidad",
    "Debilidad",
    "Amenaza",
    "Objetivo del plan",
    "Acción administrativa",
    "Acción productiva",
    "Tema asistencia técnica",
    "Modalidad educativa",
    "Pertenece a organización",
    "Interés en participar",
]


def buscar_valor_derecha(ws, etiqueta):
    """
    Busca una etiqueta dentro de la hoja y retorna
    el primer valor encontrado.

    Estrategia:
    1. Buscar la etiqueta normalizada.
    2. Buscar contenido a la derecha.
    3. Si no existe, buscar contenido debajo.
    """

    etiqueta_norm = normalizar(etiqueta)

    for fila in ws.iter_rows():

        for celda in fila:

            if celda.value is None:
                continue

            texto = normalizar(str(celda.value))

            if etiqueta_norm in texto:

                fila_idx = celda.row
                col_idx = celda.column

                # ── Buscar hacia la derecha ──────────────────────────────
                for c in range(col_idx + 1, col_idx + 6):

                    valor = ws.cell(
                        row=fila_idx,
                        column=c,
                    ).value

                    if valor is None:
                        continue

                    if isinstance(valor, str) and not valor.strip():
                        continue

                    return valor

                # ── Buscar hacia abajo ──────────────────────────────────
                for r in range(fila_idx + 1, fila_idx + 4):

                    valor = ws.cell(
                        row=r,
                        column=col_idx,
                    ).value

                    if valor is None:
                        continue

                    if isinstance(valor, str) and not valor.strip():
                        continue

                    return valor

                return None

    return None

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

        for etiqueta in CAMPOS_CONTROL:
            valor = buscar_valor_derecha(
                ws,
                etiqueta,
            )
            if valor is None:
                faltantes.append(
                    etiqueta
                )

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