"""
Análisis del archivo Plan de Negocio.

Verifica que los campos mínimos de control definidos para el formulario
estén diligenciados en la primera hoja del archivo Excel.

Retorna:
- completo: True/False
- faltantes: lista de campos vacíos
"""

from typing import Dict, List

import openpyxl


# Campos de control definidos para la primera versión
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

    Args:
        ruta_excel:
            Ruta del archivo .xlsx

    Returns:
        {
            "completo": bool,
            "faltantes": [str]
        }
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