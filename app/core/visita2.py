"""
Validación de la carpeta 02_VISITA_2_DIAGNOSTICO.

Detecta la presencia de los documentos obligatorios mediante el nombre
del archivo, utilizando texto normalizado (sin tildes, minúsculas).
"""

from pathlib import Path

from app.core.normalizacion import normalizar
from app.core.reglas import (
    DOCS_VISITA2_ORDEN,
    PALABRAS_CLAVE_VISITA2,
)


def validar_visita2(ruta_carpeta: str) -> dict:
    """
    Valida la carpeta 02_VISITA_2_DIAGNOSTICO.

    Args:
        ruta_carpeta:
            Ruta local donde se descargó la carpeta.

    Returns:
        {
            "ACTA_VISITA_2": True/False,
            "DIAGNOSTICO": True/False,
            "PLAN_NEGOCIO": True/False,
        }
    """

    presentes = {
        doc: False
        for doc in DOCS_VISITA2_ORDEN
    }

    carpeta = Path(ruta_carpeta)

    if not carpeta.exists():
        return presentes

    for archivo in carpeta.rglob("*"):
        if not archivo.is_file():
            continue

        nombre = normalizar(archivo.stem)

        for documento, palabras in PALABRAS_CLAVE_VISITA2.items():
            if any(palabra in nombre for palabra in palabras):
                presentes[documento] = True

    return presentes