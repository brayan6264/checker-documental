"""
Validación de la carpeta 02_VISITA_2_DIAGNOSTICO.

Detecta la presencia de los documentos obligatorios mediante el nombre
del archivo, utilizando texto normalizado (sin tildes, minúsculas).
"""

from pathlib import Path

from app.core.normalizacion import coincide_nombre
from app.core.reglas import (
    DOCS_VISITA2_ORDEN,
    EXTENSIONES_VISITA2,
    PALABRAS_CLAVE_VISITA2,
)


def clasificar_archivo_visita2(nombre_archivo: str) -> str | None:
    """
    Retorna el documento de la visita 2 que corresponde al archivo, o None.

    Cada archivo se asigna a un solo documento: el primero que coincida en
    PALABRAS_CLAVE_VISITA2 y admita su extensión.
    """
    ruta = Path(nombre_archivo)

    for documento, palabras in PALABRAS_CLAVE_VISITA2.items():
        if not coincide_nombre(ruta.stem, palabras):
            continue

        extensiones = EXTENSIONES_VISITA2.get(documento)
        if extensiones is None or ruta.suffix.lower() in extensiones:
            return documento

        return None

    return None


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

        documento = clasificar_archivo_visita2(archivo.name)

        if documento is not None:
            presentes[documento] = True

    return presentes