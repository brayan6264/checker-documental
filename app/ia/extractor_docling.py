from pathlib import Path

from docling.document_converter import DocumentConverter

_converter = DocumentConverter()


def extraer_texto_docling(ruta_pdf: Path) -> str:
    """
    Extrae el texto de un PDF utilizando Docling.
    """
    try:
        resultado = _converter.convert(str(ruta_pdf))
        return resultado.document.export_to_text().strip()

    except Exception:
        return ""