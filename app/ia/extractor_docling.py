from docling.document_converter import DocumentConverter

_converter = DocumentConverter()

def extraer_texto_docling(ruta_pdf) -> str:
    try:
        result = _converter.convert(str(ruta_pdf))
        return result.document.export_to_text()
    except Exception:
        return ""