"""
Extracción de contenido de documentos para análisis con IA.

Responsabilidad: dada la ruta de un archivo, devolver su texto/contenido
en una forma que pueda ser enviada a un LLM para validación semántica.

Formatos soportados actualmente: PDF, imágenes (JPG, PNG, TIFF).
Extensible a DOCX, XLSX, etc.

Dependencias opcionales (instalar según necesidad):
  pip install pymupdf pillow          # PDF e imágenes
  pip install python-docx             # DOCX
"""

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# Tipos de archivo soportados por extractor
_EXTENSIONES_PDF    = {".pdf"}
_EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"}


def extraer_texto(ruta: Path) -> str:
    """
    Extrae el texto de un documento.

    Returns:
        Texto plano del documento, o cadena vacía si no se pudo extraer.
    """
    ext = ruta.suffix.lower()

    if ext in _EXTENSIONES_PDF:
        return _extraer_pdf(ruta)
    if ext in _EXTENSIONES_IMAGEN:
        return _extraer_imagen(ruta)

    logger.debug("Formato no soportado para extracción: %s", ext)
    return ""


def extraer_imagen_base64(ruta: Path) -> str | None:
    """
    Codifica una imagen en base64 para enviarla como contenido visual a un LLM.

    Returns:
        String base64 de la imagen, o None si no aplica/falla.
    """
    import base64

    ext = ruta.suffix.lower()
    if ext not in _EXTENSIONES_IMAGEN:
        return None

    try:
        with open(ruta, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as exc:
        logger.warning("No se pudo codificar imagen '%s': %s", ruta, exc)
        return None


# ── Extractores por tipo ──────────────────────────────────────────────────────

def _extraer_pdf(ruta: Path) -> str:
    """Extrae texto de un PDF con PyMuPDF (fitz). Incluye texto de imágenes si hay OCR."""
    try:
        import fitz  # pymupdf

        doc   = fitz.open(str(ruta))
        texto = "\n".join(pagina.get_text() for pagina in doc)
        doc.close()
        return texto.strip()
    except ImportError:
        logger.warning("PyMuPDF no instalado. Instalar con: pip install pymupdf")
        return ""
    except Exception as exc:
        logger.warning("Error extrayendo PDF '%s': %s", ruta, exc)
        return ""


def _extraer_imagen(ruta: Path) -> str:
    """
    Para imágenes, devuelve cadena vacía — el análisis visual se hace
    directamente desde extraer_imagen_base64() enviando la imagen al LLM.
    Si se necesita OCR local, integrar pytesseract aquí.
    """
    return ""
