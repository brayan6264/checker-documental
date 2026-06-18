"""
Extracción de contenido de documentos para análisis con IA.

Responsabilidad: dada la ruta de un archivo, devolver su texto/contenido
en una forma que pueda ser enviada a un LLM para validación semántica.

Formatos soportados: PDF (texto nativo + OCR para escaneados), imágenes.

Dependencias:
  pip install pymupdf pillow pytesseract
  # Tesseract binario: C:\Program Files\Tesseract-OCR\tesseract.exe
  # Idioma español:    C:\Program Files\Tesseract-OCR\tessdata\spa.traineddata
"""

import base64
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_EXTENSIONES_PDF    = {".pdf"}
_EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"}

# Si el PDF tiene menos de este promedio de caracteres por página se considera escaneado
_CHARS_POR_PAGINA_UMBRAL = 50

# Si el OCR devuelve menos de este promedio de chars/página se considera fallido
_CHARS_OCR_MINIMO = 80

# Ruta del ejecutable Tesseract en Windows
_TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def extraer_texto(ruta: Path) -> str:
    """
    Extrae el texto de un documento.
    Para PDFs escaneados aplica OCR automáticamente.
    Devuelve cadena vacía si no se pudo extraer.
    """
    ext = ruta.suffix.lower()

    if ext in _EXTENSIONES_PDF:
        return _extraer_pdf(ruta)
    if ext in _EXTENSIONES_IMAGEN:
        return _extraer_imagen_ocr(ruta)

    logger.debug("Formato no soportado para extracción: %s", ext)
    return ""


def extraer_imagen_base64(ruta: Path) -> str | None:
    """
    Codifica una imagen en base64 para enviarla como contenido visual a un LLM.
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


# ── Extractores ───────────────────────────────────────────────────────────────

def _extraer_pdf(ruta: Path) -> str:
    """
    Extrae texto de un PDF.
    1. Intenta extracción nativa con PyMuPDF.
    2. Si el resultado es escaso (PDF escaneado), aplica OCR con Tesseract.
    """
    try:
        import fitz
    except ImportError:
        logger.warning("PyMuPDF no instalado (pip install pymupdf)")
        return ""

    try:
        doc = fitz.open(str(ruta))
    except Exception as exc:
        logger.warning("No se pudo abrir PDF '%s': %s", ruta.name, exc)
        return ""

    paginas_texto = [p.get_text() for p in doc]
    total_chars   = sum(len(t.strip()) for t in paginas_texto)
    promedio      = total_chars / max(len(doc), 1)

    if promedio >= _CHARS_POR_PAGINA_UMBRAL:
        # PDF con texto nativo — no necesita OCR
        doc.close()
        return "\n".join(paginas_texto).strip()

    # PDF escaneado — aplicar OCR página por página
    logger.info(
        "PDF '%s' escaneado (%.0f chars/pág). Aplicando OCR...",
        ruta.name, promedio,
    )
    texto_ocr = _ocr_pdf(doc, ruta)
    doc.close()

    # Evaluar si el OCR fue suficiente
    chars_ocr = len(texto_ocr.strip())
    if chars_ocr >= _CHARS_OCR_MINIMO:
        return texto_ocr

    # OCR insuficiente — fallback a GPT-4o-mini con el PDF completo
    logger.warning(
        "PDF '%s': OCR insuficiente (%d chars). Enviando a GPT-4o-mini...",
        ruta.name, chars_ocr,
    )
    texto_gpt = _extraer_con_gpt(ruta)
    if texto_gpt:
        return texto_gpt

    # Último recurso: devolver lo poco que haya
    return texto_ocr or "\n".join(paginas_texto).strip()


def _ocr_pdf(doc, ruta: Path) -> str:
    """
    Renderiza cada página del PDF como imagen y aplica Tesseract OCR.
    Usa 300 DPI para buena precisión en documentos escaneados.
    """
    try:
        import pytesseract
        from PIL import Image
        import io
    except ImportError:
        logger.warning("pytesseract/Pillow no instalados (pip install pytesseract pillow). Sin OCR.")
        return ""

    # Apuntar al ejecutable en Windows si no está en el PATH
    import os
    if os.path.exists(_TESSERACT_CMD):
        pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD

    partes: list[str] = []
    for i, pagina in enumerate(doc):
        try:
            pix = pagina.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            # spa+eng: español primero, inglés como respaldo para números y siglas
            texto = pytesseract.image_to_string(img, lang="spa+eng", config="--psm 1")
            partes.append(texto)
            logger.debug("  OCR pág %d/%d de '%s': %d chars", i + 1, len(doc), ruta.name, len(texto))
        except Exception as exc:
            logger.warning("  OCR error pág %d de '%s': %s", i + 1, ruta.name, exc)

    resultado = "\n".join(partes).strip()
    logger.info("  OCR completado '%s': %d chars totales", ruta.name, len(resultado))
    return resultado


def _extraer_con_gpt(ruta: Path) -> str:
    """
    Fallback cuando el OCR de Tesseract es insuficiente en un PDF escaneado.
    Envía el PDF completo a gpt-4o-mini vía OpenAI Responses API (PDF inline en base64)
    y pide que transcriba todo el texto visible.
    Solo se activa para PDFs escaneados con OCR fallido; no se llama para PDFs con texto nativo.
    """
    try:
        from openai import OpenAI
    except ImportError:
        logger.warning("openai no instalado — no se puede usar fallback GPT para OCR.")
        return ""

    try:
        from app.config import OPENAI_API_KEY
    except Exception:
        OPENAI_API_KEY = ""

    if not OPENAI_API_KEY:
        logger.warning("OPENAI_API_KEY no configurada — no se puede usar fallback GPT para OCR.")
        return ""

    try:
        pdf_b64 = base64.b64encode(ruta.read_bytes()).decode()
        client  = OpenAI(api_key=OPENAI_API_KEY)

        response = client.responses.create(
            model="gpt-4o-mini",
            input=[{
                "role": "user",
                "content": [
                    {
                        "type": "input_file",
                        "filename": ruta.name,
                        "file_data": f"data:application/pdf;base64,{pdf_b64}",
                    },
                    {
                        "type": "input_text",
                        "text": (
                            "Transcribe TODO el texto visible en este documento exactamente "
                            "como aparece, incluyendo nombres, fechas, números, direcciones y "
                            "cualquier otro dato. No resumas ni interpretes; solo transcribe el "
                            "contenido completo del documento."
                        ),
                    },
                ],
            }],
        )
        texto = (response.output_text or "").strip()
        logger.info(
            "  GPT-4o-mini OCR fallback '%s': %d chars extraídos.", ruta.name, len(texto)
        )
        return texto

    except Exception as exc:
        logger.error("  GPT-4o-mini OCR fallback error '%s': %s", ruta.name, exc)
        return ""


def _extraer_imagen_ocr(ruta: Path) -> str:
    """
    Aplica OCR a una imagen directamente con Tesseract.
    """
    try:
        import pytesseract
        from PIL import Image
        import os

        if os.path.exists(_TESSERACT_CMD):
            pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD

        img   = Image.open(str(ruta))
        texto = pytesseract.image_to_string(img, lang="spa+eng", config="--psm 1")
        return texto.strip()
    except ImportError:
        logger.warning("pytesseract/Pillow no instalados. Sin OCR para imágenes.")
        return ""
    except Exception as exc:
        logger.warning("OCR imagen '%s': %s", ruta.name, exc)
        return ""
