"""
Extracción de texto mediante EasyOCR.

Responsabilidad:
    Obtener el texto visible de PDFs escaneados e imágenes.

No contiene lógica de selección ni validación.
Únicamente devuelve un string con el texto reconocido.
"""

from __future__ import annotations

import logging
from pathlib import Path

import easyocr
import fitz
import numpy as np

logger = logging.getLogger(__name__)

# -----------------------------------------------------------------------------
# Singleton del lector OCR
# -----------------------------------------------------------------------------

_reader = easyocr.Reader(
    ["es", "en"],
    gpu=False,
)


# -----------------------------------------------------------------------------
# API pública
# -----------------------------------------------------------------------------

def extraer_texto_easyocr(ruta: Path) -> str:
    """
    Extrae texto usando EasyOCR.

    Soporta:

    - PDF
    - JPG
    - PNG
    - TIFF
    - BMP

    Devuelve:
        str
    """

    ext = ruta.suffix.lower()

    if ext == ".pdf":
        return _ocr_pdf(ruta)

    return _ocr_imagen(ruta)


# -----------------------------------------------------------------------------
# PDF
# -----------------------------------------------------------------------------

def _ocr_pdf(ruta: Path) -> str:
    """
    Convierte cada página del PDF en imagen y aplica EasyOCR.
    """

    texto_paginas: list[str] = []

    doc = fitz.open(ruta)

    try:
        for pagina in doc:

            # 300 DPI aprox.
            pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))

            img = np.frombuffer(
                pix.samples,
                dtype=np.uint8,
            ).reshape(
                pix.height,
                pix.width,
                pix.n,
            )

            resultado = _reader.readtext(
                img,
                detail=0,
                paragraph=True,
            )

            texto_paginas.append("\n".join(resultado))

    finally:
        doc.close()

    texto = "\n\n".join(texto_paginas).strip()

    logger.info(
        "EasyOCR '%s': %d caracteres",
        ruta.name,
        len(texto),
    )

    return texto


# -----------------------------------------------------------------------------
# Imagen
# -----------------------------------------------------------------------------

def _ocr_imagen(ruta: Path) -> str:
    """
    OCR sobre una imagen.
    """

    resultado = _reader.readtext(
        str(ruta),
        detail=0,
        paragraph=True,
    )

    texto = "\n".join(resultado).strip()

    logger.info(
        "EasyOCR '%s': %d caracteres",
        ruta.name,
        len(texto),
    )

    return texto