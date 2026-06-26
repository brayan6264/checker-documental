"""
Selección del mejor resultado de extracción de texto.

Responsabilidad:
    Comparar el texto obtenido por distintos motores OCR y seleccionar
    el que tenga mayor calidad.

Versión 1:
    - Prioriza cantidad de palabras.
    - Prioriza cantidad de números.
    - Penaliza textos muy cortos.
"""

from __future__ import annotations

import logging
import re

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------
# Expresiones regulares
# ---------------------------------------------------------------------

_RE_PALABRAS = re.compile(r"\b[\wÁÉÍÓÚÜÑáéíóúüñ]{2,}\b")
_RE_NUMEROS = re.compile(r"\d+")


# ---------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------

def seleccionar_mejor_texto(
    texto_easyocr: str,
    texto_docling: str,
) -> str:
    """
    Devuelve el mejor texto entre EasyOCR y Docling.

    Si ambos están vacíos devuelve "".
    """

    score_easy = calcular_score(texto_easyocr)
    score_doc = calcular_score(texto_docling)

    logger.info(
        "Score OCR -> EasyOCR=%d | Docling=%d",
        score_easy,
        score_doc,
    )

    if score_easy >= score_doc:
        return texto_easyocr, "EasyOCR"

    return texto_docling, "Docling"


# ---------------------------------------------------------------------
# Calidad del texto
# ---------------------------------------------------------------------

def calcular_score(texto: str) -> int:
    """
    Calcula una puntuación sencilla de calidad.

    La fórmula da más importancia a:

    - palabras
    - números

    y penaliza textos demasiado cortos.
    """

    if not texto:
        return 0

    palabras = contar_palabras(texto)
    numeros = contar_numeros(texto)
    longitud = len(texto)

    score = (
        palabras * 2
        + numeros * 5
        + longitud // 100
    )

    if longitud < 100:
        score -= 50

    return score


# ---------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------

def contar_palabras(texto: str) -> int:
    return len(_RE_PALABRAS.findall(texto))


def contar_numeros(texto: str) -> int:
    return len(_RE_NUMEROS.findall(texto))