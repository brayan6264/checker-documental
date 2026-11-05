"""Utilidades de normalización de texto para comparación insensible a tildes/mayúsculas."""

import unicodedata


def normalizar(texto: str) -> str:
    """
    Normaliza un texto para comparación robusta: sin tildes, sin mayúsculas, sin espacios extremos.

    Pasos:
      1. NFKD descompone caracteres combinados (é → e + ́).
      2. Se eliminan las marcas diacríticas (combining characters).
      3. casefold() convierte a minúsculas (más exhaustivo que lower() para Unicode).
      4. strip() quita espacios al inicio/final.

    Ejemplos:
      "CÉDULA"            → "cedula"
      "00_DOCUMENTACIÓN"  → "00_documentacion"
      "Cámara de Comercio"→ "camara de comercio"
    """
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    sin_diacriticos = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sin_diacriticos.casefold().strip()


def coincide_nombre(nombre: str, palabras: list[str]) -> bool:
    """
    True si alguna palabra clave aparece en el nombre, ignorando separadores.

    Espacios, "_", "-" y demás signos no cuentan, de modo que "plan de negocio"
    coincide con "PLAN_DE_NEGOCIO_123", "Plan-de-negocio" y "PlanDeNegocio".
    """
    def compactar(texto: str) -> str:
        return "".join(c for c in normalizar(texto) if c.isalnum())

    nombre_compacto = compactar(nombre)
    return any(compactar(palabra) in nombre_compacto for palabra in palabras)
