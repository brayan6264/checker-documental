"""
Validación semántica de documentos mediante LLM (Anthropic Claude).

Responsabilidad: dado el contenido de un archivo y el tipo de documento esperado,
determinar si el contenido corresponde realmente a ese tipo de documento.

Esto complementa (no reemplaza) la detección por nombre de archivo:
  Fase 1 (ya implementada) → ¿existe un archivo cuyo nombre contiene "cedula"?
  Fase 2 (este módulo)     → ¿el contenido de ese archivo ES realmente una cédula?

Solo se activa si IA_HABILITADO=true en la configuración.
"""

import logging
from pathlib import Path
from typing import Optional

from app.config import ANTHROPIC_API_KEY, IA_HABILITADO, IA_MODEL
from app.ia.extractor import extraer_imagen_base64, extraer_texto

logger = logging.getLogger(__name__)


# ── Prompts por tipo de documento ─────────────────────────────────────────────
#
# Cada prompt le pide al LLM una respuesta estructurada simple: SI/NO + razón breve.
# Se prefiere respuesta corta para minimizar tokens y latencia.

_PROMPTS: dict[str, str] = {
    "CEDULA": (
        "Analiza el siguiente documento. "
        "¿Es una cédula de ciudadanía, tarjeta de identidad o documento de identificación personal? "
        "Responde ÚNICAMENTE con el formato: VALIDO o INVALIDO, seguido de dos puntos y una razón breve. "
        "Ejemplo: 'VALIDO: contiene número de cédula y foto del titular'."
    ),
    "COMERCIO": (
        "Analiza el siguiente documento. "
        "¿Es un certificado de existencia y representación legal, cámara de comercio "
        "o documento que acredita la existencia de una empresa? "
        "Responde ÚNICAMENTE: VALIDO o INVALIDO, seguido de dos puntos y razón breve."
    ),
    "RUT": (
        "Analiza el siguiente documento. "
        "¿Es un Registro Único Tributario (RUT) de la DIAN o documento equivalente "
        "que muestra el NIT de una persona natural o jurídica? "
        "Responde ÚNICAMENTE: VALIDO o INVALIDO, seguido de dos puntos y razón breve."
    ),
    "TENENCIA": (
        "Analiza el siguiente documento. "
        "¿Es una certificación de tenencia, contrato de arrendamiento, comodato "
        "o documento que acredita el derecho de uso de un inmueble? "
        "Responde ÚNICAMENTE: VALIDO o INVALIDO, seguido de dos puntos y razón breve."
    ),
}


# ── Resultado de validación ───────────────────────────────────────────────────

class ResultadoIA:
    def __init__(self, valido: bool, razon: str, modelo: str):
        self.valido  = valido
        self.razon   = razon
        self.modelo  = modelo

    def __repr__(self) -> str:
        estado = "VALIDO" if self.valido else "INVALIDO"
        return f"ResultadoIA({estado}: {self.razon})"


# ── Validador principal ───────────────────────────────────────────────────────

def validar_documento(ruta: Path, tipo_documento: str) -> Optional[ResultadoIA]:
    """
    Valida semánticamente un documento usando el LLM configurado.

    Args:
        ruta:            Ruta local del archivo a validar.
        tipo_documento:  Clave del documento ("CEDULA", "COMERCIO", "RUT", "TENENCIA").

    Returns:
        ResultadoIA si la validación se completó, None si IA_HABILITADO=false o falla.
    """
    if not IA_HABILITADO:
        return None

    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada. Validación IA omitida.")
        return None

    prompt_sistema = _PROMPTS.get(tipo_documento)
    if not prompt_sistema:
        logger.warning("Tipo de documento sin prompt configurado: %s", tipo_documento)
        return None

    try:
        return _llamar_llm(ruta, tipo_documento, prompt_sistema)
    except Exception as exc:
        logger.error("Error en validación IA para '%s' (%s): %s", ruta.name, tipo_documento, exc)
        return None


def _llamar_llm(ruta: Path, tipo_documento: str, prompt_sistema: str) -> ResultadoIA:
    """Construye el mensaje y llama a la API de Anthropic."""
    import anthropic

    cliente = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    contenido_mensaje = _construir_contenido(ruta)

    if not contenido_mensaje:
        return ResultadoIA(
            valido=False,
            razon="No se pudo extraer contenido del archivo.",
            modelo=IA_MODEL,
        )

    respuesta = cliente.messages.create(
        model=IA_MODEL,
        max_tokens=256,
        system=prompt_sistema,
        messages=[{"role": "user", "content": contenido_mensaje}],
    )

    texto = respuesta.content[0].text.strip()
    valido = texto.upper().startswith("VALIDO")
    razon  = texto.split(":", 1)[1].strip() if ":" in texto else texto

    return ResultadoIA(valido=valido, razon=razon, modelo=IA_MODEL)


def _construir_contenido(ruta: Path) -> list | None:
    """
    Construye el bloque de contenido para el mensaje al LLM.
    Prioriza imagen (visión) sobre texto para documentos escaneados.
    """
    ext = ruta.suffix.lower()

    # Documento visual: enviar como imagen para aprovechar visión del modelo
    if ext in {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"}:
        b64 = extraer_imagen_base64(ruta)
        if b64:
            media_type = "image/jpeg" if ext in {".jpg", ".jpeg"} else "image/png"
            return [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": media_type, "data": b64},
                },
                {"type": "text", "text": "Analiza este documento."},
            ]

    # PDF u otro: extraer texto
    texto = extraer_texto(ruta)
    if texto:
        return [{"type": "text", "text": f"Contenido del documento:\n\n{texto[:8000]}"}]

    return None
