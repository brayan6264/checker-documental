"""
Reglas de validación documental por modalidad y detección de documentos por nombre de archivo.

Fuente de verdad única: si cambian las reglas de negocio, solo cambia este archivo.
"""

from enum import Enum
from typing import Dict


class Estado(str, Enum):
    """Estado resultante de cada documento lógico para una fila del checklist."""
    OK        = "OK"                   # Obligatorio + presente
    FALTA     = "FALTA (obligatorio)"  # Obligatorio + ausente  → alerta roja
    PRESENTE  = "Presente"             # Opcional    + presente
    AUSENTE   = "Ausente"              # Opcional    + ausente
    NO_APLICA = "N/A"                  # No aplica (independiente de presencia)


# ── Tabla de reglas por modalidad ─────────────────────────────────────────────
#
# CEDULA se colapsa en un único tipo porque la detección por nombre de archivo
# es idéntica para "cédula propietario" y "cédula representante/participantes".

REGLAS: Dict[str, Dict[str, str]] = {
    "M1": {
        "CEDULA":   "obligatorio",
        "COMERCIO": "no_aplica",
        "RUT":      "opcional",
        "TENENCIA": "opcional",
    },
    "M2": {
        "CEDULA":   "obligatorio",
        "COMERCIO": "opcional",
        "RUT":      "opcional",
        "TENENCIA": "obligatorio",
    },
    "M3": {
        "CEDULA":   "obligatorio",
        "COMERCIO": "obligatorio",
        "RUT":      "obligatorio",
        "TENENCIA": "obligatorio",
    },
    "M4": {
        "CEDULA":   "obligatorio",
        "COMERCIO": "obligatorio",
        "RUT":      "obligatorio",
        "TENENCIA": "obligatorio",
    },
}

# ── Palabras clave para detección por nombre de archivo ───────────────────────
#
# Un archivo se considera del tipo X si su stem normalizado contiene
# alguna de las palabras clave asociadas a ese tipo.

PALABRAS_CLAVE: Dict[str, list] = {
    "CEDULA":   ["cedula"],
    "COMERCIO": ["comercio"],
    "RUT":      ["rut"],
    "TENENCIA": ["tenencia"],
}

# Orden canónico de columnas en el checklist.
DOCUMENTOS_ORDEN = ["CEDULA", "COMERCIO", "RUT", "TENENCIA"]


# ── Lógica de evaluación ──────────────────────────────────────────────────────

def evaluar_documentos(modalidad: str, presentes: Dict[str, bool]) -> Dict[str, Estado]:
    """
    Determina el estado de cada documento para la modalidad dada.

    Args:
        modalidad: "M1", "M2", "M3" o "M4".
        presentes: {nombre_doc: bool} indicando si el archivo fue encontrado.

    Returns:
        {nombre_doc: Estado} con el resultado de cada documento.

    Raises:
        KeyError si la modalidad no está en REGLAS (verificar antes de llamar).
    """
    reqs = REGLAS[modalidad]
    resultado: Dict[str, Estado] = {}

    for doc in DOCUMENTOS_ORDEN:
        req        = reqs.get(doc, "no_aplica")
        encontrado = presentes.get(doc, False)

        if req == "no_aplica":
            resultado[doc] = Estado.NO_APLICA
        elif req == "obligatorio":
            resultado[doc] = Estado.OK if encontrado else Estado.FALTA
        else:  # opcional
            resultado[doc] = Estado.PRESENTE if encontrado else Estado.AUSENTE

    return resultado


def estados_por_defecto(modalidad: str) -> Dict[str, Estado]:
    """
    Estados de fallback cuando no se puede validar (carpeta no encontrada,
    descarga fallida, etc.).

    Los documentos obligatorios quedan como FALTA para que el checklist los
    resalte en rojo; las observaciones explicarán el motivo real.
    """
    if modalidad not in REGLAS:
        return {doc: Estado.FALTA for doc in DOCUMENTOS_ORDEN}

    reqs = REGLAS[modalidad]
    return {
        doc: (
            Estado.FALTA     if reqs[doc] == "obligatorio" else
            Estado.NO_APLICA if reqs[doc] == "no_aplica"  else
            Estado.AUSENTE
        )
        for doc in DOCUMENTOS_ORDEN
    }
