"""
Análisis IA de la carpeta 03 (Capacitación / Formación).

Dos funciones principales:
  extraer_nombre_encuesta  — extrae el nombre del asistente del PDF de encuesta
                             usando texto del PDF (sin IA, formato fijo: "ENCUESTA – NOMBRE").
  validar_nombres_en_txrx  — usa GPT-4.1 para verificar que los nombres extraídos
                             de las encuestas aparezcan en el documento TX_RX del módulo.
"""

import json
import logging
import re
from pathlib import Path

from app.config import IA_TIMEOUT, OPENAI_API_KEY, OPENAI_MODEL
from app.ia.analizador_visita import _obtener_imagenes_b64, _parsear_json, ResultadoAnalisis

logger = logging.getLogger(__name__)

_MODELO_TXRX = "gpt-4.1"


# ── Extracción de nombre desde encuesta (sin IA) ──────────────────────────────

def extraer_nombre_encuesta(ruta: Path) -> str:
    """
    Extrae el nombre del asistente del encabezado de la encuesta PDF.
    Formato esperado en el título: "ENCUESTA – NOMBRE COMPLETO"
    Usa PyMuPDF para leer el texto del PDF; no consume tokens de IA.
    """
    try:
        import fitz
    except ImportError:
        logger.error("PyMuPDF no instalado.")
        return ""

    try:
        doc   = fitz.open(str(ruta))
        texto = doc[0].get_text() if len(doc) > 0 else ""
        doc.close()
    except Exception as exc:
        logger.warning("Error leyendo encuesta '%s': %s", ruta.name, exc)
        return ""

    for linea in texto.splitlines():
        linea = linea.strip()
        # Acepta guion largo (–), guion normal (-) y raya (—)
        m = re.match(r'ENCUESTA\s*[–\-—]\s*(.+)', linea, re.IGNORECASE)
        if m:
            nombre = m.group(1).strip()
            if nombre:
                return nombre

    logger.warning("No se encontró nombre en encuesta '%s'", ruta.name)
    return ""


# ── Validación de nombres en TX_RX con GPT-4.1 ───────────────────────────────

def validar_nombres_en_txrx(
    ruta: Path,
    nombres: list[str],
    modulo: str,
) -> ResultadoAnalisis:
    """
    Convierte el documento TX_RX a imágenes y le pide a GPT-4.1 que verifique
    cuáles de los nombres de las encuestas están presentes en la lista de asistencia.

    Compara ignorando tildes y mayúsculas; acepta variaciones menores de escritura
    manuscrita típicas de listas de asistencia colombianas.
    """
    if not OPENAI_API_KEY:
        return ResultadoAnalisis(ok=True, alerta="")

    if not nombres:
        return ResultadoAnalisis(ok=True, alerta="")

    imagenes = _obtener_imagenes_b64(ruta, 15)
    if not imagenes:
        return ResultadoAnalisis(
            ok=False,
            alerta=f"{modulo} — No se pudo leer el archivo TX_RX",
        )

    from openai import OpenAI

    cliente = OpenAI(api_key=OPENAI_API_KEY, timeout=float(IA_TIMEOUT))

    lista_nombres = "\n".join(f"- {n}" for n in nombres)
    prompt = f"""Analiza este documento de lista de asistencia (TX_RX o RX_TX) del {modulo}.

Necesito verificar si las siguientes personas aparecen en este documento.
El documento puede estar escrito a mano y los nombres pueden estar abreviados,
incompletos o con variaciones caligráficas.

REGLAS DE COINCIDENCIA — una persona se considera PRESENTE si cualquier combinación
de sus palabras aparece en el documento. Las personas colombianas tienen hasta 2 nombres
y 2 apellidos; en listas de asistencia escritas a mano suelen aparecer de forma abreviada.

  1. Toma TODAS las palabras del nombre completo (nombres y apellidos por separado).
     Busca CUALQUIER subconjunto de al menos 2 palabras que aparezca junto en el documento.
     Ejemplos válidos para "SILVIA FERNANDA GARCIA VARGAS":
       • SILVIA GARCIA           ✅ (1er nombre + 1er apellido)
       • FERNANDA GARCIA         ✅ (2do nombre + 1er apellido)
       • SILVIA VARGAS           ✅ (1er nombre + 2do apellido)
       • FERNANDA VARGAS         ✅ (2do nombre + 2do apellido)
       • SILVIA FERNANDA GARCIA  ✅ (nombres + 1er apellido)
       • GARCIA VARGAS           ✅ (solo apellidos)
       • S. GARCIA               ✅ (inicial + apellido)
  2. Se ignoran tildes, mayúsculas y pequeñas variaciones de escritura manuscrita.
  3. Si la escritura a mano es ambigua pero hay similitud fonética o visual clara,
     se considera presente.
  4. Solo marcar como AUSENTE cuando ninguna combinación razonable de las palabras
     del nombre aparezca en el documento.

Personas a verificar:
{lista_nombres}

Responde SOLO con este JSON:
{{"presentes": ["nombres que SÍ aparecen (usa el nombre original de la lista)"], "ausentes": ["nombres que definitivamente NO aparecen tras aplicar las reglas de coincidencia"], "observacion": "nota breve si hubo ambigüedad o coincidencia parcial, null si no"}}"""

    contenido: list[dict] = [{"type": "text", "text": prompt}]
    for img_b64 in imagenes:
        contenido.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{img_b64}", "detail": "high"},
        })

    try:
        resp = cliente.chat.completions.create(
            model=_MODELO_TXRX,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Eres un validador de listas de asistencia colombianas. "
                        "Respondes ÚNICAMENTE con JSON válido, sin texto adicional."
                    ),
                },
                {"role": "user", "content": contenido},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        data = _parsear_json(resp.choices[0].message.content or "{}")
    except Exception as exc:
        logger.error("Error IA TX_RX '%s' (%s): %s", ruta.name, modulo, exc)
        return ResultadoAnalisis(
            ok=False,
            alerta=f"{modulo} — Error IA al revisar TX_RX: {exc}",
        )

    ausentes = [str(n).strip() for n in data.get("ausentes", []) if str(n).strip()]
    if ausentes:
        return ResultadoAnalisis(
            ok=False,
            alerta=f"{modulo} — No aparecen en TX_RX: {', '.join(ausentes)}",
        )

    return ResultadoAnalisis(ok=True)
