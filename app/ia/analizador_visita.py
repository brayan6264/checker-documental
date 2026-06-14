"""
Análisis con IA de los documentos de la carpeta 01_VISITA_1_CARACTERIZACION.

Modelo: GPT-4o (OpenAI)
  - Visión de alta fidelidad en documentos escaneados
  - Mejor lectura de escritura manuscrita y marcas en formularios colombianos

Los PDFs se convierten a imágenes con PyMuPDF antes de enviarlos.
Para documentos largos se envían las primeras N/2 y las últimas N/2 páginas.

Criterio central: un campo se considera VACÍO SOLO si está completamente en
blanco. Ante la duda (escaneo borroso, marca tenue, escritura pequeña) se
considera DILIGENCIADO para minimizar falsas alertas.
"""

import base64
import json
import logging
from dataclasses import dataclass
from pathlib import Path

from app.config import (
    IA_DPI,
    IA_HABILITADO,
    IA_MAX_PAGINAS_COMPROMISO,
    IA_MAX_PAGINAS_TRATAMIENTO,
    IA_MAX_PAGINAS_VISITA,
    IA_TIMEOUT,
    OPENAI_API_KEY,
    OPENAI_MODEL,
)

logger = logging.getLogger(__name__)

_EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff", ".tif"}


@dataclass
class ResultadoAnalisis:
    ok: bool
    alerta: str = ""
    gestor_nombre: str = ""
    gestor_cedula: str = ""           # cédula más probable
    gestor_cedulas_alt: list = None   # variantes alternativas si hay ambigüedad

    def __post_init__(self):
        if self.gestor_cedulas_alt is None:
            self.gestor_cedulas_alt = []

    def __str__(self) -> str:
        return "OK" if self.ok else f"ALERTA: {self.alerta}"


# ── Conversión PDF → imagen ───────────────────────────────────────────────────

def _pdf_a_imagenes_b64(ruta: Path, max_paginas: int) -> list[str]:
    try:
        import fitz
    except ImportError:
        logger.error("PyMuPDF no instalado. Ejecutar: pip install pymupdf")
        return []

    try:
        
        doc     = fitz.open(str(ruta))
        n       = len(doc)
        mitad   = max_paginas // 2
        indices = (
            list(range(n)) if n <= max_paginas
            else list(range(mitad)) + list(range(max(mitad, n - mitad), n))
        )
        mat      = fitz.Matrix(IA_DPI / 72, IA_DPI / 72)
        imagenes = []
        for i in indices:
            pix = doc[i].get_pixmap(matrix=mat)
            imagenes.append(base64.b64encode(pix.tobytes("jpeg")).decode())
        doc.close()
        return imagenes
    except Exception as exc:
        logger.warning("Error convirtiendo PDF '%s': %s", ruta.name, exc)
        return []


def _imagen_a_b64(ruta: Path) -> str | None:
    try:
        return base64.b64encode(ruta.read_bytes()).decode()
    except Exception as exc:
        logger.warning("Error leyendo imagen '%s': %s", ruta.name, exc)
        return None


def _obtener_imagenes_b64(ruta: Path, max_paginas: int) -> list[str]:
    ext = ruta.suffix.lower()
    if ext == ".pdf":
        return _pdf_a_imagenes_b64(ruta, max_paginas)
    if ext in _EXTENSIONES_IMAGEN:
        b64 = _imagen_a_b64(ruta)
        return [b64] if b64 else []
    logger.warning("Formato no soportado para análisis IA: %s", ext)
    return []


# ── Cliente OpenAI ────────────────────────────────────────────────────────────

def _llamar_gpt(imagenes_b64: list[str], prompt_sistema: str, prompt_usuario: str) -> str:
    """
    Envía imágenes + prompts a GPT-4o.
    Usa detail='high' para leer manuscrito y marcas en formularios escaneados.
    """
    from openai import OpenAI

    cliente = OpenAI(api_key=OPENAI_API_KEY, timeout=float(IA_TIMEOUT))

    contenido_usuario: list[dict] = [{"type": "text", "text": prompt_usuario}]
    for img_b64 in imagenes_b64:
        contenido_usuario.append({
            "type": "image_url",
            "image_url": {
                "url":    f"data:image/jpeg;base64,{img_b64}",
                "detail": "high",  # necesario para leer escritura manuscrita en escaneados
            },
        })

    resp = cliente.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[
            {"role": "system", "content": prompt_sistema},
            {"role": "user",   "content": contenido_usuario},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )
    contenido = resp.choices[0].message.content
    return contenido if contenido is not None else "{}"


def _parsear_json(texto: str) -> dict:
    try:
        return json.loads(texto)
    except Exception:
        logger.warning("Respuesta IA no era JSON válido: %s", texto[:200])
        return {}


# ── Prompt base (compartido) ──────────────────────────────────────────────────

_SISTEMA_BASE = """Eres un revisor de documentos colombianos escaneados (Programa DPS).
Respondes ÚNICAMENTE con JSON válido, sin texto adicional.

REGLA CRÍTICA sobre campos:
- Un campo está VACÍO solo si está completamente en blanco, sin NINGUNA marca, trazo ni escritura.
- Cualquier escritura (aunque sea ilegible), número, inicial, tachón, sello o marca hace que
  el campo se considere DILIGENCIADO.
- Si el escaneo es borroso o de baja calidad pero hay indicios de contenido, marca como DILIGENCIADO.
- Solo reporta un campo como vacío cuando tengas CERTEZA ABSOLUTA de que está en blanco.
- Ante cualquier duda: DILIGENCIADO."""


# ── Analizadores por tipo de documento ───────────────────────────────────────

def analizar_acta_compromiso(ruta: Path) -> ResultadoAnalisis:
    """
    Acta de Compromiso (Formato 1 DPS).
    Verifica: campos diligenciados + FIRMA + HUELLA al final.
    """
    if not IA_HABILITADO or not OPENAI_API_KEY:
        return ResultadoAnalisis(ok=True, alerta="")

    imagenes = _obtener_imagenes_b64(ruta, IA_MAX_PAGINAS_COMPROMISO)
    if not imagenes:
        return ResultadoAnalisis(ok=False, alerta="No se pudo leer el archivo para análisis IA")

    prompt_usuario = """Analiza este Acta de Compromiso (Formato 1 DPS) escaneada.

Verifica los siguientes puntos:
1. ¿Todos los campos del formulario tienen contenido escrito? (nombres, cédula, dirección, etc.)
   RECUERDA: solo marca como vacío si el campo está COMPLETAMENTE en blanco.
2. ¿Hay FIRMA del participante? (cualquier garabato, rúbrica o marca en el espacio de firma)
3. ¿Hay HUELLA dactilar? (cualquier mancha, marca o impresión visible en el espacio de huella)

Para firma y huella: si hay cualquier marca en el espacio designado = SÍ tiene.

Responde SOLO con este JSON:
{"campos_completos": true/false, "tiene_firma": true/false, "tiene_huella": true/false, "campos_vacios": ["solo campos COMPLETAMENTE en blanco, lista vacía si todos tienen contenido"], "alerta": "descripción breve del problema o null"}"""

    try:
        data = _parsear_json(_llamar_gpt(imagenes, _SISTEMA_BASE, prompt_usuario))
    except Exception as exc:
        logger.error("Error IA acta compromiso '%s': %s", ruta.name, exc)
        return ResultadoAnalisis(ok=False, alerta=f"Error en análisis IA: {exc}")

    problemas = []
    if not data.get("campos_completos", True):
        vacios = data.get("campos_vacios", [])
        if vacios:
            problemas.append(f"Campos vacíos: {', '.join(vacios)}")
    if not data.get("tiene_firma", True):
        problemas.append("Falta FIRMA")
    if not data.get("tiene_huella", True):
        problemas.append("Falta HUELLA")

    if problemas:
        return ResultadoAnalisis(ok=False, alerta="; ".join(problemas))
    return ResultadoAnalisis(ok=True)


def analizar_acta_visita(ruta: Path) -> ResultadoAnalisis:
    """
    Acta de Visita 1 (Formato 6 DPS).
    Verifica campos diligenciados; Anexos y Registros fotográficos pueden quedar vacíos.
    """
    if not IA_HABILITADO or not OPENAI_API_KEY:
        return ResultadoAnalisis(ok=True, alerta="")

    imagenes = _obtener_imagenes_b64(ruta, IA_MAX_PAGINAS_VISITA)
    if not imagenes:
        return ResultadoAnalisis(ok=False, alerta="No se pudo leer el archivo para análisis IA")

    prompt_usuario = """Analiza esta Acta de Visita 1 (Formato 6 DPS) escaneada.

CAMPOS QUE SIEMPRE PUEDEN ESTAR VACÍOS (ignóralos aunque estén en blanco):
  - Anexos
  - Registro Fotográfico / Registros fotográficos
  - Fecha de próxima reunión
  - Cualquier campo que diga "(si aplica)", "si aplica", "opcional", "cuando aplique"
    o cualquier variante similar en su etiqueta o instrucción del formulario.

PASO 1 — Clasifica cada campo:
  - OBLIGATORIO: campo sin ninguna indicación de opcionalidad. Debe tener contenido.
  - OPCIONAL: cualquiera de los listados arriba o que el formulario indique que puede quedar en blanco.

PASO 2 — Evalúa si cada campo OBLIGATORIO tiene contenido escrito.
  Un campo OBLIGATORIO está vacío SOLO si está completamente en blanco, sin ningún trazo.
  Escritura difícil de leer por baja calidad del escaneo = DILIGENCIADO.
  Los campos OPCIONALES vacíos NO son un error, ignóralos por completo.

PASO 3 — Extrae los datos del GESTOR escritos a mano en el acta.
  Busca un campo que diga "Gestor", "Nombre del gestor", "Responsable" o similar.

  LECTURA DE LA CÉDULA — sigue estos pasos con máxima atención:
  a) Amplía mentalmente la zona donde está escrita la cédula.
  b) Lee dígito por dígito de izquierda a derecha. Las cédulas colombianas tienen entre 6 y 10 dígitos.
  c) Ten especial cuidado con estos pares que frecuentemente se confunden en escritura a mano:
       • 2 ↔ 3  (el 2 puede verse como un 3 con la curva inferior cerrada)
       • 1 ↔ 7  (el 7 puede verse como un 1 con trazo superior)
       • 0 ↔ 8  (el 0 puede verse como un 8 incompleto, o viceversa)
       • 4 ↔ 9  (el 4 abierto se confunde con 9 según la escritura)
       • 5 ↔ 6  (el bucle inferior puede ser ambiguo)
       • 6 ↔ 0  (el 6 con bucle pequeño parece un 0)
  d) Para CADA dígito donde exista cualquier duda, anota las dos opciones más probables.
  e) Si tienes certeza total en todos los dígitos: devuelve solo "gestor_cedula" con ese número.
  f) Si hay 1 o más dígitos dudosos: devuelve en "gestor_cedula" la variante más probable Y en
     "gestor_cedulas_alt" una lista con TODAS las combinaciones alternativas plausibles
     (máximo 6 variantes), reemplazando cada dígito dudoso por su otra opción.
  g) No omitas ni agregues dígitos. Si no puedes leer la cédula, devuelve "gestor_cedula": "".

  LECTURA DEL NOMBRE — lee letra por letra. Si hay ambigüedad entre dos letras,
  elige la que forme una palabra reconocible en español.

Responde SOLO con este JSON:
{"campos_completos": true/false, "campos_vacios": ["lista de campos OBLIGATORIOS completamente en blanco; lista vacía si todos los obligatorios tienen contenido"], "gestor_nombre": "nombre completo del gestor o vacío si no aparece", "gestor_cedula": "cédula más probable o vacío si no se puede leer", "gestor_cedulas_alt": ["variantes alternativas de la cédula si hubo dígitos dudosos, si no hay dudas lista vacía"], "alerta": "descripción breve del problema o null"}"""

    try:
        data = _parsear_json(_llamar_gpt(imagenes, _SISTEMA_BASE, prompt_usuario))
    except Exception as exc:
        logger.error("Error IA acta visita '%s': %s", ruta.name, exc)
        return ResultadoAnalisis(ok=False, alerta=f"Error en análisis IA: {exc}")

    gestor_nombre     = str(data.get("gestor_nombre") or "").strip()
    gestor_cedula     = str(data.get("gestor_cedula") or "").strip()
    cedulas_alt_raw   = data.get("gestor_cedulas_alt", [])
    gestor_cedulas_alt = [str(c).strip() for c in (cedulas_alt_raw or []) if str(c).strip()]

    vacios = data.get("campos_vacios", [])
    if not data.get("campos_completos", True) and vacios:
        return ResultadoAnalisis(
            ok=False, alerta=f"Campos vacíos: {', '.join(vacios)}",
            gestor_nombre=gestor_nombre, gestor_cedula=gestor_cedula,
            gestor_cedulas_alt=gestor_cedulas_alt,
        )

    return ResultadoAnalisis(
        ok=True,
        gestor_nombre=gestor_nombre, gestor_cedula=gestor_cedula,
        gestor_cedulas_alt=gestor_cedulas_alt,
    )


def analizar_tratamiento_datos(ruta: Path) -> ResultadoAnalisis:
    """
    Autorización de Tratamiento de Datos Personales.
    Verifica: campos llenos + opción 'SÍ autorizo' marcada.
    """
    if not IA_HABILITADO or not OPENAI_API_KEY:
        return ResultadoAnalisis(ok=True, alerta="")

    imagenes = _obtener_imagenes_b64(ruta, IA_MAX_PAGINAS_TRATAMIENTO)
    if not imagenes:
        return ResultadoAnalisis(ok=False, alerta="No se pudo leer el archivo para análisis IA")

    prompt_usuario = """Analiza esta Autorización de Tratamiento de Datos Personales escaneada.

Verifica:
1. ¿Los campos del formulario están diligenciados? (nombres, cédula, fecha, etc.)
   RECUERDA: campo vacío = completamente en blanco. Escritura difícil de leer = DILIGENCIADO.

2. En la sección de autorización hay dos opciones:
   [ ] SÍ autorizo el uso de mis registros fotográficos, de video y/o de voz.
   [ ] NO autorizo el uso de mis registros fotográficos, de video y/o de voz.

   ¿Está marcada la opción "SÍ autorizo" con X, visto, tilde, tachón u otra marca?
   - Si "SÍ" está marcado: autoriza = true
   - Si "NO" está marcado o ninguna opción tiene marca: autoriza = false

Responde SOLO con este JSON:
{"campos_completos": true/false, "autoriza": true/false, "campos_vacios": ["campos COMPLETAMENTE en blanco, lista vacía si todo está bien"], "alerta": "descripción breve del problema o null"}"""

    try:
        data = _parsear_json(_llamar_gpt(imagenes, _SISTEMA_BASE, prompt_usuario))
    except Exception as exc:
        logger.error("Error IA tratamiento datos '%s': %s", ruta.name, exc)
        return ResultadoAnalisis(ok=False, alerta=f"Error en análisis IA: {exc}")

    problemas = []
    vacios = data.get("campos_vacios", [])
    if not data.get("campos_completos", True) and vacios:
        problemas.append(f"Campos vacíos: {', '.join(vacios)}")
    if not data.get("autoriza", True):
        problemas.append("No marcó 'SÍ autorizo' o marcó 'NO autorizo'")

    if problemas:
        return ResultadoAnalisis(ok=False, alerta="; ".join(problemas))
    return ResultadoAnalisis(ok=True)
