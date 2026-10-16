"""
Configuración centralizada. Todos los parámetros ajustables del sistema
se leen aquí desde variables de entorno, con valores por defecto sensatos.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# El .env MANDA sobre las variables ya definidas en el entorno del sistema. Antes
# no las pisaba y una OPENAI_API_KEY vieja de Windows (sin crédito) tapaba la
# clave vigente del .env, causando errores 429 "no credits remaining".
load_dotenv(override=True)

# Directorio raíz del proyecto (un nivel arriba de este archivo)
BASE_DIR = Path(__file__).resolve().parent.parent

# ── Rutas de trabajo ──────────────────────────────────────────────────────────
JOBS_DIR      = Path(os.getenv("JOBS_DIR",      str(BASE_DIR / "jobs")))
DOWNLOADS_DIR = Path(os.getenv("DOWNLOADS_DIR", str(BASE_DIR / "descargas_sharepoint")))

JOBS_DIR.mkdir(parents=True, exist_ok=True)
DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)

# ── Descarga SharePoint ───────────────────────────────────────────────────────
MAX_WORKERS_LISTADO  = int(os.getenv("MAX_WORKERS_LISTADO",  "8"))
MAX_WORKERS_DESCARGA = int(os.getenv("MAX_WORKERS_DESCARGA", "12"))
CHUNK_SIZE           = 4 * 1024 * 1024   # 4 MB

# ── Validación ────────────────────────────────────────────────────────────────
CHECKLIST_LOTE_GUARDADO = int(os.getenv("CHECKLIST_LOTE_GUARDADO", "50"))
TIMEOUT_DESCARGA        = int(os.getenv("TIMEOUT_DESCARGA", "120"))

# ── Análisis IA (OpenAI) ──────────────────────────────────────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
# gpt-4o: mejor comprensión visual de documentos escaneados con escritura manuscrita.
OPENAI_MODEL   = os.getenv("OPENAI_MODEL", "gpt-4o")
IA_HABILITADO  = os.getenv("IA_HABILITADO", "false").lower() == "true"

# Modelos del estudio de mercado (carpeta 04).
#   OPENAI_MODEL_PRINCIPAL : visión que valida capturas y lee precios.
#   OPENAI_MODEL_WEBSEARCH : modelo que ejecuta web_search_preview (búsqueda de URLs).
# Si el modelo configurado no existe / no soporta la herramienta, el flujo cae
# automáticamente a OPENAI_MODEL_RESPALDO (ver cotizacion_web._con_respaldo).
OPENAI_MODEL_PRINCIPAL = os.getenv("OPENAI_MODEL_PRINCIPAL", "gpt-5.6-luna")
OPENAI_MODEL_WEBSEARCH = os.getenv("OPENAI_MODEL_WEBSEARCH", OPENAI_MODEL_PRINCIPAL)
OPENAI_MODEL_RESPALDO  = os.getenv("OPENAI_MODEL_RESPALDO",  "gpt-4o")

# ── Búsqueda web adicional (estudio de mercado carpeta 04) ────────────────────
# SerpAPI: búsqueda Google Shopping + orgánico para Colombia.
#   Obtener clave en https://serpapi.com/  (plan gratuito disponible)
SERPAPI_KEY = os.getenv("SERPAPI_KEY", "")

# Google Custom Search: búsqueda alternativa.
#   Requiere clave en https://console.cloud.google.com/ y motor en
#   https://programmablesearch.google.com/
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
GOOGLE_CSE_ID  = os.getenv("GOOGLE_CSE_ID",  "")

# Páginas máximas a enviar a la IA por documento (primeras N/2 + últimas N/2).
IA_MAX_PAGINAS_COMPROMISO  = int(os.getenv("IA_MAX_PAGINAS_COMPROMISO",  "10"))
IA_MAX_PAGINAS_VISITA      = int(os.getenv("IA_MAX_PAGINAS_VISITA",      "15"))
IA_MAX_PAGINAS_TRATAMIENTO = int(os.getenv("IA_MAX_PAGINAS_TRATAMIENTO", "10"))
IA_DPI     = int(os.getenv("IA_DPI",     "200"))
IA_TIMEOUT = int(os.getenv("IA_TIMEOUT", "120"))
