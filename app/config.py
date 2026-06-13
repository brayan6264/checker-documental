"""
Configuración centralizada. Todos los parámetros ajustables del sistema
se leen aquí desde variables de entorno, con valores por defecto sensatos.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()  # carga .env si existe, sin pisar variables ya definidas en el entorno

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
TIMEOUT_DESCARGA        = int(os.getenv("TIMEOUT_DESCARGA", "300"))

# ── Análisis IA (OpenAI) ──────────────────────────────────────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
# gpt-4o: mejor comprensión visual de documentos escaneados con escritura manuscrita.
OPENAI_MODEL   = os.getenv("OPENAI_MODEL", "gpt-4o")
IA_HABILITADO  = os.getenv("IA_HABILITADO", "false").lower() == "true"

# Páginas máximas a enviar a la IA por documento (primeras N/2 + últimas N/2).
IA_MAX_PAGINAS_COMPROMISO  = int(os.getenv("IA_MAX_PAGINAS_COMPROMISO",  "10"))
IA_MAX_PAGINAS_VISITA      = int(os.getenv("IA_MAX_PAGINAS_VISITA",      "15"))
IA_MAX_PAGINAS_TRATAMIENTO = int(os.getenv("IA_MAX_PAGINAS_TRATAMIENTO", "10"))
IA_DPI     = int(os.getenv("IA_DPI",     "200"))
IA_TIMEOUT = int(os.getenv("IA_TIMEOUT", "120"))
