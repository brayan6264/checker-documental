"""
Configuración centralizada. Todos los parámetros ajustables del sistema
se leen aquí desde variables de entorno, con valores por defecto sensatos.
"""

import os
from pathlib import Path

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

# ── Análisis IA ───────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
IA_MODEL          = os.getenv("IA_MODEL", "claude-haiku-4-5-20251001")
IA_HABILITADO     = os.getenv("IA_HABILITADO", "false").lower() == "true"
