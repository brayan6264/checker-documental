"""
FastAPI app factory.

Monta los routers de descarga y validación bajo sus prefijos:
  /descarga/*    → app.api.descarga
  /validacion/*  → app.api.validacion

Arranque:
  uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
"""

import logging

from fastapi import FastAPI

from app.api import descarga, validacion

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

app = FastAPI(
    title="Validador Documental SharePoint",
    description=(
        "Descarga carpetas de SharePoint, valida la presencia de documentos "
        "obligatorios según modalidad y genera un checklist en Excel."
    ),
    version="2.0.0",
)

app.include_router(descarga.router,   prefix="/descarga",   tags=["Descarga"])
app.include_router(validacion.router, prefix="/validacion", tags=["Validación"])


@app.get("/health", tags=["Sistema"])
def health():
    return {"status": "ok", "version": app.version}
