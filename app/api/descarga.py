"""
Router de descarga SharePoint.
Expone la lógica de app.services.sharepoint como endpoint HTTP
para clientes externos o para pruebas independientes.
"""

import shutil
import uuid

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from app.config import DOWNLOADS_DIR, MAX_WORKERS_DESCARGA, MAX_WORKERS_LISTADO
from app.services.sharepoint import descargar_carpeta

router = APIRouter()


class DescargarRequest(BaseModel):
    url: str


@router.get("/health")
def health():
    return {
        "status":            "ok",
        "service":           "descarga-sharepoint",
        "downloads_dir":     str(DOWNLOADS_DIR.resolve()),
        "workers_listado":   MAX_WORKERS_LISTADO,
        "workers_descarga":  MAX_WORKERS_DESCARGA,
    }


@router.post("/descargar")
def descargar(req: DescargarRequest):
    """
    Descarga una carpeta SharePoint y retorna su ruta local.

    Request:  { "url": "<URL compartida de SharePoint>" }
    Response: { "ruta_local": "...", "total": N, "exitosos": N, "fallidos": N }
    """
    if not req.url.strip():
        raise HTTPException(status_code=400, detail="La URL viene vacía.")

    dest_base = DOWNLOADS_DIR / str(uuid.uuid4())

    try:
        ruta_local, stats = descargar_carpeta(req.url, dest_base)
        return {"ruta_local": str(ruta_local.resolve()), **stats}

    except HTTPException:
        raise
    except Exception as exc:
        if dest_base.exists():
            shutil.rmtree(dest_base, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(exc))
