"""
Router de validación documental.
Gestiona el ciclo de vida de los trabajos (job): inicio, estado y descarga del resultado.
"""

import asyncio
import logging
import uuid
from pathlib import Path
from typing import Any, Dict

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import JOBS_DIR
from app.services.procesador import ValidadorDocumental

logger   = logging.getLogger(__name__)
router   = APIRouter()

# Almacén en memoria de trabajos activos.
# Para producción multi-instancia reemplazar por Redis o base de datos.
_jobs: Dict[str, Dict[str, Any]] = {}


@router.post("/validar", summary="Inicia la validación de una matriz Excel")
async def iniciar_validacion(
    archivo: UploadFile = File(..., description="Archivo .xlsx de la matriz de entrada"),
):
    """
    Sube el Excel de la matriz y lanza el proceso en background.
    Responde inmediatamente con job_id y URLs de estado/resultado.
    """
    job_id          = str(uuid.uuid4())
    job_dir         = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    ruta_entrada    = job_dir / "matriz.xlsx"
    ruta_checklist  = job_dir / "checklist.xlsx"

    contenido = await archivo.read()
    ruta_entrada.write_bytes(contenido)
    logger.info("Job %s: archivo recibido (%d bytes).", job_id, len(contenido))

    _jobs[job_id] = {
        "estado":           "iniciado",
        "filas_total":      0,
        "filas_procesadas": 0,
        "errores":          0,
        "ruta_checklist":   str(ruta_checklist),
    }

    task = asyncio.create_task(
        asyncio.to_thread(_ejecutar_validacion, job_id, str(ruta_entrada), str(ruta_checklist))
    )
    _jobs[job_id]["_task"] = task

    return {
        "job_id":        job_id,
        "estado":        "iniciado",
        "estado_url":    f"/validacion/estado/{job_id}",
        "resultado_url": f"/validacion/resultado/{job_id}",
    }


@router.get("/estado/{job_id}", summary="Consulta el estado de un trabajo")
async def estado_trabajo(job_id: str):
    _verificar_job(job_id)
    return {k: v for k, v in _jobs[job_id].items() if not k.startswith("_")}


@router.get(
    "/resultado/{job_id}",
    summary="Descarga el checklist Excel resultante",
    response_class=FileResponse,
)
async def descargar_resultado(job_id: str):
    """Solo disponible si el estado es 'completado'."""
    _verificar_job(job_id)
    job = _jobs[job_id]

    if job["estado"] != "completado":
        raise HTTPException(
            status_code=400,
            detail=(
                f"El trabajo aún no terminó (estado: {job['estado']}). "
                f"Procesadas: {job['filas_procesadas']}/{job['filas_total']}."
            ),
        )

    ruta = Path(job["ruta_checklist"])
    if not ruta.exists():
        raise HTTPException(status_code=404, detail="Archivo de resultado no encontrado en disco.")

    return FileResponse(
        str(ruta),
        filename="checklist_validacion.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Función de fondo ──────────────────────────────────────────────────────────

def _ejecutar_validacion(job_id: str, ruta_entrada: str, ruta_checklist: str) -> None:
    try:
        _jobs[job_id]["estado"] = "procesando"

        def _progreso(procesadas: int, total: int, errores: int) -> None:
            _jobs[job_id]["filas_procesadas"] = procesadas
            _jobs[job_id]["filas_total"]      = total
            _jobs[job_id]["errores"]          = errores

        ValidadorDocumental(
            ruta_checklist=ruta_checklist,
            callback_progreso=_progreso,
        ).procesar_matriz(ruta_entrada)

        _jobs[job_id]["estado"] = "completado"
        logger.info("Job %s: completado.", job_id)

    except Exception as exc:
        logger.exception("Job %s: error fatal: %s", job_id, exc)
        _jobs[job_id]["estado"]    = "error"
        _jobs[job_id]["error_msg"] = str(exc)


def _verificar_job(job_id: str) -> None:
    if job_id not in _jobs:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' no encontrado.")
