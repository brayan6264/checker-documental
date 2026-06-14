"""
Router de validación documental.
Gestiona el ciclo de vida de los trabajos (job): inicio, estado y descarga del resultado.
"""

import asyncio
import logging
import time
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import JOBS_DIR
from app.services.procesador import ValidadorDocumental

logger   = logging.getLogger(__name__)
router   = APIRouter()

# Almacén en memoria de trabajos activos.
# Para producción multi-instancia reemplazar por Redis o base de datos.
_jobs: Dict[str, Dict[str, Any]] = {}


_FLUJOS_VALIDOS = {"00", "01", "02", "03"}


@router.post("/validar", summary="Inicia la validación de una matriz Excel")
async def iniciar_validacion(
    archivo: UploadFile = File(..., description="Archivo .xlsx de la matriz de entrada"),
    flujos: Optional[str] = Form(
        default=None,
        description=(
            "Flujos a ejecutar, separados por coma. "
            "Valores posibles: 00, 01, 02, 03. "
            "Dejar vacío para ejecutar todos. "
            "Ejemplos: '00,01'  |  '02,03'  |  '00'"
        ),
    ),
):
    """
    Sube el Excel de la matriz y lanza el proceso en background.
    Responde inmediatamente con job_id y URLs de estado/resultado.

    El parámetro `flujos` permite elegir qué carpetas validar:
    - **00** : Documentación (cédula, RUT, comercio, tenencia)
    - **01** : Visita 1 — Caracterización (acta, fotos, tratamiento de datos)
    - **02** : Visita 2 — Diagnóstico y plan de negocio
    - **03** : Capacitación (encuestas, grupos, módulos TX/RX)

    Se pueden combinar: `00,01` valida solo documentación y primera visita.
    """
    # Parsear y validar flujos
    if flujos and flujos.strip():
        flujos_set = {f.strip() for f in flujos.split(",")}
        invalidos = flujos_set - _FLUJOS_VALIDOS
        if invalidos:
            raise HTTPException(
                status_code=422,
                detail=f"Flujos inválidos: {invalidos}. Valores permitidos: {_FLUJOS_VALIDOS}",
            )
    else:
        flujos_set = None  # None = todos

    job_id          = str(uuid.uuid4())
    job_dir         = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    ruta_entrada    = job_dir / "matriz.xlsx"
    ruta_checklist  = job_dir / "checklist.xlsx"

    contenido = await archivo.read()
    ruta_entrada.write_bytes(contenido)
    logger.info(
        "Job %s: archivo recibido (%d bytes), flujos=%s.",
        job_id, len(contenido), flujos_set or "todos",
    )

    _jobs[job_id] = {
        "estado":           "iniciado",
        "filas_total":      0,
        "filas_procesadas": 0,
        "errores":          0,
        "flujos":           sorted(flujos_set) if flujos_set else ["00", "01", "02", "03"],
        "ruta_checklist":   str(ruta_checklist),
    }

    task = asyncio.create_task(
        asyncio.to_thread(
            _ejecutar_validacion, job_id, str(ruta_entrada), str(ruta_checklist), flujos_set,
        )
    )
    _jobs[job_id]["_task"] = task

    return {
        "job_id":        job_id,
        "estado":        "iniciado",
        "flujos":        sorted(flujos_set) if flujos_set else ["00", "01", "02", "03"],
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

def _ejecutar_validacion(
    job_id: str,
    ruta_entrada: str,
    ruta_checklist: str,
    flujos: Optional[set] = None,
) -> None:
    t0 = time.perf_counter()
    try:
        _jobs[job_id]["estado"]      = "procesando"
        _jobs[job_id]["inicio_utc"]  = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())

        def _progreso(procesadas: int, total: int, errores: int) -> None:
            _jobs[job_id]["filas_procesadas"] = procesadas
            _jobs[job_id]["filas_total"]      = total
            _jobs[job_id]["errores"]          = errores

        ValidadorDocumental(
            ruta_checklist=ruta_checklist,
            callback_progreso=_progreso,
            flujos=flujos,
        ).procesar_matriz(ruta_entrada)

        elapsed = time.perf_counter() - t0
        _jobs[job_id]["estado"]           = "completado"
        _jobs[job_id]["duracion_segundos"] = round(elapsed, 1)
        logger.info(
            "Job %s: completado en %s.",
            job_id, _fmt_duracion(elapsed),
        )

    except Exception as exc:
        elapsed = time.perf_counter() - t0
        logger.exception("Job %s: error fatal tras %s: %s", job_id, _fmt_duracion(elapsed), exc)
        _jobs[job_id]["estado"]            = "error"
        _jobs[job_id]["error_msg"]         = str(exc)
        _jobs[job_id]["duracion_segundos"] = round(elapsed, 1)


def _fmt_duracion(segundos: float) -> str:
    """Formatea segundos como '1h 23m 45s' o '2m 03s' o '45.3s'."""
    s = int(segundos)
    h, rem = divmod(s, 3600)
    m, seg = divmod(rem, 60)
    if h:
        return f"{h}h {m:02d}m {seg:02d}s"
    if m:
        return f"{m}m {seg:02d}s"
    return f"{segundos:.1f}s"


def _verificar_job(job_id: str) -> None:
    if job_id not in _jobs:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' no encontrado.")
