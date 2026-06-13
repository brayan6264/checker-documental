"""
Motor principal de validación documental.

Responsabilidades:
  1. Leer la matriz Excel (hoja activa).
  2. Por cada fila: descargar la carpeta SharePoint, localizar 00_DOCUMENTACION,
     detectar documentos presentes y evaluar según modalidad.
  3. Registrar el resultado en el ChecklistWriter (incremental + reanudable).
  4. Borrar la carpeta descargada tras cada fila para liberar disco.
  5. Nunca abortar por un error individual: los errores van a "observaciones".

La descarga se delega directamente a app.services.sharepoint (llamada de función,
sin overhead HTTP), aunque el router /descargar sigue disponible para uso externo.
"""

import logging
import shutil
from pathlib import Path
from typing import Callable, Dict, Optional

import openpyxl
import pandas as pd

from app.config import CHECKLIST_LOTE_GUARDADO, DOWNLOADS_DIR
from app.core.normalizacion import normalizar
from app.core.reglas import (
    DOCUMENTOS_ORDEN,
    PALABRAS_CLAVE,
    REGLAS,
    Estado,
    evaluar_documentos,
    estados_por_defecto,
)
from app.services.checklist import ChecklistWriter
from app.services.sharepoint import descargar_carpeta_doc

logger = logging.getLogger(__name__)


class ValidadorDocumental:
    """
    Componente de validación documental.

    Uso mínimo:
        v = ValidadorDocumental("checklist.xlsx")
        v.procesar_matriz("matriz.xlsx")
    """

    def __init__(
        self,
        ruta_checklist: str,
        callback_progreso: Optional[Callable[[int, int, int], None]] = None,
        lote_guardado: int = CHECKLIST_LOTE_GUARDADO,
    ):
        """
        Args:
            ruta_checklist:    Ruta del Excel de salida. Si existe, reanuda desde ahí.
            callback_progreso: Función opcional (filas_procesadas, total, errores)
                               para reportar progreso (p. ej. al job de FastAPI).
            lote_guardado:     Filas a acumular antes de persistir el checklist a disco.
        """
        self.ruta_checklist    = ruta_checklist
        self.callback_progreso = callback_progreso
        self.lote_guardado     = lote_guardado

    # ── Punto de entrada principal ────────────────────────────────────────────

    def procesar_matriz(self, ruta_excel: str) -> None:
        """
        Procesa todas las filas de la matriz Excel y genera el checklist.
        Reanudable: salta los ID_unico que ya estén en el checklist.
        """
        logger.info("Cargando matriz: %s", ruta_excel)
        df    = self._cargar_matriz(ruta_excel)
        total = len(df)
        logger.info("Total filas: %d", total)

        checklist      = ChecklistWriter(self.ruta_checklist, lote_guardado=self.lote_guardado)
        ya_procesados  = checklist.ids_procesados()
        procesadas     = len(ya_procesados)
        errores        = 0

        if ya_procesados:
            logger.info("Reanudando: %d filas ya procesadas, se saltearán.", procesadas)

        for _, fila in df.iterrows():
            id_unico = str(fila.get("ID_unico", "")).strip()

            if not id_unico or id_unico.lower() in ("nan", "none", ""):
                continue
            if id_unico in ya_procesados:
                continue

            logger.info("[%d/%d] ID: %s", procesadas + 1, total, id_unico)

            resultado = self._procesar_fila(fila)

            if resultado.get("_tiene_error"):
                errores += 1
                logger.warning("  Alerta: %s", resultado.get("observaciones"))

            checklist.agregar_fila(resultado)
            procesadas += 1

            if self.callback_progreso:
                self.callback_progreso(procesadas, total, errores)

        checklist.guardar()
        logger.info(
            "Proceso terminado. Procesadas: %d/%d. Errores/alertas: %d.",
            procesadas, total, errores,
        )

    # ── Procesamiento de una fila ─────────────────────────────────────────────

    def _procesar_fila(self, fila: pd.Series) -> dict:
        """Procesa una fila. Nunca lanza excepción: los errores van a observaciones."""
        id_unico   = str(fila.get("ID_unico",      "")).strip()
        modalidad  = str(fila.get("modalidad",     "")).strip().upper()
        unidad_doc = str(fila.get("unidad_doc",    "")).strip()
        link       = str(fila.get("carpetas_link", "")).strip()
        integrante = self._componer_integrante(fila)

        observaciones: list[str] = []
        tiene_error = False

        # ── 1. Validar modalidad ──────────────────────────────────────────────
        if modalidad not in REGLAS:
            observaciones.append(f"Modalidad inválida: '{modalidad}'")
            return self._armar_resultado(
                id_unico, modalidad, unidad_doc, integrante,
                estados_por_defecto(modalidad), observaciones, tiene_error=True,
            )

        # ── 2. Buscar y descargar solo 00_DOCUMENTACION desde SharePoint ─────
        # descargar_carpeta_doc() localiza la subcarpeta en SharePoint antes de
        # descargar, evitando bajar carpetas irrelevantes (fotos, visitas, etc.).
        ruta_doc: Optional[Path] = None
        dest_base = DOWNLOADS_DIR / id_unico
        try:
            logger.info("  [%s] Buscando y descargando 00_DOCUMENTACION...", id_unico)
            ruta_doc, _ = descargar_carpeta_doc(link, dest_base)
            if ruta_doc is None:
                observaciones.append("Carpeta 00_DOCUMENTACION no encontrada en SharePoint")
                return self._armar_resultado(
                    id_unico, modalidad, unidad_doc, integrante,
                    estados_por_defecto(modalidad), observaciones, tiene_error=True,
                )
            logger.info("  [%s] Descarga completa → %s", id_unico, ruta_doc)
        except Exception as exc:
            observaciones.append(f"Descarga fallida: {exc}")
            return self._armar_resultado(
                id_unico, modalidad, unidad_doc, integrante,
                estados_por_defecto(modalidad), observaciones, tiene_error=True,
            )

        # ── 3-4. Validar documentos (limpieza garantizada en finally) ─────────
        docs_estado = estados_por_defecto(modalidad)
        try:
            archivos    = self._listar_archivos(str(ruta_doc))
            presentes   = self._detectar_documentos(archivos)
            docs_estado = evaluar_documentos(modalidad, presentes)
            logger.info("  [%s] Documentos detectados: %s", id_unico, presentes)

            faltantes = [d for d, e in docs_estado.items() if e == Estado.FALTA]
            if faltantes:
                observaciones.append(f"Obligatorios ausentes: {', '.join(faltantes)}")
                tiene_error = True

        except Exception as exc:
            observaciones.append(f"Error en validación: {exc}")
            tiene_error = True

        finally:
            # ── 5. Limpiar carpeta temporal ───────────────────────────────────
            if dest_base.exists():
                try:
                    shutil.rmtree(dest_base)
                    logger.info("  [%s] Carpeta temporal eliminada.", id_unico)
                except Exception as exc:
                    logger.warning("No se pudo eliminar '%s': %s", dest_base, exc)

        return self._armar_resultado(
            id_unico, modalidad, unidad_doc, integrante,
            docs_estado, observaciones, tiene_error,
        )

    # ── Listado y detección de documentos ────────────────────────────────────

    @staticmethod
    def _listar_archivos(ruta: str) -> list:
        import os
        nombres: list[str] = []
        for _, _, archivos in os.walk(ruta):
            nombres.extend(archivos)
        return nombres

    @staticmethod
    def _detectar_documentos(archivos: list) -> Dict[str, bool]:
        """
        Compara el stem normalizado de cada archivo contra las palabras clave.
        Usa el stem (sin extensión) para evitar falsos positivos por extensiones.
        """
        presentes = {doc: False for doc in PALABRAS_CLAVE}
        for archivo in archivos:
            stem_norm = normalizar(Path(archivo).stem)
            for doc, palabras in PALABRAS_CLAVE.items():
                if not presentes[doc] and any(p in stem_norm for p in palabras):
                    presentes[doc] = True
        return presentes

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _componer_integrante(fila: pd.Series) -> str:
        claves = [
            "integrante_nombre1", "integrante_nombre2",
            "integrante_apellido1", "integrante_apellido2",
        ]
        partes = []
        for clave in claves:
            valor = str(fila.get(clave, "")).strip()
            if valor and valor.lower() not in ("nan", "none"):
                partes.append(valor)
        return " ".join(partes)

    @staticmethod
    def _armar_resultado(
        id_unico: str,
        modalidad: str,
        unidad_doc: str,
        integrante: str,
        docs_estado: Dict[str, Estado],
        observaciones: list,
        tiene_error: bool,
    ) -> dict:
        return {
            "ID_unico":        id_unico,
            "modalidad":       modalidad,
            "unidad_doc":      unidad_doc,
            "integrante":      integrante,
            "docs":            {doc: estado.value for doc, estado in docs_estado.items()},
            "docs_estado_raw": docs_estado,
            "observaciones":   "; ".join(observaciones),
            "_tiene_error":    tiene_error,
        }

    # ── Carga de la matriz ────────────────────────────────────────────────────

    @staticmethod
    def _cargar_matriz(ruta: str) -> pd.DataFrame:
        """Carga la hoja activa sin hardcodear su nombre."""
        wb_tmp     = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        hoja_activa = wb_tmp.active.title
        wb_tmp.close()

        REQUERIDAS = {"ID_unico", "modalidad", "unidad_doc", "carpetas_link"}
        OPCIONALES = {
            "integrante_nombre1", "integrante_nombre2",
            "integrante_apellido1", "integrante_apellido2",
        }

        df = pd.read_excel(ruta, sheet_name=hoja_activa, dtype=str)

        faltantes = REQUERIDAS - set(df.columns)
        if faltantes:
            raise ValueError(
                f"Columnas requeridas no encontradas: {sorted(faltantes)}\n"
                f"Columnas presentes: {sorted(df.columns.tolist())}"
            )

        todas = list(REQUERIDAS | (OPCIONALES & set(df.columns)))
        return df[todas].fillna("")
