"""
Motor principal de validación documental.

Flujo por fila:
  1. Leer matriz Excel (hoja activa).
  2. Descargar y validar 00_DOCUMENTACION (CEDULA, COMERCIO, RUT, TENENCIA).
  3. Descargar y validar 01_VISITA_1_CARACTERIZACION:
       - Existencia de ACTA_COMPROMISO, ACTA_VISITA_1, TRATAMIENTO_DATOS.
       - Conteo mínimo de 4 fotos/videos.
       - Análisis IA de cada documento (si IA_HABILITADO=true).
  4. Registrar resultado en ChecklistWriter (incremental + reanudable).
  5. Borrar carpetas temporales.
  6. Nunca abortar por error individual: los errores van a "observaciones".
"""

import logging
import shutil
from pathlib import Path
from typing import Callable, Dict, Optional

import openpyxl
import pandas as pd
from app.core.visita2 import validar_visita2

from app.config import CHECKLIST_LOTE_GUARDADO, DOWNLOADS_DIR, IA_HABILITADO
from app.core.normalizacion import normalizar
from app.core.reglas import (
    DOCS_VISITA_ORDEN,
    DOCUMENTOS_ORDEN,
    DOCS_VISITA2_ORDEN,
    EXTENSIONES_MEDIA,
    MIN_ARCHIVOS_MEDIA,
    PALABRAS_CLAVE,
    PALABRAS_CLAVE_VISITA,
    REGLAS,
    Estado,
    evaluar_documentos,
    estados_por_defecto,
)
from app.ia.analizador_visita import (
    ResultadoAnalisis,
    analizar_acta_compromiso,
    analizar_acta_visita,
    analizar_tratamiento_datos,
)
from app.core.gestores import buscar_gestor
from app.services.checklist import ChecklistWriter
from app.services.sharepoint import descargar_carpeta_doc, descargar_visita_selectiva

from app.core.reglas import DOCS_VISITA2_ORDEN
from app.ia.analizador_visita2 import (analizar_plan_negocio, analizar_diagnostico)
from app.services.sharepoint import descargar_visita2_selectiva

logger = logging.getLogger(__name__)

_NO_ANALIZADO = ResultadoAnalisis(ok=True, alerta="No analizado (IA desactivada)")

# Mapa de keyword → función de análisis IA
_ANALIZADORES = {
    "ACTA_COMPROMISO":   analizar_acta_compromiso,
    "ACTA_VISITA_1":     analizar_acta_visita,
    "TRATAMIENTO_DATOS": analizar_tratamiento_datos,

    # Carpeta 02
    "ACTA_VISITA_2":     analizar_acta_visita,
}


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
        self.ruta_checklist    = ruta_checklist
        self.callback_progreso = callback_progreso
        self.lote_guardado     = lote_guardado

    # ── Punto de entrada principal ────────────────────────────────────────────

    def procesar_matriz(self, ruta_excel: str) -> None:
        logger.info("Cargando matriz: %s", ruta_excel)
        df    = self._cargar_matriz(ruta_excel)
        total = len(df)
        logger.info("Total filas: %d", total)

        checklist     = ChecklistWriter(self.ruta_checklist, lote_guardado=self.lote_guardado)
        ya_procesados = checklist.ids_procesados()
        procesadas    = len(ya_procesados)
        errores       = 0

        if ya_procesados:
            logger.info("Reanudando: %d filas ya procesadas.", procesadas)

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
        logger.info("Terminado. %d/%d procesadas, %d alertas.", procesadas, total, errores)

    # ── Procesamiento de una fila ─────────────────────────────────────────────

    def _procesar_fila(self, fila: pd.Series) -> dict:
        id_unico   = str(fila.get("ID_unico",      "")).strip()
        modalidad  = str(fila.get("modalidad",     "")).strip().upper()
        unidad_doc = str(fila.get("unidad_doc",    "")).strip()
        link       = str(fila.get("carpetas_link", "")).strip()
        integrante = self._componer_integrante(fila)

        observaciones: list[str] = []
        tiene_error = False
        dest_base   = DOWNLOADS_DIR / id_unico

        # ── 1. Validar modalidad ──────────────────────────────────────────────
        if modalidad not in REGLAS:
            observaciones.append(f"Modalidad inválida: '{modalidad}'")
            return self._armar_resultado(
                id_unico, modalidad, unidad_doc, integrante,
                estados_por_defecto(modalidad), {}, {}, observaciones, tiene_error=True,
            )

        # ── 2. Carpeta 00_DOCUMENTACION ───────────────────────────────────────
        docs_estado = estados_por_defecto(modalidad)
        ruta_doc: Optional[Path] = None
        try:
            logger.info("  [%s] Buscando 00_DOCUMENTACION...", id_unico)
            ruta_doc, _ = descargar_carpeta_doc(link, dest_base / "doc")
            if ruta_doc is None:
                observaciones.append("00_DOCUMENTACION no encontrada en SharePoint")
                tiene_error = True
            else:
                archivos    = self._listar_archivos(str(ruta_doc))
                presentes   = self._detectar_documentos(archivos)
                docs_estado = evaluar_documentos(modalidad, presentes)
                faltantes   = [d for d, e in docs_estado.items() if e == Estado.FALTA]
                if faltantes:
                    observaciones.append(f"00 — Obligatorios ausentes: {', '.join(faltantes)}")
                    tiene_error = True
                logger.info("  [%s] 00 OK — detectados: %s", id_unico, presentes)
        except Exception as exc:
            observaciones.append(f"00 — Descarga fallida: {exc}")
            tiene_error = True

        # ── 3. Carpeta 01_VISITA_1_CARACTERIZACION ────────────────────────────
        resultado_visita = self._procesar_visita(id_unico, link, dest_base / "visita")
        observaciones.extend(resultado_visita["observaciones"])
        if resultado_visita["tiene_error"]:
            tiene_error = True

        # ── 4. Carpeta 02_VISITA_2_DIAGNOSTICO ───────────────────────────────
        resultado_visita2 = self._procesar_visita2(id_unico, link, dest_base / "visita2")
        observaciones.extend(resultado_visita2["observaciones"])
        if resultado_visita2["tiene_error"]:
            tiene_error = True

        # ── 5. Limpiar temporales ─────────────────────────────────────────────
        if dest_base.exists():
            try:
                shutil.rmtree(dest_base)
                logger.info("  [%s] Temporales eliminados.", id_unico)
            except Exception as exc:
                logger.warning("No se pudo eliminar '%s': %s", dest_base, exc)

        return self._armar_resultado(
            id_unico, modalidad, unidad_doc, integrante,
            docs_estado, resultado_visita, resultado_visita2, observaciones, tiene_error,
        )

    # ── Validación carpeta 01 ─────────────────────────────────────────────────

    def _procesar_visita(self, id_unico: str, link: str, dest_base: Path) -> dict:
        """
        Descarga selectivamente la carpeta 01_VISITA, cuenta media,
        y ejecuta análisis IA sobre los 3 documentos.
        """
        obs: list[str] = []
        tiene_error    = False

        # Estados iniciales: FALTA para todos
        docs_visita = {k: "FALTA" for k in DOCS_VISITA_ORDEN}
        ia_resultados = {
            "IA_COMPROMISO":  _NO_ANALIZADO,
            "IA_VISITA":      _NO_ANALIZADO,
            "IA_TRATAMIENTO": _NO_ANALIZADO,
        }
        conteo_media  = 0
        gestor_nombre = ""
        gestor_cedula = ""
        gestor_ok     = None  # None = no analizado, True/False = resultado

        try:
            logger.info("  [%s] Buscando 01_VISITA...", id_unico)
            info = descargar_visita_selectiva(link, dest_base)

            if not info["encontrada"]:
                obs.append("01_VISITA_1_CARACTERIZACION no encontrada")
                tiene_error = True
                return {
                    "encontrada": False, "docs_visita": docs_visita,
                    "conteo_media": 0, "ia": ia_resultados,
                    "gestor_nombre": "", "gestor_cedula": "", "gestor_ok": None,
                    "observaciones": obs, "tiene_error": tiene_error,
                }

            conteo_media = info["conteo_media"]
            logger.info("  [%s] 01 — media encontrada: %d", id_unico, conteo_media)

            # Verificar mínimo de fotos/videos
            if conteo_media < MIN_ARCHIVOS_MEDIA:
                obs.append(
                    f"01 — Fotos/videos insuficientes: {conteo_media} "
                    f"(mínimo {MIN_ARCHIVOS_MEDIA})"
                )
                tiene_error = True

            # Verificar y analizar los 3 documentos
            for keyword, ruta_archivo in info["docs_rutas"].items():
                if ruta_archivo is None:
                    docs_visita[keyword] = "FALTA"
                    obs.append(f"01 — {keyword} no encontrado")
                    tiene_error = True
                    continue

                docs_visita[keyword] = "OK"

                if IA_HABILITADO:
                    logger.info("  [%s] Analizando IA: %s", id_unico, keyword)
                    try:
                        res = _ANALIZADORES[keyword](ruta_archivo)
                        ia_key = {
                            "ACTA_COMPROMISO":   "IA_COMPROMISO",
                            "ACTA_VISITA_1":     "IA_VISITA",
                            "TRATAMIENTO_DATOS": "IA_TRATAMIENTO",
                        }[keyword]
                        ia_resultados[ia_key] = res
                        if not res.ok:
                            obs.append(f"IA {keyword}: {res.alerta}")
                            tiene_error = True

                        # Verificar gestor extraído del acta de visita
                        if keyword == "ACTA_VISITA_1":
                            gestor_nombre     = res.gestor_nombre
                            gestor_cedula     = res.gestor_cedula
                            gestor_cedulas_alt = res.gestor_cedulas_alt or []
                            if gestor_nombre or gestor_cedula:
                                encontrado, nombre_bd = buscar_gestor(
                                    gestor_cedula, gestor_nombre, gestor_cedulas_alt
                                )
                                gestor_ok = encontrado
                                if not encontrado:
                                    alt_txt = f" | variantes: {', '.join(gestor_cedulas_alt)}" if gestor_cedulas_alt else ""
                                    obs.append(
                                        f"Gestor no registrado: {gestor_nombre} "
                                        f"(Cédula: {gestor_cedula}{alt_txt})"
                                    )
                                    tiene_error = True
                                else:
                                    logger.info("  [%s] Gestor verificado: %s", id_unico, nombre_bd)
                            else:
                                logger.warning("  [%s] No se extrajo gestor del acta de visita", id_unico)
                    except Exception as exc:
                        logger.error("Error IA %s: %s", keyword, exc)
                        obs.append(f"IA {keyword}: error — {exc}")
                        tiene_error = True

        except Exception as exc:
            obs.append(f"01 — Error procesando visita: {exc}")
            tiene_error = True

        return {
            "encontrada":     True,
            "docs_visita":    docs_visita,
            "conteo_media":   conteo_media,
            "ia":             ia_resultados,
            "gestor_nombre":  gestor_nombre,
            "gestor_cedula":  gestor_cedula,
            "gestor_ok":      gestor_ok,
            "observaciones":  obs,
            "tiene_error":   tiene_error,
        }


    # ── Validación carpeta 02 ─────────────────────────────────────────────────
    def _procesar_visita2(self, id_unico: str, link: str, dest_base: Path) -> dict:
        """
        Descarga selectivamente la carpeta 02_VISITA_2_DIAGNOSTICO
        y valida:
        - ACTA_VISITA_2
        - DIAGNOSTICO
        - PLAN_NEGOCIO
        """

        obs: list[str] = []
        tiene_error = False

        docs_visita2 = {
            k: "FALTA"
            for k in DOCS_VISITA2_ORDEN
        }

        acta_visita2 = "No analizado"
        diagnostico = "No analizado"
        plan_negocio = "No analizado"

        try:
            logger.info("  [%s] Buscando 02_VISITA...", id_unico)

            info = descargar_visita2_selectiva(
                link,
                dest_base,
            )

            if not info["encontrada"]:
                obs.append("02_VISITA_2_DIAGNOSTICO no encontrada")
                tiene_error = True

                return {
                    "encontrada": False,
                    "docs_visita2": docs_visita2,
                    "acta_visita_2": acta_visita2,
                    "diagnostico": diagnostico,
                    "plan_negocio": plan_negocio,
                    "observaciones": obs,
                    "tiene_error": tiene_error,
                }

            for keyword, ruta_archivo in info["docs_rutas"].items():

                if ruta_archivo is None:
                    docs_visita2[keyword] = "FALTA"
                    obs.append(f"02 — {keyword} no encontrado")
                    tiene_error = True
                    continue

                docs_visita2[keyword] = "OK"

                # ACTA_VISITA_2
                if keyword == "ACTA_VISITA_2":

                    if IA_HABILITADO:
                        try:
                            res = analizar_acta_visita(ruta_archivo)

                            if res.ok:
                                acta_visita2 = "OK"
                            else:
                                acta_visita2 = res.alerta
                                obs.append(
                                    f"Acta Visita 2: {res.alerta}"
                                )
                                tiene_error = True

                        except Exception as exc:
                            acta_visita2 = f"Error: {exc}"
                            obs.append(
                                f"Acta Visita 2: error — {exc}"
                            )
                            tiene_error = True


                # DIAGNOSTICO
                elif keyword == "DIAGNOSTICO":

                    if IA_HABILITADO:

                        try:

                            res = analizar_diagnostico(
                                str(ruta_archivo)
                            )

                            if res.ok:
                                diagnostico = "OK"

                            else:
                                diagnostico = res.alerta

                                obs.append(
                                    f"Diagnóstico: {res.alerta}"
                                )

                                tiene_error = True

                        except Exception as exc:

                            diagnostico = f"Error: {exc}"

                            obs.append(
                                f"Diagnóstico: error — {exc}"
                            )

                            tiene_error = True

                # PLAN_NEGOCIO
                elif keyword == "PLAN_NEGOCIO":

                    try:
                        resultado_plan = analizar_plan_negocio(
                            str(ruta_archivo)
                        )

                        if resultado_plan["completo"]:
                            plan_negocio = "OK"
                        else:
                            faltantes = len(
                                resultado_plan["faltantes"]
                            )

                            plan_negocio = (
                                f"Faltan {faltantes} campos"
                            )

                            obs.append(
                                f"Plan de Negocio incompleto "
                                f"({faltantes} campos)"
                            )

                            tiene_error = True

                    except Exception as exc:
                        plan_negocio = f"Error: {exc}"

                        obs.append(
                            f"Plan de Negocio: error — {exc}"
                        )

                        tiene_error = True

        except Exception as exc:
            obs.append(
                f"02 — Error procesando visita: {exc}"
            )
            tiene_error = True

        return {
            "encontrada": True,
            "docs_visita2": docs_visita2,
            "acta_visita_2": acta_visita2,
            "diagnostico": diagnostico,
            "plan_negocio": plan_negocio,
            "observaciones": obs,
            "tiene_error": tiene_error,
        }


    # ── Detección de documentos en 00 ────────────────────────────────────────

    @staticmethod
    def _listar_archivos(ruta: str) -> list:
        import os
        nombres: list[str] = []
        for _, _, archivos in os.walk(ruta):
            nombres.extend(archivos)
        return nombres

    @staticmethod
    def _detectar_documentos(archivos: list) -> Dict[str, bool]:
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
        id_unico: str, modalidad: str, unidad_doc: str, integrante: str,
        docs_estado: Dict[str, Estado],
        resultado_visita: dict,
        resultado_visita2: dict,
        observaciones: list,
        tiene_error: bool,
    ) -> dict:
        ia = resultado_visita.get("ia", {})
        dv = resultado_visita.get("docs_visita", {k: "N/A" for k in DOCS_VISITA_ORDEN})
        cm = resultado_visita.get("conteo_media", 0)
        encontrada = resultado_visita.get("encontrada", False)
        dv2 = resultado_visita2.get("docs_visita2", {})
        encontrada2 = resultado_visita2.get("encontrada", False)

        # ── 01_DOCUMENTOS: estado de existencia de los 3 docs ────────────────
        _NOMBRE_CORTO = {
            "ACTA_COMPROMISO":   "COMPROMISO",
            "ACTA_VISITA_1":     "VISITA",
            "TRATAMIENTO_DATOS": "TRATAMIENTO",
        }
        if not encontrada:
            docs_01_valor = "N/A"
        else:
            faltantes = [_NOMBRE_CORTO[k] for k in DOCS_VISITA_ORDEN if dv.get(k) == "FALTA"]
            presentes  = [_NOMBRE_CORTO[k] for k in DOCS_VISITA_ORDEN if dv.get(k) == "OK"]
            if faltantes:
                partes = [f"FALTA: {', '.join(faltantes)}"]
                if presentes:
                    partes.append(f"OK: {', '.join(presentes)}")
                docs_01_valor = " | ".join(partes)
            else:
                docs_01_valor = f"OK ({', '.join(presentes)})"

        # ── 01_FOTOS_VIDEOS ───────────────────────────────────────────────────
        if not encontrada:
            fotos_01_valor = "N/A"
        elif cm >= MIN_ARCHIVOS_MEDIA:
            fotos_01_valor = f"OK ({cm} archivos)"
        else:
            fotos_01_valor = f"FALTA ({cm}/{MIN_ARCHIVOS_MEDIA} mínimo)"

        # ── 02_DOCUMENTOS ────────────────────────────────────────────────────────────
        _NOMBRE_CORTO_02 = {
            "ACTA_VISITA_2": "ACTA",
            "DIAGNOSTICO": "DIAGNOSTICO",
            "PLAN_NEGOCIO": "PLAN",
        }

        if not encontrada2:
            docs_02_valor = "N/A"
        else:
            faltantes = [
                _NOMBRE_CORTO_02[k]
                for k in DOCS_VISITA2_ORDEN
                if dv2.get(k) == "FALTA"
            ]

            presentes = [
                _NOMBRE_CORTO_02[k]
                for k in DOCS_VISITA2_ORDEN
                if dv2.get(k) == "OK"
            ]

            if faltantes:
                partes = [f"FALTA: {', '.join(faltantes)}"]

                if presentes:
                    partes.append(f"OK: {', '.join(presentes)}")

                docs_02_valor = " | ".join(partes)
            else:
                docs_02_valor = f"OK ({', '.join(presentes)})"

        # ── Revisión por documento (columnas individuales) ────────────────────
        def _valor_revision(ia_key: str) -> tuple[str, bool]:
            """Retorna (texto_celda, tiene_alerta) para una columna de revisión."""
            if not IA_HABILITADO:
                return "No aplica (IA desactivada)", False
            if not encontrada:
                return "N/A", False
            res = ia.get(ia_key, _NO_ANALIZADO)
            if res is None or "No analizado" in res.alerta:
                return "No analizado", False
            if res.ok:
                return "OK", False
            return res.alerta, True

        val_compromiso,  alerta_compromiso  = _valor_revision("IA_COMPROMISO")
        val_visita,      alerta_visita      = _valor_revision("IA_VISITA")
        val_tratamiento, alerta_tratamiento = _valor_revision("IA_TRATAMIENTO")

        # ── Gestor ────────────────────────────────────────────────────────────
        gn = resultado_visita.get("gestor_nombre", "")
        gc = resultado_visita.get("gestor_cedula", "")
        gok = resultado_visita.get("gestor_ok", None)

        if not encontrada or not IA_HABILITADO:
            gestor_valor  = "N/A"
            alerta_gestor = False
        elif gok is None:
            gestor_valor  = "No extraído"
            alerta_gestor = False
        elif gok:
            gestor_valor  = f"OK — {gn} ({gc})"
            alerta_gestor = False
        else:
            gestor_valor  = f"NO REGISTRADO — {gn} ({gc})"
            alerta_gestor = True

        # ── Acta Visita 2 ─────────────────────────────────────────────────────
        if not encontrada2:
            acta2_valor = "N/A"
            alerta_acta2 = False
        else:
            acta2_valor = resultado_visita2.get(
                "acta_visita_2",
                "No analizado",
            )

            if acta2_valor in ("N/A", "No analizado"):
                alerta_acta2 = False
            else:
                alerta_acta2 = "OK" not in str(acta2_valor)
            
        # ── Diagnóstico ──────────────────────────────────────────────────────
        if not encontrada2:
            diagnostico_valor = "N/A"
            alerta_diagnostico = False

        else:
            diagnostico_valor = resultado_visita2.get(
                "diagnostico",
                "No analizado",
            )

            if diagnostico_valor in ("N/A", "No analizado"):
                alerta_diagnostico = False
            else:
                alerta_diagnostico = (
                    "OK" not in str(diagnostico_valor)
                )

        # ── Plan de Negocio ───────────────────────────────────────────────────
        if not encontrada2:
            plan_valor = "N/A"
            alerta_plan = False
        else:
            plan_valor = resultado_visita2.get(
                "plan_negocio",
                "No analizado",
            )

            if plan_valor in ("N/A", "No analizado"):
                alerta_plan = False
            else:
                alerta_plan = "OK" not in str(plan_valor)

        return {
            "ID_unico":        id_unico,
            "modalidad":       modalidad,
            "unidad_doc":      unidad_doc,
            "integrante":      integrante,
            # Carpeta 00
            "docs":            {doc: estado.value for doc, estado in docs_estado.items()},
            "docs_estado_raw": docs_estado,
            # Carpeta 01
            "01_documentos":      docs_01_valor,
            "01_fotos_videos":    fotos_01_valor,
            "01_acta_compromiso": val_compromiso,
            "01_acta_visita":     val_visita,
            "01_gestor":          gestor_valor,
            "01_tratamiento":     val_tratamiento,
            "01_alertas_por_doc": {
                "ACTA_COMPROMISO":   alerta_compromiso,
                "ACTA_VISITA":       alerta_visita,
                "GESTOR":            alerta_gestor,
                "TRATAMIENTO_DATOS": alerta_tratamiento,
            },
            # Carpeta 02
            "02_documentos": docs_02_valor,
            "02_acta_visita": acta2_valor,
            "02_diagnostico": diagnostico_valor,
            "02_plan_negocio": plan_valor,

            "02_alertas_por_doc": {
                "02_ACTA_VISITA": alerta_acta2,
                "02_DIAGNOSTICO": alerta_diagnostico,
                "02_PLAN_NEGOCIO": alerta_plan,
            },
            "observaciones":   "; ".join(observaciones),
            "_tiene_error":    tiene_error,
        }



    # ── Carga de la matriz ────────────────────────────────────────────────────

    @staticmethod
    def _cargar_matriz(ruta: str) -> pd.DataFrame:
        wb_tmp      = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
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
