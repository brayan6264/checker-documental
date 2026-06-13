"""
Escritura incremental del checklist de validación documental en formato Excel.

Características:
- Si el archivo ya existe, lo carga y extrae los ID_unico ya procesados (reanudable).
- Cada llamada a agregar_fila() acumula en memoria y persiste cada `lote_guardado` filas.
- Aplica el formato visual especificado: Arial 9, centrado, encabezado lila,
  celdas rojas para documentos obligatorios ausentes.
"""

import os
from typing import Dict, Set

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from app.core.reglas import Estado

# ── Paleta de colores (ARGB: FF = opaco) ─────────────────────────────────────
_LILA  = "FFCC8FCC"
_ROJO  = "FFFF6B6B"
_BLANC = "FFFFFFFF"

FILL_LILA = PatternFill(start_color=_LILA, end_color=_LILA, fill_type="solid")
FILL_ROJO = PatternFill(start_color=_ROJO, end_color=_ROJO, fill_type="solid")

FONT_HEADER = Font(name="Arial", size=9, bold=True, color=_BLANC)
FONT_NORMAL = Font(name="Arial", size=9)
FONT_FALTA  = Font(name="Arial", size=9, bold=True, color=_BLANC)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ── Estructura de columnas ────────────────────────────────────────────────────

COLUMNAS = [
    "ID_unico",
    "modalidad",
    "unidad_doc",
    "integrante",
    "CEDULA",
    "COMERCIO",
    "RUT",
    "TENENCIA",
    "observaciones",
]

_ANCHOS_MIN = {
    "ID_unico":      20,
    "modalidad":     12,
    "unidad_doc":    22,
    "integrante":    32,
    "CEDULA":        20,
    "COMERCIO":      20,
    "RUT":           20,
    "TENENCIA":      20,
    "observaciones": 42,
}

_DOC_COL_IDX = {
    doc: COLUMNAS.index(doc) + 1
    for doc in ("CEDULA", "COMERCIO", "RUT", "TENENCIA")
}


class ChecklistWriter:
    """Mantiene y escribe el archivo Excel del checklist de validación."""

    def __init__(self, ruta: str, lote_guardado: int = 50):
        self.ruta         = ruta
        self.lote_guardado = lote_guardado
        self._sin_guardar  = 0
        self._max_ancho    = [_ANCHOS_MIN[c] for c in COLUMNAS]

        if os.path.exists(ruta):
            self._wb = openpyxl.load_workbook(ruta)
            self._ws = self._wb.active
            self._ids_procesados: Set[str] = self._leer_ids_existentes()
        else:
            self._wb = openpyxl.Workbook()
            self._ws = self._wb.active
            self._ws.title = "Checklist"
            self._ids_procesados = set()
            self._escribir_encabezado()

    def ids_procesados(self) -> Set[str]:
        return self._ids_procesados

    def agregar_fila(self, resultado: dict) -> None:
        """Agrega una fila y aplica formato. Persiste cada `lote_guardado` filas."""
        docs_estado: Dict[str, Estado] = resultado.get("docs_estado_raw", {})

        valores = [
            resultado.get("ID_unico",     ""),
            resultado.get("modalidad",    ""),
            resultado.get("unidad_doc",   ""),
            resultado.get("integrante",   ""),
            resultado.get("docs", {}).get("CEDULA",   ""),
            resultado.get("docs", {}).get("COMERCIO", ""),
            resultado.get("docs", {}).get("RUT",      ""),
            resultado.get("docs", {}).get("TENENCIA", ""),
            resultado.get("observaciones", ""),
        ]

        self._ws.append(valores)
        fila_num = self._ws.max_row

        for col_idx in range(1, len(COLUMNAS) + 1):
            celda           = self._ws.cell(row=fila_num, column=col_idx)
            celda.font      = FONT_NORMAL
            celda.alignment = ALIGN_CENTER

        for doc, col_idx in _DOC_COL_IDX.items():
            if docs_estado.get(doc) == Estado.FALTA:
                celda       = self._ws.cell(row=fila_num, column=col_idx)
                celda.fill  = FILL_ROJO
                celda.font  = FONT_FALTA

        for i, valor in enumerate(valores):
            longitud = len(str(valor)) if valor else 0
            if longitud + 3 > self._max_ancho[i]:
                self._max_ancho[i] = longitud + 3

        self._ids_procesados.add(str(resultado.get("ID_unico", "")))
        self._sin_guardar += 1

        if self._sin_guardar >= self.lote_guardado:
            self.guardar()

    def guardar(self) -> None:
        """Persiste a disco. Aplica auto-fit de columnas antes de escribir."""
        self._ajustar_anchos()
        self._wb.save(self.ruta)
        self._sin_guardar = 0

    def _escribir_encabezado(self) -> None:
        self._ws.append(COLUMNAS)
        for col_idx in range(1, len(COLUMNAS) + 1):
            celda           = self._ws.cell(row=1, column=col_idx)
            celda.fill      = FILL_LILA
            celda.font      = FONT_HEADER
            celda.alignment = ALIGN_CENTER
        self._ws.row_dimensions[1].height = 25
        self._ws.freeze_panes = "A2"

    def _leer_ids_existentes(self) -> Set[str]:
        ids: Set[str] = set()
        for fila in self._ws.iter_rows(min_row=2, values_only=True):
            if fila[0]:
                ids.add(str(fila[0]))
        return ids

    def _ajustar_anchos(self) -> None:
        for col_idx, ancho in enumerate(self._max_ancho, start=1):
            letra = get_column_letter(col_idx)
            self._ws.column_dimensions[letra].width = min(ancho, 60)
