"""
Script CLI para ejecutar la validación sin levantar el servidor.

Uso:
    python scripts/run.py <matriz.xlsx> [checklist.xlsx]

Argumentos:
    matriz.xlsx     (obligatorio) Ruta del Excel de entrada.
    checklist.xlsx  (opcional)    Ruta de salida. Default: checklist_validacion.xlsx

Ejemplo:
    python scripts/run.py datos/matriz_2026.xlsx resultados/checklist.xlsx
"""

import logging
import sys
from pathlib import Path

# Asegurar que el directorio raíz del proyecto esté en el path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.services.procesador import ValidadorDocumental

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)


def main() -> None:
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)

    ruta_excel     = args[0]
    ruta_checklist = args[1] if len(args) > 1 else "checklist_validacion.xlsx"

    print(f"Matriz:    {ruta_excel}")
    print(f"Checklist: {ruta_checklist}")
    print("─" * 50)

    def _progreso(procesadas: int, total: int, errores: int) -> None:
        print(f"\r  {procesadas}/{total} filas  |  {errores} alertas", end="", flush=True)

    validador = ValidadorDocumental(
        ruta_checklist=ruta_checklist,
        callback_progreso=_progreso,
    )

    try:
        validador.procesar_matriz(ruta_excel)
        print(f"\nChecklist generado en: {ruta_checklist}")
    except Exception as exc:
        print(f"\nError fatal: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
