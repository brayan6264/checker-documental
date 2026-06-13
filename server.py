"""
Punto de entrada único. Levanta descarga + validación en el mismo proceso.

Uso:
    python server.py              # puerto 8000, sin reload
    python server.py --reload     # con recarga automática (desarrollo)
    python server.py --port 9000  # puerto personalizado
"""

import sys
import uvicorn

if __name__ == "__main__":
    reload = "--reload" in sys.argv
    port   = 8000
    for arg in sys.argv:
        if arg.startswith("--port="):
            port = int(arg.split("=")[1])
        elif arg == "--port" and sys.argv.index(arg) + 1 < len(sys.argv):
            port = int(sys.argv[sys.argv.index(arg) + 1])

    uvicorn.run(
        "app.main:app",
        host="127.0.0.1",
        port=port,
        reload=reload,
    )
