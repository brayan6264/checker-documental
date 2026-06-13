# Validador Documental SharePoint

Sistema que recibe un archivo Excel (la "matriz"), descarga únicamente la carpeta de documentación de cada registro desde SharePoint, valida que estén los documentos obligatorios según la modalidad, y produce un **checklist en Excel** con el resultado.

---

## Arquitectura

```
validador_documental/
├── app/
│   ├── config.py                  # Variables de entorno y constantes
│   ├── main.py                    # FastAPI app (monta ambos routers)
│   │
│   ├── core/                      # Lógica de dominio pura
│   │   ├── normalizacion.py       # Normalización de texto (tildes, mayúsculas)
│   │   └── reglas.py              # Tabla de reglas por modalidad + evaluación
│   │
│   ├── services/                  # Lógica de negocio
│   │   ├── sharepoint.py          # Descarga selectiva de SharePoint
│   │   ├── procesador.py          # Motor de validación (orquesta todo)
│   │   └── checklist.py           # Escritura incremental del Excel de salida
│   │
│   ├── api/                       # Capa HTTP (routers FastAPI)
│   │   ├── descarga.py            # POST /descarga/descargar
│   │   └── validacion.py          # POST /validacion/validar, GET /estado, /resultado
│   │
│   └── ia/                        # Análisis semántico con IA (fase futura)
│       ├── extractor.py           # Extrae texto/imagen de documentos
│       └── validador.py           # Valida contenido con Claude (Anthropic)
│
├── scripts/
│   └── run.py                     # CLI sin servidor
├── server.py                      # Punto de entrada único
├── requirements.txt
└── .env.example
```

### Flujo por fila de la matriz

```
Excel (matriz)
    │
    ▼
procesador.py
    │
    ├─► sharepoint.py
    │       1. resolver()  →  sigue la URL compartida, obtiene cookies de sesión
    │       2. _buscar_carpeta_doc_en_sharepoint()  →  localiza 00_DOCUMENTACION
    │          en SharePoint SIN descargar el resto de la carpeta
    │       3. _descargar_arbol()  →  descarga en paralelo solo 00_DOCUMENTACION
    │
    ├─► Detecta documentos por nombre de archivo (normalizado)
    ├─► Evalúa obligatorios / opcionales / no aplica según modalidad
    ├─► Registra resultado en checklist.xlsx (incremental)
    └─► Borra la carpeta temporal descargada
```

**Por qué descarga selectiva:** las carpetas en SharePoint contienen subcarpetas adicionales (fotos de visita, caracterización, etc.) que pueden tener decenas de imágenes pesadas. Descargar solo `00_DOCUMENTACION` reduce el tiempo de ~5 minutos a segundos por registro.

---

## Reglas de validación

### Documentos detectados por nombre de archivo

| Documento | Palabra clave en el nombre |
|-----------|---------------------------|
| CEDULA    | `cedula`                  |
| COMERCIO  | `comercio`                |
| RUT       | `rut`                     |
| TENENCIA  | `tenencia`                |

La comparación es normalizada: sin tildes, sin distinción de mayúsculas/minúsculas.

### Matriz documento × modalidad

| Documento | M1          | M2          | M3          | M4          |
|-----------|-------------|-------------|-------------|-------------|
| CEDULA    | Obligatorio | Obligatorio | Obligatorio | Obligatorio |
| COMERCIO  | No aplica   | Opcional    | Obligatorio | Obligatorio |
| RUT       | Opcional    | Opcional    | Obligatorio | Obligatorio |
| TENENCIA  | Opcional    | Obligatorio | Obligatorio | Obligatorio |

### Estados del checklist

| Estado             | Significado                              | Color en Excel |
|--------------------|------------------------------------------|----------------|
| `OK`               | Obligatorio y presente                   | —              |
| `FALTA (obligatorio)` | Obligatorio y ausente               | Rojo           |
| `Presente`         | Opcional y presente                      | —              |
| `Ausente`          | Opcional y ausente                       | —              |
| `N/A`              | No aplica para esta modalidad            | —              |

---

## Instalación

```bash
# 1. Clonar / descomprimir el proyecto
cd validador_documental

# 2. Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # Linux/Mac

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Configurar variables de entorno (opcional)
copy .env.example .env
# Editar .env si se necesita cambiar rutas o habilitar IA
```

---

## Arranque

```bash
python server.py
```

El servidor queda disponible en `http://127.0.0.1:8000`.

**Opciones:**
```bash
python server.py --reload          # recarga automática al cambiar código (desarrollo)
python server.py --port 9000       # puerto personalizado
```

> No se necesita levantar ningún segundo proceso. Descarga y validación corren en el mismo servidor.

---

## Probar desde Swagger

### 1. Abrir Swagger UI

Navegar a:
```
http://127.0.0.1:8000/docs
```

---

### 2. Verificar que el servidor está vivo

**`GET /health`**

1. Hacer clic en `GET /health` → **Try it out** → **Execute**
2. Respuesta esperada:
```json
{
  "status": "ok",
  "version": "2.0.0"
}
```

---

### 3. Lanzar una validación

**`POST /validacion/validar`**

1. Hacer clic en `POST /validacion/validar` → **Try it out**
2. En el campo `archivo` hacer clic en **Choose File** y seleccionar el Excel de la matriz
3. Hacer clic en **Execute**
4. La respuesta llega inmediatamente (el proceso corre en background):

```json
{
  "job_id": "abc123...",
  "estado": "iniciado",
  "estado_url": "/validacion/estado/abc123...",
  "resultado_url": "/validacion/resultado/abc123..."
}
```

> Copiar el `job_id` para los siguientes pasos.

---

### 4. Consultar el progreso

**`GET /validacion/estado/{job_id}`**

1. Hacer clic en `GET /validacion/estado/{job_id}` → **Try it out**
2. Pegar el `job_id` copiado en el paso anterior
3. Hacer clic en **Execute**
4. Respuesta durante el proceso:

```json
{
  "estado": "procesando",
  "filas_total": 4,
  "filas_procesadas": 2,
  "errores": 0,
  "ruta_checklist": "jobs/abc123.../checklist.xlsx"
}
```

5. Cuando `estado` sea `"completado"`, el checklist está listo.

**Estados posibles:** `iniciado` → `procesando` → `completado` / `error`

---

### 5. Descargar el checklist

**`GET /validacion/resultado/{job_id}`**

1. Hacer clic en `GET /validacion/resultado/{job_id}` → **Try it out**
2. Pegar el `job_id`
3. Hacer clic en **Execute**
4. Hacer clic en **Download file** para guardar el Excel `checklist_validacion.xlsx`

---

### 6. (Opcional) Probar solo la descarga de una carpeta

**`POST /descarga/descargar`**

Útil para verificar que la conexión con SharePoint funciona antes de procesar la matriz completa.

1. Hacer clic en `POST /descarga/descargar` → **Try it out**
2. Pegar el body:
```json
{
  "url": "https://...sharepoint.com/:f:/s/..."
}
```
3. Respuesta:
```json
{
  "ruta_local": "C:\\...\\descargas_sharepoint\\uuid\\00_DOCUMENTACION",
  "total": 8,
  "exitosos": 8,
  "fallidos": 0
}
```

---

## Columnas del Excel de entrada (matriz)

El archivo Excel debe tener una hoja activa con **exactamente** estas columnas:

| Columna               | Requerida | Descripción                          |
|-----------------------|-----------|--------------------------------------|
| `ID_unico`            | Sí        | Identificador único del registro     |
| `modalidad`           | Sí        | M1, M2, M3 o M4                     |
| `unidad_doc`          | Sí        | Unidad documental                    |
| `carpetas_link`       | Sí        | URL anónima de la carpeta SharePoint |
| `integrante_nombre1`  | No        | Primer nombre del integrante         |
| `integrante_nombre2`  | No        | Segundo nombre                       |
| `integrante_apellido1`| No        | Primer apellido                      |
| `integrante_apellido2`| No        | Segundo apellido                     |

---

## Variables de entorno

Copiar `.env.example` a `.env` y ajustar según el entorno:

| Variable                | Default                    | Descripción                                  |
|-------------------------|----------------------------|----------------------------------------------|
| `JOBS_DIR`              | `jobs`                     | Directorio donde se guardan los trabajos     |
| `DOWNLOADS_DIR`         | `descargas_sharepoint`     | Directorio temporal de descargas             |
| `MAX_WORKERS_LISTADO`   | `8`                        | Hilos paralelos para listar carpetas         |
| `MAX_WORKERS_DESCARGA`  | `12`                       | Hilos paralelos para descargar archivos      |
| `CHECKLIST_LOTE_GUARDADO` | `50`                     | Filas antes de persistir el checklist a disco|
| `IA_HABILITADO`         | `false`                    | Activa validación semántica con IA           |
| `ANTHROPIC_API_KEY`     | *(vacío)*                  | API key de Anthropic (solo si IA_HABILITADO) |
| `IA_MODEL`              | `claude-haiku-4-5-20251001`| Modelo a usar para análisis de documentos    |

---

## Módulo de IA (fase futura)

El módulo `app/ia/` está preparado pero desactivado por defecto (`IA_HABILITADO=false`).

Cuando se active, complementa la validación por nombre de archivo con validación **semántica**:
- **Fase 1 (activa):** ¿existe un archivo cuyo nombre contiene "cedula"?
- **Fase 2 (IA):** ¿el contenido de ese archivo ES realmente una cédula?

Para activarlo:
```env
IA_HABILITADO=true
ANTHROPIC_API_KEY=sk-ant-...
```

Instalar dependencias adicionales según los formatos a analizar:
```bash
pip install anthropic pymupdf pillow
```

---

## Ejecución por CLI (sin servidor)

Para procesar un Excel directamente desde la terminal:

```bash
python scripts/run.py ruta/matriz.xlsx ruta/checklist.xlsx
```

El proceso es **reanudable**: si se interrumpe, al volver a ejecutar con el mismo archivo de salida continúa desde donde quedó, saltando los registros ya procesados.
