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

## Uso de la API desde Swagger UI

Swagger UI es la interfaz visual integrada que permite ejecutar todos los endpoints sin escribir código ni usar Postman.

### Flujo completo de una validación

```
 PASO 1          PASO 2              PASO 3            PASO 4          PASO 5
Abrir Swagger → Subir matriz.xlsx → Copiar job_id → Consultar estado → Descargar checklist
   /docs         POST /validar        (respuesta)     GET /estado/{id}   GET /resultado/{id}
```

---

### PASO 1 — Abrir Swagger UI

Con el servidor corriendo (`python server.py`), abrir en el navegador:

```
http://127.0.0.1:8000/docs
```

Verás una pantalla similar a esta:

```
┌─────────────────────────────────────────────────────────────────┐
│  Validador Documental SharePoint   OAS 3.0                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ▼  validacion                                                  │
│  ├── POST  /validacion/validar      Lanzar validación           │
│  ├── GET   /validacion/estado/{id}  Consultar progreso          │
│  └── GET   /validacion/resultado/{id} Descargar checklist       │
│                                                                 │
│  ▼  descarga                                                    │
│  └── POST  /descarga/descargar      Probar descarga SharePoint  │
│                                                                 │
│  ▼  default                                                     │
│  └── GET   /health                  Estado del servidor         │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

> **Tip:** Antes de todo, verificar que el servidor responde ejecutando `GET /health`. Debe retornar `{"status": "ok"}`.

---

### PASO 2 — Lanzar una validación

**`POST /validacion/validar`**

```
┌─────────────────────────────────────────────────────────────────┐
│ POST /validacion/validar                            [ Try it out]│
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Request body     multipart/form-data                           │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ archivo  [  matriz.xlsx          ] [Choose File]          │  │
│  │                                                           │  │
│  │ flujos   [                       ]  (opcional)            │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│                                          [ Clear ]  [ Execute ] │
└─────────────────────────────────────────────────────────────────┘
```

**Pasos:**
1. Clic en `POST /validacion/validar` para expandir el endpoint
2. Clic en **Try it out** (esquina superior derecha del endpoint)
3. Clic en **Choose File** → seleccionar el archivo Excel de la matriz
4. En el campo `flujos` escribir los flujos que deseas ejecutar (ver tabla abajo)
5. Clic en **Execute**

**Campo `flujos` — selección de carpetas a validar:**

| Valor en `flujos` | Carpeta validada | Qué revisa |
|---|---|---|
| `00` | `00_DOCUMENTACION` | Cédula, RUT, Comercio, Tenencia |
| `01` | `01_VISITA_1_CARACTERIZACION` | Acta compromiso, acta visita, fotos, gestor, tratamiento datos |
| `02` | `02_VISITA_2_DIAGNOSTICO` | Acta visita 2, diagnóstico, plan de negocio |
| `03` | `03_CAPACITACION` | Encuestas, grupal, individual, módulos TX/RX |

**Ejemplos de combinaciones:**

| Qué ejecutar | Valor del campo `flujos` |
|---|---|
| Todo el flujo completo | *(dejar vacío)* |
| Solo documentación | `00` |
| Solo primera visita | `01` |
| Solo segunda visita | `02` |
| Solo capacitación | `03` |
| Documentación + primera visita | `00,01` |
| Las dos visitas | `01,02` |
| Todo excepto capacitación | `00,01,02` |
| Capacitación + documentación | `00,03` |

Las columnas de las carpetas no ejecutadas aparecerán como `—` en el checklist de salida.

**Respuesta esperada (HTTP 200):**
```json
{
  "job_id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "estado": "iniciado",
  "flujos": ["00", "01", "02", "03"],
  "estado_url": "/validacion/estado/a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "resultado_url": "/validacion/resultado/a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

> ⚠️ **Importante:** el proceso corre en **background** — la respuesta llega en menos de 1 segundo aunque la validación tarde varios minutos. **Copiar el `job_id`** para los pasos siguientes.

---

### PASO 3 — Consultar el progreso

**`GET /validacion/estado/{job_id}`**

```
┌─────────────────────────────────────────────────────────────────┐
│ GET /validacion/estado/{job_id}                     [ Try it out]│
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Parameters                                                     │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ job_id *  a1b2c3d4-e5f6-7890-abcd-ef1234567890           │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│                                          [ Clear ]  [ Execute ] │
└─────────────────────────────────────────────────────────────────┘
```

**Pasos:**
1. Clic en `GET /validacion/estado/{job_id}` para expandir
2. Clic en **Try it out**
3. En el campo `job_id` pegar el ID copiado en el Paso 2
4. Clic en **Execute**
5. **Repetir hasta que `estado` sea `"completado"`**

**Respuesta mientras procesa:**
```json
{
  "job_id": "a1b2c3d4-...",
  "estado": "procesando",
  "filas_total": 120,
  "filas_procesadas": 47,
  "errores": 2,
  "inicio_utc": "2025-06-13T14:32:10Z",
  "duracion_segundos": 83.4
}
```

**Respuesta al finalizar:**
```json
{
  "job_id": "a1b2c3d4-...",
  "estado": "completado",
  "filas_total": 120,
  "filas_procesadas": 120,
  "errores": 5,
  "inicio_utc": "2025-06-13T14:32:10Z",
  "duracion_segundos": 312.7
}
```

**Ciclo de vida del estado:**

```
  [iniciado] ──► [procesando] ──► [completado]
                      │
                      └──────────► [error]  ← si falla algo crítico
```

| Estado       | Significado                                        |
|--------------|----------------------------------------------------|
| `iniciado`   | Job registrado, aún no comenzó a procesar filas    |
| `procesando` | Validando filas, ver `filas_procesadas` para avance|
| `completado` | Todas las filas procesadas, checklist disponible   |
| `error`      | Falla crítica (archivo inválido, SharePoint caído) |

---

### PASO 4 — Descargar el checklist

**`GET /validacion/resultado/{job_id}`**

```
┌─────────────────────────────────────────────────────────────────┐
│ GET /validacion/resultado/{job_id}                  [ Try it out]│
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Parameters                                                     │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ job_id *  a1b2c3d4-e5f6-7890-abcd-ef1234567890           │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│                                          [ Clear ]  [ Execute ] │
├─────────────────────────────────────────────────────────────────┤
│ Response body                                                   │
│                                                                 │
│  [Download file]  ← hacer clic aquí para guardar el Excel      │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

**Pasos:**
1. Clic en `GET /validacion/resultado/{job_id}` para expandir
2. Clic en **Try it out**
3. Pegar el `job_id`
4. Clic en **Execute**
5. En la sección **Response body**, hacer clic en **Download file**
6. Guardar el archivo `checklist_validacion.xlsx`

> Solo disponible cuando el estado es `"completado"`. Si se llama antes, retorna HTTP 404.

---

### (Opcional) Probar la conexión con SharePoint antes de procesar

**`POST /descarga/descargar`**

Útil para verificar que una URL de SharePoint es accesible y los archivos se descargan correctamente, **antes** de correr la validación completa.

**Pasos:**
1. Clic en `POST /descarga/descargar` → **Try it out**
2. En el campo `Request body`, pegar:
```json
{
  "url": "https://tuorganizacion.sharepoint.com/:f:/s/Sitio/XXXX"
}
```
3. Clic en **Execute**

**Respuesta esperada:**
```json
{
  "ruta_local": "C:\\...\\descargas_sharepoint\\uuid\\00_DOCUMENTACION",
  "total": 8,
  "exitosos": 8,
  "fallidos": 0
}
```

Si `fallidos > 0` o aparece error 401/403, la URL de SharePoint no tiene permisos de acceso anónimo.

---

### Errores comunes en Swagger

| Error HTTP | Causa más probable | Qué hacer |
|---|---|---|
| `422 Unprocessable Entity` | El Excel tiene columnas faltantes o nombre incorrecto | Verificar que las columnas coincidan con la tabla de [Columnas del Excel de entrada](#columnas-del-excel-de-entrada-matriz) |
| `404 Not Found` en `/resultado` | El job aún no terminó | Esperar a que `/estado` devuelva `"completado"` |
| `500 Internal Server Error` | Error en el servidor | Revisar la consola donde corre `python server.py` |
| `0 filas procesadas` en `/estado` | El Excel está vacío o la primera fila no es de datos | Verificar que la hoja activa tenga datos desde la fila 2 |

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
| `IA_HABILITADO`         | `false`                    | Activa validación semántica con IA (OpenAI)  |
| `OPENAI_API_KEY`        | *(vacío)*                  | API key de OpenAI (requerida si IA_HABILITADO=true) |
| `OPENAI_MODEL`          | `gpt-4o`                   | Modelo OpenAI para análisis visual de documentos |

---

## Módulo de IA

El análisis con IA está desactivado por defecto (`IA_HABILITADO=false`). Cuando se activa, complementa la validación por nombre de archivo con revisión del **contenido** de los documentos usando GPT-4o con visión.

**Qué analiza la IA por carpeta:**

| Carpeta | Documento | Qué verifica |
|---------|-----------|--------------|
| 01 Visita 1 | Acta de Compromiso | Firmas, fechas, datos del gestor |
| 01 Visita 1 | Acta de Visita | Extrae nombre y cédula del gestor para verificar en BD |
| 01 Visita 1 | Tratamiento de Datos | Autorización firmada presente |
| 02 Visita 2 | Acta de Visita 2 | Contenido y firmas |
| 02 Visita 2 | Diagnóstico | Documento completo y coherente |
| 02 Visita 2 | Plan de Negocio | Campos obligatorios completados |
| 03 Capacitación | TX_RX por módulo | Nombres de encuestas presentes en lista de asistencia |

**Para activarla:**
```env
IA_HABILITADO=true
OPENAI_API_KEY=sk-...tu-key-de-openai...
OPENAI_MODEL=gpt-4o
```

Las dependencias ya están en `requirements.txt`. Si faltaran:
```bash
pip install openai pymupdf pillow
```

---

## Ejecución por CLI (sin servidor)

Para procesar un Excel directamente desde la terminal:

```bash
python scripts/run.py ruta/matriz.xlsx ruta/checklist.xlsx
```

El proceso es **reanudable**: si se interrumpe, al volver a ejecutar con el mismo archivo de salida continúa desde donde quedó, saltando los registros ya procesados.
