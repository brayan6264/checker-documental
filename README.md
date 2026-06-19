# Validador Documental SharePoint

Sistema que recibe un archivo Excel (la "matriz"), descarga selectivamente las carpetas de cada registro desde SharePoint, valida los documentos obligatorios y el contenido según la modalidad, y produce un **checklist en Excel** con el resultado.

Cubre **5 flujos** de validación (00 a 04). El flujo **04 (Capitalización)** es el más completo: valida el PLAN_INVERSION.xlsx, las firmas, **cruza las 3 cotizaciones de cada producto contra los PDFs de la carpeta `02_COTIZACIONES_Y_COMPRA`**, y hace una **búsqueda web de precios de referencia** (Playwright + OpenAI) para detectar sobrecostos frente al mercado.

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
│   │   ├── reglas.py              # Tabla de reglas por modalidad + evaluación
│   │   ├── plan_inversion.py      # Validación PLAN_INVERSION.xlsx, firmas y
│   │   │                          #   cotizaciones vs PDFs (carpeta 02 + OCR + GPT)
│   │   └── cotizacion_web.py      # Búsqueda web de precios (Playwright + OpenAI)
│   │
│   ├── services/                  # Lógica de negocio
│   │   ├── sharepoint.py          # Descarga selectiva de SharePoint
│   │   ├── procesador.py          # Motor de validación (orquesta todos los flujos)
│   │   └── checklist.py           # Escritura incremental del Excel de salida
│   │
│   ├── api/                       # Capa HTTP (routers FastAPI)
│   │   ├── descarga.py            # POST /descarga/descargar
│   │   └── validacion.py          # POST /validacion/validar, GET /estado, /resultado
│   │
│   └── ia/                        # Análisis semántico con IA
│       ├── extractor.py           # Extrae texto/imagen (texto nativo + OCR + fallback GPT)
│       ├── validador.py           # Valida contenido con Claude (Anthropic)
│       ├── analizador_visita.py   # Análisis visual GPT-4o de visita 1
│       └── analizador_visita2.py  # Análisis visual GPT-4o de visita 2
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
# 1. Clonar el repositorio
git clone <URL_DEL_REPOSITORIO>
cd validador_documental

# 2. Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # Linux/Mac

# 3. Instalar dependencias de Python
pip install -r requirements.txt

# 4. Instalar el navegador para Playwright (REQUERIDO para el flujo 04 web)
playwright install chromium

# 5. Configurar variables de entorno
copy .env.example .env          # Windows  (Linux/Mac: cp .env.example .env)
# Editar .env: como mínimo agregar OPENAI_API_KEY para el flujo 04
```

### Dependencias del sistema (no son paquetes pip)

| Dependencia | Para qué | Cómo instalar |
|-------------|----------|---------------|
| **Navegador Chromium de Playwright** | Capturas de pantalla en la búsqueda web (flujo 04) | `playwright install chromium` |
| **Tesseract OCR** | OCR de PDFs escaneados (flujo 04 — firmas/plan/cotizaciones) | Windows: instalar en `C:\Program Files\Tesseract-OCR\` con el paquete de idioma español (`spa`). Linux: `sudo apt install tesseract-ocr tesseract-ocr-spa` |

> La ruta de Tesseract en Windows está fijada en `app/ia/extractor.py` y `app/core/plan_inversion.py` como `C:\Program Files\Tesseract-OCR\tesseract.exe`. Si lo instalas en otra ruta, ajústala ahí.

> El flujo 04 (búsqueda web y OCR con fallback a IA) requiere `OPENAI_API_KEY`. Los flujos 00–03 funcionan sin OpenAI salvo que actives `IA_HABILITADO=true`.

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
| `04` | `04_CAPITALIZACION` | PLAN_INVERSION.xlsx, firmas, cotizaciones vs carpeta 02, precios de mercado (web) |

**Ejemplos de combinaciones:**

| Qué ejecutar | Valor del campo `flujos` |
|---|---|
| Todo el flujo completo | *(dejar vacío)* |
| Solo documentación | `00` |
| Solo primera visita | `01` |
| Solo segunda visita | `02` |
| Solo capacitación | `03` |
| Solo capitalización | `04` |
| Documentación + primera visita | `00,01` |
| Las dos visitas | `01,02` |
| Todo excepto capacitación | `00,01,02,04` |
| Solo cotizaciones/precios | `04` |

Las columnas de las carpetas no ejecutadas aparecerán como `—` en el checklist de salida.

**Respuesta esperada (HTTP 200):**
```json
{
  "job_id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "estado": "iniciado",
  "flujos": ["00", "01", "02", "03", "04"],
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
| `Departamento`        | No        | Departamento (flujo 04): si el producto es un ser vivo (plantas, semillas, animales), la búsqueda web se restringe a este departamento |
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
| `IA_HABILITADO`         | `false`                    | Activa validación semántica con IA (flujos 01/02) |
| `OPENAI_API_KEY`        | *(vacío)*                  | API key de OpenAI. **Requerida para el flujo 04** (búsqueda web + OCR fallback) y si `IA_HABILITADO=true` |
| `OPENAI_MODEL`          | `gpt-4o`                   | Modelo OpenAI para análisis visual de documentos (flujos 01/02) |
| `TIMEOUT_DESCARGA`      | `120`                      | Timeout (s) por archivo descargado de SharePoint |

> Los modelos del flujo 04 están fijados en `app/core/cotizacion_web.py`: `gpt-4o` para búsqueda web y análisis visual, `gpt-4o-mini` para normalizar términos. El OCR fallback usa `gpt-4o-mini`.

---

## Flujo 04 — Capitalización (PLAN_INVERSION y cotizaciones)

Es el flujo más completo. Se ejecuta con `flujos=04` (o vacío para todo). Descarga la carpeta `04_CAPITALIZACION` y procesa, en orden:

```
04_CAPITALIZACION
├── 01_APROBACION_Y_PLAN
│   ├── PLAN_INVERSION_*.xlsx   ← se valida estructura y cotizaciones
│   └── FIRMA_UP_*.pdf / PLAN..pdf  ← firmas + cruce de cotización ganadora
└── 02_COTIZACIONES_Y_COMPRA
    ├── COTIZACION GANADORA/  → PDF del proveedor (ej. ...-ANGELNET.pdf)
    ├── COTIZACION TED CEL/   → PDF del proveedor
    └── COTIZACION WARRIORS/  → PDF del proveedor
```

**Pasos del flujo 04:**

1. **PLAN_INVERSION.xlsx** — valida que cada ítem tenga sus 3 cotizaciones completas (proveedor, descripción, valor total) y detecta la cotización seleccionada (la marcada con `X`).

2. **Firmas** — cuenta las imágenes de firma embebidas en el PDF del plan (mínimo 2).

3. **Cruce de la cotización ganadora vs PDF del plan** — verifica que el valor total de la cotización seleccionada aparezca en el PDF. Para PDFs escaneados usa un pipeline de 3 capas: texto nativo → OCR (Tesseract) → GPT-4o-mini como segunda opinión.

4. **Cotizaciones vs carpeta `02_COTIZACIONES_Y_COMPRA`** *(verificación nueva)*:
   - Localiza el PDF de **cada proveedor** por **coincidencia parcial** del nombre de la empresa. El nombre del Excel (fila "Nombre proveedor") no es idéntico al del PDF: `ANGELNET INGENIERIA SAS` → busca `angelnet`. Se ignoran palabras genéricas (SAS, INGENIERIA, CIA, etc.).
   - El texto se extrae **nativo, sin OCR ni IA** (los PDFs de cotización tienen caracteres).
   - Valida que el valor total de **las 3 cotizaciones de cada producto** coincida con el precio del PDF del proveedor correspondiente.
   - **Si falta el PDF de alguna cotización → se DETIENE el proceso** (no se hace búsqueda web).

5. **Búsqueda web de precios de referencia** *(solo si los pasos anteriores pasan)*:
   - OpenAI (`web_search_preview`, obligatorio vía `tool_choice`) busca URLs reales de tiendas colombianas para cada producto seleccionado. Excluye MercadoLibre.
   - Playwright visita las URLs, toma capturas y extrae el precio del DOM.
   - GPT-4o (visión) recibe las capturas como **imágenes** y valida que muestren el producto correcto y con precio.
   - **Regla obligatoria: 3 capturas válidas por producto.** Si faltan, se hacen rondas de reemplazo (hasta 5). Si aún así no llega a 3 → se marca *completación manual*.
   - Genera un Excel `{ID}.xlsx` (una hoja por producto con capturas + precios) y compara el precio cotizado contra la **mediana** de internet: alerta si supera 50 % (productos < 500.000 COP) o 20 % (≥ 500.000 COP).

### Columnas del flujo 04 en el checklist de salida

| Columna | Qué indica | Colores |
|---------|------------|---------|
| `04_PLAN_INVERSION_ENCONTRADO` | Estructura del PLAN_INVERSION.xlsx | 🔴 si falta/incompleto |
| `04_PDF_APROBACION_ENCONTRADO` | PDF del plan presente | 🔴 si falta |
| `04_PDF_APROBACION_FIRMADO` | Firmas encontradas en el PDF | 🔴 si insuficientes |
| `04_CONSISTENCIA_COTIZACION_GANADORA` | Valor de la cotización ganadora coincide con el PDF del plan | 🔴 si no coincide |
| `04_COTIZACIONES (carpeta 02)` | Las 3 cotizaciones existen y sus precios coinciden con los PDFs de la carpeta 02 | 🟡 si hay algo que revisar |
| `04_VALIDACION_PRECIO_MERCADO` | Resultado de la búsqueda web (Excel generado / completación manual) | 🟡 si faltan capturas (completar manual) |

**Código de colores general del checklist:**

| Color | Significado |
|-------|-------------|
| 🔴 Rojo | Documento/dato faltante o error crítico |
| 🟡 Amarillo | Requiere revisión o completación manual |
| 🟠 Naranja | Alerta de desviación de precio de mercado (en la columna `observaciones`) |

> El detalle completo de cada alerta (producto, proveedor, precio) siempre queda en la columna `observaciones`.

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

El CLI ejecuta **todos los flujos (00–04)**; no permite seleccionar flujos (para eso usar la API). Requiere las mismas dependencias del sistema que el flujo 04: `OPENAI_API_KEY` en `.env`, navegador de Playwright (`playwright install chromium`) y Tesseract OCR.

El proceso es **reanudable**: si se interrumpe, al volver a ejecutar con el mismo archivo de salida continúa desde donde quedó, saltando los registros ya procesados.

> Para que aparezcan las columnas nuevas/renombradas del flujo 04, usa un archivo de checklist **nuevo** (vacío). `ChecklistWriter` solo escribe los encabezados al crear el archivo; si reutilizas uno viejo, conserva las columnas anteriores.
