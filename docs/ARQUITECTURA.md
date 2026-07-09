# Arquitectura técnica

Documento para desarrolladores que necesiten **mantener, depurar o extender** el sistema.

> Si eres consultor SST sin formación técnica, esto NO es para ti. Lee `GUIA_CONSULTOR.md`.

---

## Diagrama general

```
┌─────────────────────────────────────────────────────────────────┐
│  USUARIO (Consultor SST)          │  CLIENTE HTTP (Fase 4)        │
│  • Edita: clientes_config/X.yaml  │  • POST /clientes/{id}/generar│
│  • Ejecuta: python cli.py ...     │  • GET /jobs/{id}             │
└─────────────────────────────────────────────────────────────────┘
                │                                  │
                ▼                                  ▼
┌───────────────────────────────┐  ┌────────────────────────────────┐
│  cli.py (Click)                │  │  api.py (FastAPI)               │
│  Comandos: generar, subir-...  │  │  Endpoints REST + jobs en        │
│                                 │  │  segundo plano (engine/jobs.py) │
└───────────────────────────────┘  └────────────────────────────────┘
                              │                  │
                              └────────┬─────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────┐
│  engine/generador.py                                             │
│  Orquesta el pipeline completo (12 pasos)                        │
└─────────────────────────────────────────────────────────────────┘
       │              │             │            │             │
       ▼              ▼             ▼            ▼             ▼
  ┌────────┐    ┌──────────┐  ┌─────────┐  ┌─────────┐  ┌──────────┐
  │cliente │    │  docx    │  │  xlsx   │  │ claude  │  │  drive   │
  │        │    │  engine  │  │ engine  │  │ service │  │ service  │
  └────────┘    └──────────┘  └─────────┘  └─────────┘  └──────────┘
                                  │
                                  ▼
              ┌──────────────────────────────────┐
              │ Módulos especializados:           │
              │ • matriz_peligros (IPVR)          │
              │ • matriz_peligros_viales (PESV)   │
              │ • listado_maestro                 │
              │ • pesv_cleanup                    │
              │ • normativa, colores, auditoria   │
              └──────────────────────────────────┘
```

---

## Pipeline `generar` paso a paso

Cuando se ejecuta `python cli.py generar clientes_config/X.yaml`, ocurre lo siguiente en `engine/generador.py::generar_sgsst()`:

| Paso | Acción | Módulo |
|---|---|---|
| 1 | Cargar el YAML del cliente | `engine/cliente.py` |
| 2 | Generar contexto de marcadores `{{KEY}}` con variantes | `engine/cliente.py` |
| 3 | Crear carpeta de salida `clientes/<RAZON_SOCIAL>/` | shutil |
| 4 | Recorrer plantilla y procesar cada archivo: <br>• `.docx` → reemplazar marcadores + humanizar con Claude <br>• `.xlsx` → reemplazar marcadores + insertar logo | `docx_engine.py`, `xlsx_engine.py` |
| 5 | Aplicar correcciones legales al normograma, filtradas por el año del SG-SST (`anio_cliente`, tomado del campo `año` del YAML) | `normativa.py` |
| 5b | Adaptar COPASST ↔ Vigía según N° trabajadores | `generador.py` |
| 5c | Aplicar colores corporativos del cliente | `colores.py` |
| 6 | Si `generar_pesv: false` → eliminar carpeta PESV + limpiar hojas PESV de matrices | `pesv_cleanup.py` |
| 6b | Generar Matriz IPVR `SST-MTZ-012` | `matriz_peligros.py` |
| 6c | Generar Matriz Vial `PESV-MTZ-002` (si aplica) | `matriz_peligros_viales.py` |
| 7 | Generar reporte de auditoría | `auditoria.py` |
| 8 | Limpiar temporales | shutil |

Al terminar, en `clientes/<RAZON_SOCIAL>/` queda la estructura PHVA completa más el `REPORTE_AUDITORIA_SST.txt`.

### Correcciones legales (`engine/normativa.py`)

`CORRECCIONES` es una lista de reglas `(texto derogado, texto vigente, vigente_desde)`. Cada regla con `vigente_desde` solo se aplica si `anio_cliente >= vigente_desde`; las reglas sin `vigente_desde` (errores de cita, nunca fueron correctas) se aplican siempre. Si no se pasa `anio_cliente` (comportamiento previo a 2026-07-08), se aplican todas las reglas sin filtrar.

El comando `python cli.py normas` lista las reglas configuradas. El script `auditar_normativa.py` (independiente del pipeline de generación) escanea toda la plantilla maestra, extrae cada cita legal y produce `inventario_normas.json` para verificar manualmente contra fuentes oficiales cuáles siguen vigentes.

---

## Pipeline `subir-drive` paso a paso

Cuando se ejecuta `python cli.py subir-drive clientes_config/X.yaml`:

| Paso | Acción | Módulo |
|---|---|---|
| 1 | Cargar YAML, identificar carpeta de origen | `cliente.py` |
| 2 | Autenticarse en Drive del cliente (OAuth, abre navegador la primera vez) | `drive_service.py` |
| 3 | Crear carpeta raíz `<RAZON_SOCIAL>/` en Drive | `drive_service.py` |
| 4 | Replicar estructura de carpetas (PHVA + PESV) | `drive_service.py` |
| 5 | Subir cada archivo, retornar `{ruta_relativa: {link, nombre, id}}` | `drive_service.py` |
| 6 | Llamar a `actualizar_matriz_gestion_documental()` con los `drive_links` | `listado_maestro.py` |
| 7 | Inyectar `=HYPERLINK("url","texto")` en cada fila de la matriz `SST-MTZ-005` | `listado_maestro.py` |
| 8 | Generar pestañas por tipo (PROGRAMA, FORMATO, ...) | `listado_maestro.py` |
| 9 | Re-subir `SST-MTZ-005` actualizado a Drive | `drive_service.py` |

---

## API REST (`api.py`, Fase 4)

Segunda puerta de entrada a `engine/`, pensada para eventualmente correr en un servidor (de ahí el flujo OAuth "web" más abajo), aunque por ahora se corre y prueba localmente:

```cmd
uvicorn api:app --reload --port 8000
```

Todas las rutas salvo `/health` y `/auth/drive/callback` requieren `Authorization: Bearer <API_AUTH_TOKEN>`, verificado por `api_auth.py::verify_api_key` (dependency de FastAPI, no es login de usuario).

### Endpoints y a qué llaman

| Endpoint | Llama a |
|---|---|
| `GET /health` | — (chequeo de vida) |
| `GET /normas` | `engine/normativa.py::listar_correcciones()` (sin cambios) |
| `POST /clientes/validar` | `engine/validacion.py::validar_cliente_dict()` |
| `POST /clientes/verificar` | `engine/validacion.py::construir_reporte_verificacion()` |
| `POST /clientes/{id}/generar` | `engine/generador.py::generar_sgsst()`, en segundo plano vía `engine/jobs.py` |
| `POST /clientes/{id}/subir-drive` | `engine/drive_upload.py::ejecutar_subida_drive()`, en segundo plano |
| `GET /jobs/{job_id}`, `GET /clientes/{id}/jobs` | `engine/jobs.py` (lectura de `jobs.db`) |
| `GET /auth/drive/{id}/start`, `GET /auth/drive/callback` | `engine/drive_service.py::iniciar_auth_web()` / `completar_auth_web()` |

### Lógica compartida entre CLI y API

Para no duplicar comportamiento, se extrajo de `cli.py` a `engine/`:

- **`engine/validacion.py`** — `validar_cliente_dict()` (antes `cli.py::_validar_yaml`, lanzaba `click.BadParameter`, específico de Click) y `construir_reporte_verificacion()` (antes la lógica inline del comando `verificar`, imprimía una tabla Rich directamente). Ahora ambas retornan datos estructurados; `cli.py` los envuelve para su propia UI (tabla Rich / excepción Click) y `api.py` los serializa a JSON.
- **`engine/drive_upload.py`** — `ejecutar_subida_drive()` (antes `cli.py::_ejecutar_subida_drive`, ya estaba separada del comando Click). Se movió para que `api.py` no dependa de `cli.py`.

### Trabajos en segundo plano (`engine/jobs.py`)

`generar_sgsst()` tarda 2-40 min — nunca corre dentro del ciclo request/response. `POST /clientes/{id}/generar` valida sincrónicamente (falla rápido con `400` si el YAML es inválido), crea un registro de job y lo agenda con `BackgroundTasks` de FastAPI; responde `202` de inmediato con un `job_id`.

Estado persistido en **SQLite** (`jobs.db`, tabla `jobs`) — mismo patrón que ya usa `diskcache` (también SQLite) para el caché de Claude. Estados: `queued` → `running` → `done`/`failed`. Al arrancar la API (`@app.on_event("startup")`), cualquier job que quedó en `queued`/`running` de una corrida anterior se marca `failed` con `"Interrumpido por reinicio del servidor"` — el trabajo en sí no se puede reanudar, pero el estado no se pierde silenciosamente.

### Login de Drive por navegador (flujo OAuth "web")

`autenticar_drive()` (usada por el CLI) abre un navegador **en la misma máquina donde corre el proceso** (`InstalledAppFlow.run_local_server()`) — funciona en una PC, es imposible en un servidor remoto. Para la API se agregó un segundo flujo en `engine/drive_service.py`, sin tocar `autenticar_drive()`:

- `iniciar_auth_web(cliente_id, redirect_uri)` — genera la URL de consentimiento de Google (`google_auth_oauthlib.flow.Flow`, no `InstalledAppFlow`).
- `completar_auth_web(cliente_id, code, redirect_uri)` — intercambia el código por un token y lo guarda en **la misma ruta/formato** que ya usa `autenticar_drive()` (`tokens/token_<id>.json`) — por eso `autenticar_drive()` no necesitó ningún cambio: no le importa si el token se creó por el flujo de escritorio o el web.

**Prerrequisito manual** (no se puede hacer por código): crear un Client ID de tipo **"Web application"** en Google Cloud Console (el `credentials.json` actual es tipo "Desktop", se deja intacto), con el redirect URI `<API_BASE_URL>/auth/drive/callback` registrado, y descargarlo como `credentials_web.json` en la raíz. Sin ese archivo, `/auth/drive/{id}/start` falla con `FileNotFoundError` (mensaje explícito, no crashea el servidor).

`/auth/drive/callback` es la única ruta protegida distinta de `/health`: no requiere `API_AUTH_TOKEN` porque la llama Google directamente (su seguridad es el `code` de un solo uso que emite Google, práctica estándar de OAuth).

**Detalle importante — PKCE**: ambas funciones pasan `autogenerate_code_verifier=False` al crear el `Flow`. Por defecto `google-auth-oauthlib` activa PKCE, generando un `code_verifier` que solo vive en el objeto `Flow` de esa petición. Como `iniciar_auth_web()` y `completar_auth_web()` corren en dos peticiones HTTP separadas (dos objetos `Flow` distintos), ese `code_verifier` nunca le llega al segundo paso y Google rechazaba el intercambio con `invalid_grant: Missing code verifier`. No hace falta PKCE aquí porque `credentials_web.json` es un cliente confidencial (tiene `client_secret`) — PKCE existe para proteger clientes públicos (SPA/apps móviles) que no pueden guardar un secreto.

Verificado en vivo (2026-07-09): login completo por navegador con un `credentials_web.json` real, token guardado en `tokens/token_<id>.json` con `refresh_token` incluido.

---

## Módulos principales

### `engine/cliente.py`

Carga el YAML y expande el contexto con todas las variantes que pueden aparecer en los marcadores `{{...}}`:

```python
contexto = generar_contexto(cliente)
# → {
#   "RAZON_SOCIAL": "FUNDACIÓN CRISTOVISIÓN",
#   "razon_social": "FUNDACIÓN CRISTOVISIÓN",
#   "RAZÓN_SOCIAL": "FUNDACIÓN CRISTOVISIÓN",
#   ...
# }
```

**Lógica clave**: para cada campo del YAML genera variantes en mayúsculas, minúsculas, sin tildes, con/sin guiones bajos, para tolerar las inconsistencias de la plantilla original.

### `engine/docx_engine.py`

Procesa archivos Word. Hace dos cosas:

1. **Reemplazo de marcadores** usando `docxtpl` (Jinja2 sobre Word).
2. **Aplicación de correcciones de formato** sobre el XML: forzar color negro en textos rojos, conservar negrillas originales, preservar tablas azules.

Las funciones clave:
- `procesar_docx(plantilla, output, contexto)` — reemplazo principal
- `extraer_secciones_contenido(docx_path)` — saca secciones para humanizar con Claude
- `_aplicar_humanizacion_post()` — re-inserta textos humanizados respetando bold, listas, definiciones

### `engine/xlsx_engine.py`

Procesa archivos Excel. Reemplaza marcadores celda por celda + inserta logos/imágenes.

### `engine/claude_service.py`

Cliente de Anthropic API. Lo más importante:

- `humanizar_secciones(secciones, cliente, nivel)` — humaniza en batch
- `actualizar_normograma(texto, cliente)` — actualiza normograma eliminando normas derogadas
- `_construir_contexto(cliente)` — genera el contexto que se le pasa a Claude (COPASST vs Vigía, número de trabajadores, etc.)
- `_corregir_tildes(texto)` — post-procesa la salida para arreglar tildes (~80 palabras SST)
- Cache con `diskcache` en `.cache/claude_responses/` (TTL 30 días)

### `engine/drive_service.py`

Cliente de Google Drive API. Sube archivos, crea carpetas, devuelve links.

- Usa OAuth 2.0, token por cliente en `tokens/token_<NIT>.json`
- Maneja rate limit, reintentos
- Cache en memoria de IDs de carpetas para no re-crearlas

### `engine/listado_maestro.py`

**Módulo crítico**. Llena `SST-MTZ-005` con los links de Drive.

Funciones clave:

| Función | Para qué |
|---|---|
| `_normalizar_codigo()` | Aplica equivalencias (FTM → FMT) |
| `_normalizar_nombre()` | Lowercase, sin tildes, sin extensiones |
| `_similitud()` | Score Jaccard + containment + tolerancia a typos |
| `_similitud_palabras()` | Compara 2 palabras tolerando typos (SequenceMatcher) |
| `_indexar_drive_links_por_codigo()` | Convierte el dict de Drive en `{codigo: [candidatos]}` y `{prefijo: [candidatos]}` |
| `_elegir_mejor_match()` | Dada una fila de la matriz, elige el archivo correcto en Drive |
| `actualizar_matriz_gestion_documental()` | Pipeline principal: itera filas, busca match, inyecta `=HYPERLINK()` |
| `_generar_pestanas_por_tipo()` | Crea pestañas PROGRAMA, FORMATO, MATRIZ, etc. |

Algoritmo de matching:

```
Para cada fila con código X y nombre N en la sección S:
  1. Buscar candidatos con código X (después de normalizar FTM→FMT)
  2. Si la sección tiene numeral (ej. 4.1.2), filtrar candidatos por carpeta
  3. Calcular similitud(N, candidato) para cada uno
  4. Si max(similitud) >= 0.40 → asignar mejor candidato
  5. Si no hay match exacto → fallback por prefijo+nombre (SST-PRG-*)
  6. Si fallback >= 0.75 → asignar
  7. Si nada supera el umbral → dejar celda vacía
```

### `engine/matriz_peligros.py` (IPVR — SST-MTZ-012)

Llena la matriz de peligros laborales según los cargos del cliente.

- `PELIGROS_BASE_ADMINISTRATIVO` — 11 peligros que aplican a todos los cargos
- `EXTRAS_POR_TIPO` — peligros adicionales por tipo (operativo, conductor, SG)
- `_detectar_tipo_cargo()` — heurística por nombre
- `_extraer_denominacion()` — lee el campo DENOMINACIÓN de los manuales .docx
- `cargar_cargos_desde_manuales(dir)` — recorre la carpeta de manuales
- `_generar_datos_cargo()` — opcional: llama a Claude para personalizar actividades/tareas
- `generar_matriz_peligros()` — pipeline principal

### `engine/matriz_peligros_viales.py` (PESV — PESV-MTZ-002)

Llena la matriz vial según los cargos con desplazamiento.

- `CATALOGO` — dict `{factor: {actor: [peligros]}}` con los 5 factores GTC 45
- `_inferir_actores_viales()` — detecta conductor/motociclista/peatón/pasajero por cargo
- `_calcular_nivel_riesgo()` — pre-calcula el color del semáforo
- `generar_matriz_peligros_viales()` — pipeline principal

### `engine/pesv_cleanup.py`

Cuando `generar_pesv: false`, elimina hojas PESV de TODAS las matrices Excel del SG-SST generado.

- `_es_hoja_pesv()` — detecta nombres de hoja con normalización (case-insensitive, sin tildes)
- `limpiar_hojas_pesv_en_matrices()` — recorre `output_base/**/*.xlsx`

---

## Convenciones del código

### Naming

- Funciones públicas: `snake_case`, sin prefijo
- Funciones internas/privadas: `_snake_case` con guion bajo inicial
- Constantes: `UPPER_SNAKE_CASE`
- Variables: `snake_case` en español si refleja un concepto SST específico (ej: `nivel_riesgo`, `cliente`, `humanizar`)

### Imports

Orden Python estándar:
1. Stdlib (`os`, `re`, `pathlib`)
2. Third-party (`openpyxl`, `docx`, `anthropic`)
3. Local (`from engine.X import Y` o `from .X import Y`)

### Logging

Usa `logging.getLogger(__name__)` por módulo. `engine/generador.py` y `cli.py` configuran el handler raíz con Rich.

### Errores

- **Faltantes recuperables** (campo opcional vacío) → warning + continuar
- **Errores de plantilla** (archivo corrupto) → loggear + continuar con el siguiente
- **Errores de configuración** (YAML inválido) → fallar inmediatamente con mensaje claro
- **Fallos de API externa** (Claude, Drive) → reintentar 3 veces con backoff, luego loggear + continuar

---

## Cómo agregar...

### Un nuevo campo al YAML del cliente

1. Añadir el campo al ejemplo en `clientes_config/ejemplo_cliente.yaml`
2. En `engine/cliente.py::generar_contexto()`, agregarlo al dict de retorno con sus variantes
3. Si afecta a Claude, agregarlo a `engine/claude_service.py::_construir_contexto()`
4. Documentar en `docs/GUIA_CONSULTOR.md`

### Un nuevo peligro al catálogo IPVR

Edita `engine/matriz_peligros.py`:

```python
PELIGROS_BASE_ADMINISTRATIVO.append({
    "peligro": "Texto del peligro en mayúsculas",
    "clasificacion": "FÍSICO",  # u otra de las 7 categorías GTC 45
    "efectos": "Efectos posibles",
    "fuente": "Controles en fuente",
    "medio": "Controles en medio",
    "individuo": "Controles en individuo",
    "nd": 2,    # 2/6/10/14
    "ne": 3,    # 1/2/3/4
    "nc": 25,   # 10/25/60/100
    "peor": "Peor consecuencia",
    "req_legal": "SI",
    "eliminacion": "N/A",
    "sustitucion": "N/A",
    "ingenieria": "N/A",
    "administrativos": "Controles administrativos",
    "epp": "EPP requerido",
})
```

Y agregar un test correspondiente en `tests/test_matriz_peligros.py`.

### Un nuevo factor o peligro a la matriz vial

Editar el dict `CATALOGO` en `engine/matriz_peligros_viales.py`. La estructura es:

```python
CATALOGO = {
    "FACTOR HUMANO": {
        "conductor": [peligro1, peligro2, ...],
        "motociclista": [...],
        ...
    },
    ...
}
```

### Un nuevo tipo de documento al listado maestro

Agregar el tipo a `TIPOS_PESTANA` en `engine/listado_maestro.py`:

```python
TIPOS_PESTANA = [
    "PROGRAMA", "PLAN", ..., "NUEVO_TIPO",
]
```

Las pestañas se generan automáticamente si hay filas con ese tipo en la columna E del listado maestro PHVA.

### Una nueva equivalencia de prefijos (typo histórico)

En `engine/listado_maestro.py`:

```python
_PREFIJOS_EQUIVALENTES = {
    "FMT": "FMT",
    "FTM": "FMT",
    "NUEVO": "OTRO",  # ← agregar
}
```

Esto hace que cualquier código `SST-NUEVO-NNN` sea tratado como `SST-OTRO-NNN`.

---

## Cache

Hay dos sistemas de caché:

### Caché de Claude (`engine/claude_service.py`)

- Backend: `diskcache` en `.cache/claude_responses/cache.db`
- TTL: 30 días (configurable con `CACHE_EXPIRY_DAYS`)
- Key: hash del texto + contexto + nivel de humanización
- Limpiar: `python limpiar_cache.py`

### Caché de carpetas Drive (`engine/drive_service.py`)

- Backend: dict en memoria (`cache_carpetas`)
- Vive solo durante la ejecución (no persiste)
- Evita re-crear carpetas ya creadas en la misma sesión

> **No hay caché persistente de links de Drive**. Cada `subir-drive` re-pregunta a Drive qué archivos hay y reconstruye los links.

---

## Tests

Ver `tests/`. Se ejecutan con `python -m pytest`.

- `conftest.py` — fixtures compartidas (clientes mock, copias temporales de plantillas, stub de Claude)
- `test_listado_maestro.py` — matcher de links (30 tests)
- `test_matriz_peligros.py` — IPVR (17 tests)
- `test_matriz_peligros_viales.py` — vial + cleanup (33 tests)

**Política**: cada bug encontrado en producción debe convertirse en un regression test.

---

## Variables de entorno

| Variable | Default | Propósito |
|---|---|---|
| `ANTHROPIC_API_KEY` | _(vacío)_ | API key de Claude |
| `CLAUDE_MODEL` | `claude-haiku-4-5-20251001` | Modelo a usar |
| `CLAUDE_PRECIO_INPUT` | `1` | USD por millón de tokens input |
| `CLAUDE_PRECIO_OUTPUT` | `5` | USD por millón de tokens output |
| `PLANTILLA_PATH` | `Plantilla SG-SST 2026` | Carpeta de plantilla maestra |
| `OUTPUT_BASE_PATH` | `clientes` | Carpeta donde se generan los SG-SST |
| `CACHE_DIR` | `.cache` | Carpeta del caché de Claude |
| `CACHE_EXPIRY_DAYS` | `30` | TTL del caché |
| `API_AUTH_TOKEN` | _(vacío)_ | Token que deben enviar los clientes de la API REST (`Authorization: Bearer <token>`) |
| `API_BASE_URL` | `http://localhost:8000` | URL base de la API, usada para armar el redirect URI de OAuth |
| `JOBS_DB_PATH` | `jobs.db` | Ruta al SQLite de estado de trabajos en segundo plano |

---

## Archivos sensibles (no commitear)

```
credentials.json           # OAuth Google Cloud (flujo de escritorio/CLI)
credentials_web.json       # OAuth Google Cloud (flujo web/API, Fase 4)
jobs.db                    # Estado de trabajos en segundo plano de la API
.env                       # API keys y tokens
tokens/                    # Tokens de Drive de cada cliente
clientes/                  # SG-SST generados (datos sensibles de clientes)
.cache/                    # Respuestas cacheadas de Claude
```

Sugerencia para `.gitignore`:

```
credentials.json
credentials_web.json
jobs.db
.env
tokens/
clientes/
.cache/
__pycache__/
.pytest_cache/
ejemplo_matriz_*/
_demo_*.py
.test_*/
```

---

## Versión Python y dependencias

- **Python 3.12+** (probado en 3.12.10)
- Ver `requirements.txt` para versiones exactas
- Principales: `python-docx`, `docxtpl`, `openpyxl`, `anthropic`, `click`, `rich`, `pyyaml`, `diskcache`, `google-api-python-client`, `fastapi`, `uvicorn` (Fase 4)

---

## Limitaciones conocidas

1. **OneDrive**: trabajar en una carpeta sincronizada con OneDrive puede causar conflictos cuando OneDrive intenta sincronizar archivos mientras Python los escribe. Solución: cerrar OneDrive temporalmente o trabajar en una carpeta local.

2. **Archivos abiertos en Excel/Word**: Windows bloquea el archivo. El generador falla con `PermissionError`. Solución: cerrar todos los Word/Excel antes de generar.

3. **Manuales con formato muy distinto al esperado**: si el manual no tiene tabla con DENOMINACIÓN/DEPENDENCIA, el sistema usa el nombre del archivo como denominación. Funciona pero pierde precisión en la inferencia de actor vial.

4. **Tildes en nombres de archivo Windows**: archivos como `INSTALACIÓN.docx` pueden dar problemas si la consola usa codificación cp1252. El sistema usa rutas largas (`\\?\C:\...`) en `engine/utils.py::get_long_path()` para mitigar eso. Aparte, hasta el 2026-07-09 la salida por consola (`rich.Console`) podía **crashear** con `UnicodeEncodeError` al imprimir un carácter fuera de cp1252 (ej. "→") cuando el proceso no corre en una terminal interactiva — no era solo cosmético, tumbaba la generación completa. Se corrigió forzando UTF-8 en `sys.stdout`/`sys.stderr` al inicio de `engine/config.py` (`reconfigure(encoding="utf-8")`); el `os.environ["PYTHONIOENCODING"]` que tenía `cli.py` no servía porque esa variable solo afecta subprocesos nuevos, no el intérprete ya arrancado. Esto era invisible usando el CLI en una terminal normal, pero bloqueaba por completo a la API REST (Fase 4), que siempre corre así.

---

## Roadmap (siguientes mejoras sugeridas)

- [ ] Tests adicionales: `docx_engine`, `xlsx_engine`, `cliente`, `drive_service` (mock)
- [ ] Modo `--dry-run` para `subir-drive` (simular sin tocar Drive)
- [ ] Migrar la subida a Drive a async para acelerar
- [x] API REST (Fase 4) — `api.py`, ver sección dedicada arriba
- [ ] Frontend web (Fase 5)
- [ ] OCR de RUT y Cámara de Comercio (Fase 6)
