# Arquitectura técnica

Documento para desarrolladores que necesiten **mantener, depurar o extender** el sistema.

> Si eres consultor SST sin formación técnica, esto NO es para ti. Lee `GUIA_CONSULTOR.md`.

---

## Diagrama general

```
┌───────────────────────────┐  ┌────────────────────────┐  ┌──────────────────────────┐
│ USUARIO (Consultor SST)   │  │ CLIENTE HTTP (Fase 4)   │  │ CLIENTE FINAL (Fase 5)    │
│ • Edita clientes_config/  │  │ • POST /clientes/{id}/  │  │ • Navega /portal/         │
│   X.yaml                  │  │   generar                │  │ • Sube RUT/Cámara/Cédula  │
│ • Ejecuta python cli.py.. │  │ • GET /jobs/{id}         │  │   y confirma cada módulo  │
└───────────────────────────┘  └────────────────────────┘  └──────────────────────────┘
                │                          │                            │
                ▼                          ▼                            ▼
┌───────────────────────────┐  ┌──────────────────────────────────────────────────────┐
│  cli.py (Click)            │  │  api.py (FastAPI)                                     │
│  Comandos: generar, subir- │  │  Endpoints REST + jobs en segundo plano (engine/jobs.py)│
│  ...                        │  │  + wizard_router (wizard_api.py, Fase 5) montado en /wizard│
│                             │  │  + StaticFiles del portal montado en /portal          │
└───────────────────────────┘  └──────────────────────────────────────────────────────┘
                              │                  │                    │
                              │                  │                    ▼
                              │                  │      ┌───────────────────────────┐
                              │                  │      │ wizard_store (SQLite) +    │
                              │                  │      │ engine/extraccion.py       │
                              │                  │      │ (Claude, texto de PDF)     │
                              │                  │      └───────────────────────────┘
                              │                  │                    │
                              │                  │   POST /wizard/borradores/{id}/finalizar
                              │                  │   escribe un YAML nuevo en clientes_config/
                              │                  │                    │
                              └────────┬─────────┴────────────────────┘
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
              │ • pie_pagina, estilo_encabezado,  │
              │   tipo_letra (Word + Excel)       │
              └──────────────────────────────────┘
```

El asistente por módulos (Fase 5) **no genera nada por sí mismo** — solo produce un YAML de cliente. La generación real sigue siendo el mismo `POST /clientes/{id}/generar` de la Fase 4, que a su vez llama al mismo `engine/generador.py` de siempre. Ver la sección dedicada más abajo.

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
| 5d | Posición del número de página en el pie (opcional, `pie_pagina_numero`) | `pie_pagina.py` |
| 5e | Estilo del encabezado: logo/título/código/versión/fecha reacomodados (opcional, `estilo_encabezado`), aplica a `.docx` y `.xlsx` | `estilo_encabezado.py` |
| 5f | Tipo de letra de todo el documento (opcional, `tipo_letra`), aplica a `.docx` y `.xlsx` | `tipo_letra.py` |
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

- `iniciar_auth_web(cliente_id, redirect_uri)` — genera la URL de consentimiento de Google (`google_auth_oauthlib.flow.Flow`, no `InstalledAppFlow`); el `state` que le pasa a Google es `cliente_id` firmado con HMAC (`_firmar_state`), no el valor crudo.
- `completar_auth_web(state, code, redirect_uri)` — verifica la firma del `state` (`_verificar_state`, rechaza uno falsificado o con un `cliente_id` inválido), intercambia el código por un token y lo guarda en **la misma ruta/formato** que ya usa `autenticar_drive()` (`tokens/token_<id>.json`) — por eso `autenticar_drive()` no necesitó ningún cambio: no le importa si el token se creó por el flujo de escritorio o el web. Devuelve `(cliente_id, email)` — el `cliente_id` verificado y el correo de la cuenta de Google que se acaba de autorizar (vía `drive.about().get(fields="user")`, no requiere scope adicional), para que la página de confirmación se lo muestre a quien autorizó y pueda notar si conectó la cuenta equivocada. `iniciar_auth_web()` además pasa `prompt="select_account consent"` (no solo `"consent"`) para que Google siempre muestre el selector de cuenta, incluso si el navegador ya tenía una sesión de Google abierta — sin esto era fácil autorizar en silencio la cuenta de quien hacía clic en vez de la del cliente.

**Prerrequisito manual** (no se puede hacer por código): crear un Client ID de tipo **"Web application"** en Google Cloud Console (el `credentials.json` actual es tipo "Desktop", se deja intacto), con el redirect URI `<API_BASE_URL>/auth/drive/callback` registrado, y descargarlo como `credentials_web.json` en la raíz. Sin ese archivo, `/auth/drive/{id}/start` falla con `FileNotFoundError` (mensaje explícito, no crashea el servidor).

`/auth/drive/callback` es la única ruta protegida distinta de `/health`: no requiere `API_AUTH_TOKEN` porque la llama Google directamente. Al ser pública, no confía en nada de lo que le llega sin verificar: la seguridad viene de tres capas — el `code` de un solo uso que emite Google (práctica estándar de OAuth), la firma HMAC del `state` (ver arriba, cierra la puerta a que alguien invente un `state` propio), y la lista blanca de `_CLIENTE_ID_VALIDO` en `_token_path()` (evita que un `cliente_id` con `/` o `..` escriba el token fuera de `tokens/`). Los valores que la respuesta HTML le muestra al usuario (`cliente_id`, mensaje de error) pasan por `html.escape()` antes de interpolarse.

**Detalle importante — PKCE**: ambas funciones pasan `autogenerate_code_verifier=False` al crear el `Flow`. Por defecto `google-auth-oauthlib` activa PKCE, generando un `code_verifier` que solo vive en el objeto `Flow` de esa petición. Como `iniciar_auth_web()` y `completar_auth_web()` corren en dos peticiones HTTP separadas (dos objetos `Flow` distintos), ese `code_verifier` nunca le llega al segundo paso y Google rechazaba el intercambio con `invalid_grant: Missing code verifier`. No hace falta PKCE aquí porque `credentials_web.json` es un cliente confidencial (tiene `client_secret`) — PKCE existe para proteger clientes públicos (SPA/apps móviles) que no pueden guardar un secreto.

Verificado en vivo (2026-07-09): login completo por navegador con un `credentials_web.json` real, token guardado en `tokens/token_<id>.json` con `refresh_token` incluido.

---

## Asistente de onboarding por módulos (`wizard_api.py`, Fase 5)

Portal de autoservicio para que el cliente final llene sus propios datos, en vez de que un consultor escriba el YAML a mano. Router separado de `api.py` (montado con `app.include_router(wizard_router)`, prefijo `/wizard`) para no arriesgar regresión en las rutas de la Fase 4 al crecer el número de endpoints. Reusa la misma `Depends(verify_api_key)` de siempre — **no** es un sistema de login por cliente, sigue siendo el mismo token estático; login real por cliente queda fuera de alcance (tema aparte, más grande).

El frontend (`static/wizard/*.html`) se sirve como archivos planos con `StaticFiles(directory=..., html=True)` montado en `/portal` — **HTML/JS sin build toolchain**, deliberadamente sin React: el proyecto vive en una carpeta sincronizada con OneDrive, y `node_modules` es justo el tipo de conflicto que ya documenta la sección de Limitaciones conocidas más abajo.

### Modelo de datos: el "borrador" (`engine/wizard_store.py`)

Cada sesión del asistente es un **borrador**, persistido en `wizard_drafts.db` (SQLite puro, mismo patrón que `engine/jobs.py`: un `_conexion()` con context manager, columnas JSON serializadas). Tabla `wizard_drafts`:

| Columna | Contenido |
|---|---|
| `borrador_id` | UUID, clave primaria |
| `modulo_actual` | Puntero entero — el módulo que falta confirmar |
| `modulos_confirmados` | Lista JSON de módulos ya confirmados (`["modulo_1", "modulo_2", ...]`) |
| `datos_json` | Los datos confirmados de cada módulo, por clave (`{"modulo_1": {...}, ...}`) |
| `archivos_json` | Rutas de los archivos subidos (`{"rut": "wizard_uploads/<id>/rut.pdf", ...}`) |
| `extraccion_json` | Lo que Claude extrajo de cada documento del módulo 1, **sin combinar** (`{"rut": {...}, "camara_comercio": {...}}`) |
| `estado` | `en_progreso` / `finalizado` |
| `config_path_final` | Ruta del YAML una vez que se llama a `/finalizar` |

`MODULOS_ORDEN = ["modulo_1", ..., "modulo_7", "modulo_8"]` define el orden. `guardar_modulo(borrador_id, modulo_key, datos)` agrega `modulo_key` a `modulos_confirmados` y avanza `modulo_actual` — nunca lo retrocede (`max(actual, indice+2)`), así que confirmar módulos fuera de orden (si el gate lo permitiera) no corrompe el puntero.

### Bloqueo secuencial (el gate)

El bloqueo **no vive en `wizard_store.py`** — esa capa solo guarda datos, no juzga el orden. Vive en `wizard_api.py::_verificar_gate(borrador, modulo_key)`: antes de aceptar la confirmación del módulo N, revisa que el módulo N-1 esté en `modulos_confirmados`; si no, responde `409`. Cada endpoint `/modulo-N/confirmar` (y las variantes del módulo 7) llama a este gate antes de guardar nada.

### Módulo 1 — documentos y extracción (la pieza no trivial)

`POST /wizard/borradores/{id}/modulo-1/archivos` acepta cualquier subconjunto de 5 archivos (`rut`, `camara_comercio`, `cedula_rl` como PDF; `logo`, `firma_rl` como imagen). Por cada PDF subido:

1. `engine/pdf_utils.py::extraer_texto_pdf()` saca el texto (pdfplumber, con fallback a pypdf; nunca lanza excepción — si el PDF es un escaneo sin capa de texto, retorna `""`).
2. `engine/extraccion.py::extraer_datos_documento(tipo, texto)` le pide a Claude un JSON con los campos esperados para ese tipo de documento (6 campos para RUT, 9 para Cámara de Comercio, 2 para Cédula — ver `CAMPOS_POR_TIPO`). Mismo patrón que ya usaba `matriz_peligros.py::_generar_datos_cargo()`: prompt pidiendo JSON puro, se le quitan los fences de markdown con regex, `json.loads`, y **cualquier falla retorna `{}`** en vez de propagar la excepción (sin `ANTHROPIC_API_KEY`, sin saldo, texto vacío, o un JSON mal formado — todos caen al mismo camino silencioso, el usuario llena a mano lo que falte).
3. Cada documento se guarda **por separado** en `extraccion_json` (`guardar_extraccion`) — nunca se combinan ahí. Combinarlos temprano perdería la posibilidad de detectar que, por ejemplo, la dirección del RUT no coincide con la de la Cámara de Comercio.

`engine/extraccion.py::detectar_conflictos(extracciones)` compara los campos que aparecen en dos o más documentos (`razon_social`, `nit`, `direccion_empresa`, `ciudad_empresa`, `departamento_empresa`), normalizando espacios/mayúsculas, y devuelve la lista de discrepancias. El endpoint la incluye en la respuesta junto con la extracción cruda; `static/wizard/modulo1.html` la muestra al usuario para que elija un valor o escriba uno propio — **nunca se queda con un valor en silencio cuando hay dos fuentes que no coinciden**.

`POST /wizard/borradores/{id}/modulo-1/confirmar` recibe los valores ya revisados/corregidos por el usuario (no vuelve a leer `extraccion_json`) y los valida con `engine/validacion.py::validar_cliente_dict()` antes de guardarlos — la misma validación que ya usan CLI y API.

**Decisión de diseño (no reabrir)**: extracción con Claude leyendo **texto** de PDF, no imágenes ni un servicio de OCR aparte, y no Gemini. La mayoría de RUT/Cámara de Comercio descargados de la DIAN/Cámara traen texto real seleccionable, así que estructurar ese texto en campos es barato (input de texto, no de imagen). Se descartó la capa gratuita de Gemini por el riesgo de que datos personales (cédula, dirección) pasen por una política de datos que puede permitir entrenamiento.

### Módulos 2 a 6

Formularios simples, cada uno con su propio modelo Pydantic (`Modulo2Request`, `Modulo3Request` con `sedes: list[Sede]`, etc.) y su endpoint `/modulo-N/confirmar` que solo llama al gate y a `wizard_store.guardar_modulo()`. Dos endpoints de pre-llenado (`GET /defaults/modulo-4` con los datos fijos de MEDINA desde `engine/wizard_defaults.py`, `GET /defaults/modulo-5` con la fecha de hoy vía `engine/cliente.py::formatear_fecha_hoy()` — la misma función que usa `generar_contexto()`, para no tener dos cálculos de fecha que se puedan desincronizar) evitan que el usuario escriba de cero lo que ya se sabe. El módulo 4 además acepta `POST /modulo-4/archivos` para la firma del responsable SST.

### Módulo 7 — cargos (opcional)

A diferencia de los módulos 2-6, tiene **tres caminos**, todos terminando en `guardar_modulo(..., "modulo_7", {...})` con un campo `modo` que decide cómo se resuelve en `finalizar()`:

| Endpoint | `modo` guardado | Efecto en el YAML final |
|---|---|---|
| `POST /modulo-7/manuales` | `"manuales"` | Guarda los manuales subidos en `wizard_uploads/<id>/manuales/`; al finalizar se copian a `clientes_config/<slug>_assets/manuales_cargos/` y el YAML recibe `manuales_cargos_dir` — el pipeline existente (`matriz_peligros.py::cargar_cargos_desde_manuales()`) detecta el cargo del contenido del manual, sin que el usuario escriba nada |
| `POST /modulo-7/lista` | `"lista"` | Los cargos escritos a mano quedan directo como `cargos: [...]` en el YAML (rechaza con `400` si la lista viene vacía) |
| `POST /modulo-7/saltar` | `"omitido"` | No agrega ni `cargos` ni `manuales_cargos_dir` — `generar_sgsst()` ya soporta ausencia total de cargos, simplemente omite la Matriz IPVR |

Es el único módulo **excluido** de la lista `requeridos` en `finalizar()` — se puede finalizar el borrador sin haberlo tocado.

### Módulo 8 — estilo del documento

Formulario simple (`Modulo8Request`: `pie_pagina_numero`, `estilo_encabezado`, `tipo_letra`, los tres opcionales con default de plantilla), pero con una vista previa en vivo en `static/wizard/modulo8.html` que arma una tabla HTML del encabezado (`<table class="mini-header">`) usando el **logo real ya subido** (`GET /wizard/borradores/{id}/archivos/logo`, servido por `obtener_archivo_subido()`) y la **versión/fecha ya confirmadas** en el módulo 5 (`GET /borradores/{id}`), para que el usuario vea cómo quedará su propio encabezado antes de confirmar — no una imagen de ejemplo estática. Los tres valores se guardan tal cual (`guardar_modulo(..., "modulo_8", body.model_dump())`) y en `finalizar()` se vuelcan directo al YAML (`cliente.update(datos.get("modulo_8", {}))`); `engine/generador.py` los aplica en el pipeline (ver pasos 5d-5f más arriba) solo si vienen no vacíos.

### Finalizar (`POST /wizard/borradores/{id}/finalizar`)

1. Verifica que todos los módulos obligatorios (1-6 y 8; el 7 es opcional) estén confirmados (`409` si falta alguno).
2. `_ensamblar_cliente(borrador)` mezcla los `datos_json` de cada módulo en un solo dict (equivalente a lo que hoy escribe un consultor a mano en el YAML).
3. Valida el dict resultante con `validar_cliente_dict()` (misma validación de siempre).
4. Calcula el `slug` desde `razon_social` y **copia** (no referencia) logo/firma_rl/firma_sst desde `wizard_uploads/<id>/` hacia `clientes_config/<slug>_assets/`, con rutas relativas al propio YAML — exactamente lo que ya espera `obtener_imagenes()` (`engine/cliente.py`). Si el módulo 7 vino en modo `"manuales"`, copia también esa carpeta.
5. Escribe `clientes_config/<slug>.yaml` con `yaml.safe_dump()` y marca el borrador `finalizado` (`wizard_store.marcar_finalizado`).
6. Responde `{"config_path": ..., "cliente_id": slug, "siguiente_paso": "POST /clientes/{slug}/generar"}` — el handoff explícito al endpoint de la Fase 4, sin duplicar la lógica de generación.

**Por qué se copia y no se referencia**: `wizard_uploads/` es una carpeta de borrador desechable (se puede borrar con `DELETE /wizard/borradores/{id}`); `clientes_config/` es el destino permanente que ya usa el resto del sistema.

### Endpoints completos

| Método | Ruta | Para qué |
|---|---|---|
| `POST` | `/wizard/borradores` | Crea un borrador nuevo, `modulo_actual=1` |
| `GET` | `/wizard/borradores/{id}` | Consulta el estado completo del borrador |
| `DELETE` | `/wizard/borradores/{id}` | Borra el borrador y su carpeta de `wizard_uploads/` |
| `POST` | `/wizard/borradores/{id}/modulo-1/archivos` | Sube cualquier subconjunto de rut/camara_comercio/cedula_rl/logo/firma_rl; extrae los PDF |
| `GET` | `/wizard/borradores/{id}/modulo-1/extraccion` | Reconsulta la extracción ya guardada (recargar página sin perder progreso) |
| `POST` | `/wizard/borradores/{id}/modulo-1/confirmar` | Guarda los datos ya revisados por el usuario |
| `POST` | `/wizard/borradores/{id}/modulo-2/confirmar` | Riesgo, ARL, N° trabajadores, color |
| `POST` | `/wizard/borradores/{id}/modulo-3/confirmar` | Lista de sedes |
| `POST` | `/wizard/borradores/{id}/modulo-4/confirmar` | Responsable SST + consultora |
| `POST` | `/wizard/borradores/{id}/modulo-4/archivos` | Firma del responsable SST |
| `POST` | `/wizard/borradores/{id}/modulo-5/confirmar` | Fechas, versión, PESV, normograma |
| `POST` | `/wizard/borradores/{id}/modulo-6/confirmar` | Descripción de la empresa |
| `POST` | `/wizard/borradores/{id}/modulo-7/manuales` | Cargos vía manuales de funciones |
| `POST` | `/wizard/borradores/{id}/modulo-7/lista` | Cargos escritos a mano |
| `POST` | `/wizard/borradores/{id}/modulo-7/saltar` | Omite la Matriz IPVR |
| `POST` | `/wizard/borradores/{id}/modulo-8/confirmar` | Posición del número de página, estilo de encabezado, tipo de letra |
| `GET` | `/wizard/borradores/{id}/defaults/modulo-4` | Pre-llenado con los datos fijos de MEDINA |
| `GET` | `/wizard/borradores/{id}/defaults/modulo-5` | Pre-llenado con la fecha de hoy |
| `GET` | `/wizard/borradores/{id}/archivos/{rol}` | Sirve un archivo ya subido (ej. el logo, para la vista previa del módulo 8) |
| `POST` | `/wizard/borradores/{id}/finalizar` | Ensambla y escribe el YAML final |

Verificado con un flujo end-to-end real (2026-07-09/10): borrador → 7 módulos → `finalizar()` → YAML resultante cargado sin errores por `cargar_cliente()`/`generar_contexto()`/`obtener_imagenes()` → `POST /clientes/{id}/generar` disparado con ese `config_path` → 117 archivos generados, 0 errores.

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

### `engine/pie_pagina.py`

Cambia la posición del número de página en el pie de todos los `.docx` generados (`izquierda`/`centro`/`derecha`/`ninguno`), manipulando directamente el XML de `word/footer*.xml` dentro del zip (mismo patrón de "abrir zip, sustituir XML con regex, reescribir zip" que usa `colores.py`).

### `engine/estilo_encabezado.py`

Reacomoda logo, título, código, versión y fecha del encabezado en 5 grillas alternativas (`ESTILOS_VALIDOS = "1".."4"`, más el original vacío). Aplica tanto a `.docx` como a `.xlsx`, con estrategias distintas por lo distinto de cada formato:

- **Word**: el encabezado real vive en una tabla de 3 filas propia (confirmado en 147/149 documentos reales de la plantilla), con anchos de columna en cm — se reconstruye con la API de objetos de `python-docx` (`_construir_estilo_N`), moviendo el párrafo con el logo en vez de reconstruir la imagen.
- **Excel**: el bloque de encabezado NO está en una columna fija (varía entre B, C, E, F, J... según la matriz) y comparte columnas con la tabla de datos de abajo. Se ubica por **búsqueda de texto** — celdas con la forma "CÓDIGO: ..." (`_buscar_celdas_codigo_xlsx`, exige los dos puntos para no confundirlo con una columna de datos que casualmente empiece con "código", ej. "CÓDIGO CIE 10") — en vez de coordenadas fijas; confirmado sobre 53 archivos reales de un cliente, 46 (87%) tienen al menos un bloque detectable así. La búsqueda cubre **toda la hoja**, no solo las primeras filas: algunas matrices repiten el bloque completo varias veces (una por sección/categoría, ej. `SST-MTZ-002` lo repite 4 veces; o un formulario que se repite más abajo para una segunda entrada) y cada repetición se restyla por separado, con su propio logo (`_asociar_imagenes_a_bloques` empareja cada bloque con la imagen más cercana a su fila superior). Los estilos se reconstruyen como fracciones del **ancho real de columna** (`_asignar_columnas`, suma de `ws.column_dimensions[...].width`, no cantidad de columnas) para no apretar textos largos en columnas angostas, y el logo se **escala y centra** dentro de su bloque (`_mover_logo_xlsx`, usando el ancho/alto real del bloque en EMU) en vez de solo reposicionarse al tamaño original — si no, quedaba desproporcionado y pegado a una esquina. Solo se procesa la primera hoja de cada libro (`wb.worksheets[0]`) y ningún bloque sale nunca de las filas que ya ocupaba, para no invadir la fila de encabezados de la tabla de datos ni las filas de otro bloque. Las matrices sin ningún "CÓDIGO: ..." detectable (consolidados/indicadores) quedan sin cambios, sin error.

`aplicar_estilo_encabezado(output_base, estilo)` recorre `*.docx` y `*.xlsx` de la carpeta de salida y agrega los contadores de ambos en un solo `{"archivos": n, "omitidos": m}`.

### `engine/tipo_letra.py`

Cambia la fuente de todo el documento (títulos, subtítulos y contenido a la vez — la plantilla no usa Estilos de Word por nivel, solo `w:docDefaults` en `styles.xml`, así que un único swap de fuente cubre todo el `.docx`). En Word sustituye `w:rFonts` en `document.xml`/`styles.xml`/headers/footers; en Excel sustituye `<name val="...">` en `xl/styles.xml` y los tipos de letra del tema en `xl/theme/*.xml`. `FUENTES_VALIDAS = ("Arial", "Calibri", "Times New Roman")`.

**Decisión de alcance (no reabrir)**: no se implementó diferenciación de **tamaño** de letra por título/subtítulo/contenido. Se investigó una heurística de "título principal" (mayor tamaño de fuente cerca del inicio del documento) sobre una muestra real de 8 documentos y resultó **no confiable** — confundía líneas de firma, oraciones del cuerpo y encabezados de sección duplicados en 3-4 de los 8 casos. Se decidió no construir sobre una heurística frágil.

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
- `test_listado_maestro.py` — matcher de links (40 tests)
- `test_matriz_peligros.py` — IPVR (17 tests)
- `test_matriz_peligros_viales.py` — vial + cleanup (23 tests)
- `test_matriz_documental.py` — matriz documental (4 tests)
- `test_normativa.py` — correcciones legales filtradas por año (13 tests)
- `test_cliente.py` — carga de YAML y contexto de marcadores (4 tests)
- `test_generador.py` — orquestador del pipeline (2 tests)
- `test_extraccion.py` — extracción de datos de RUT/Cámara/Cédula con Claude mockeado (9 tests) + 1 `integration` real
- `test_wizard_store.py` — CRUD del borrador sin HTTP (8 tests)
- `test_wizard_api.py` — endpoints del asistente con `TestClient`, incluye bloqueo secuencial (`409`) y flujo completo hasta `finalizar()` (14 tests)
- `test_pie_pagina.py` — posición del número de página en el pie (5 tests)
- `test_estilo_encabezado.py` — reacomodo del encabezado en Word y Excel, incluida la búsqueda por texto de "CÓDIGO" en distintas columnas (17 tests)
- `test_tipo_letra.py` — cambio de fuente en Word y Excel (7 tests)

170 tests en total: 165 corren por defecto, 4 están marcados `integration` (requieren `ANTHROPIC_API_KEY` con saldo real) y 1 marcado `slow`.

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
| `WIZARD_DB_PATH` | `wizard_drafts.db` | Ruta al SQLite de borradores del asistente por módulos (Fase 5) |
| `WIZARD_UPLOADS_DIR` | `wizard_uploads` | Carpeta donde quedan temporalmente los archivos subidos en el asistente |

---

## Archivos sensibles (no commitear)

```
credentials.json           # OAuth Google Cloud (flujo de escritorio/CLI)
credentials_web.json       # OAuth Google Cloud (flujo web/API, Fase 4)
jobs.db                    # Estado de trabajos en segundo plano de la API
wizard_drafts.db           # Estado de los borradores del asistente por módulos (Fase 5)
wizard_uploads/            # Archivos (RUT, Cámara, Cédula, logo, firmas) subidos en el asistente
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
wizard_drafts.db
wizard_uploads/
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
- Principales: `python-docx`, `docxtpl`, `openpyxl`, `anthropic`, `click`, `rich`, `pyyaml`, `diskcache`, `google-api-python-client`, `fastapi`, `uvicorn` (Fase 4), `python-multipart` (subida de archivos en la API/asistente, Fase 4-5), `pdfplumber`/`pypdf` (extracción de texto de PDF, Fase 5)

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
- [x] Portal/asistente por módulos (Fase 5) — `wizard_api.py` + `static/wizard/`, ver sección dedicada arriba
- [ ] Login real por cliente (hoy el portal usa el mismo API key estático de la Fase 4)
- [ ] Cobro/facturación de clientes cuando el software corra en un servidor
- [ ] OCR de respaldo para PDFs escaneados sin capa de texto (Fase 6)
