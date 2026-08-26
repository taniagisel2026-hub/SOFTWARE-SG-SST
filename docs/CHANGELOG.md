# Changelog

Registro de cambios del SG-SST Generator. Formato basado en [Keep a Changelog](https://keepachangelog.com/).

## [No publicado]

### Agregado (auditoría técnica y endurecimiento, 2026-08-26)
- **Validación real de archivos subidos en el portal** (`wizard_api.py`): antes solo se confiaba en la extensión del nombre del archivo. Ahora se valida la firma binaria real (PNG/JPG/PDF) y un límite de 15 MB, y el archivo se sirve de vuelta con un `Content-Type` explícito por extensión — cierra un vector de XSS almacenado vía la vista previa del Módulo 8.
- **Callback de OAuth de Drive blindado** (`engine/drive_service.py`, `api.py`): `/auth/drive/callback` es la única ruta pública de la API (Google la llama sin API key) y antes confiaba en el `state` tal cual. Ahora `iniciar_auth_web()` firma el `state` con HMAC y `completar_auth_web()` verifica esa firma; `cliente_id` se valida contra una lista blanca antes de usarse para nombrar cualquier archivo (evita path traversal); la respuesta HTML escapa cualquier valor que muestre (evita XSS reflejado).
- **Reintentos ante fallos de red**: `claude_service._llamar_claude()` ahora reintenta también ante `APIConnectionError`/`InternalServerError`, no solo `RateLimitError`. Las 7 llamadas a la API de Drive usan `execute(num_retries=3)`, el mecanismo propio de `googleapiclient`.
- **Tests nuevos**: `tests/test_docx_engine.py` (31), ampliación de `tests/test_claude_service.py` (+10), `tests/test_drive_service.py` (17, incluye la firma del state y la validación de cliente_id), `tests/test_api_drive_callback.py` (2). Total: 122 → 250 tests.
- **Limpieza de normas de COVID-19 obsoletas** en 3 plantillas Word (`PESV-PLN-003`, `SST-FMT-073`, `SST-PLN-004`) — la emergencia sanitaria terminó el 30-jun-2022, estas eran medidas exclusivas de esa coyuntura. Quedan pendientes 7 documentos más con la Circular 015/2024 y la Resolución 132/2004 (ver `auditoria_normativa_2026-07-08.html`), y 5 normas más que solo viven en `SST-MTZ-006 MATRIZ DE REQUISITOS LEGALES.xlsx` (el motor de corrección no procesa `.xlsx` todavía).
- **Contraste de los `<select>` del portal**: el texto heredaba el color claro del tema oscuro, pero el navegador dibuja la lista desplegable con su propio fondo (casi siempre blanco) — quedaba casi ilegible. Fix en `static/wizard/wizard.css`.
- **Fixture de tests desactualizada**: `tests/conftest.py::output_listado_maestro` buscaba un `SST-MTZ-005` sin etiqueta que ya no existe en la plantilla (ahora ships `[CAT-I]`/`[CAT-II]`/`[CAT-III]`, mismo mecanismo que `[COPASST]`/`[VIGÍA]` en `generador.py`) — 4 tests quedaban saltados en silencio. Corregida para tomar cualquiera de las 3 variantes.

### Agregado
- Documentación inicial: `README.md`, `docs/GUIA_CONSULTOR.md`, `docs/ARQUITECTURA.md`, `docs/CHANGELOG.md`
- Suite de tests automatizados (122 tests) en `tests/`
- Archivo `.env.example` para facilitar setup inicial
- Configuración de pytest (`pytest.ini`)
- **Auditoría de normativa** (`auditar_normativa.py`): extrae todas las citas de leyes/decretos/resoluciones de la plantilla maestra y las cruza contra fuentes oficiales (Suin-Juriscol, Función Pública, Mintrabajo, Minsalud). Genera `inventario_normas.json` con las 150 normas encontradas. 32 correcciones verificadas agregadas a `CORRECCIONES` en `engine/normativa.py`; 13 normas quedaron marcadas para revisión legal manual (ver `auditoria_normativa_2026-07-08.html`).
- **Correcciones legales filtradas por año del SG-SST**: `aplicar_correcciones_docx()` acepta `anio_cliente` (tomado del campo `año` del YAML), de modo que un SG-SST histórico solo recibe las correcciones vigentes para su año y no normas que aún no existían.
- **Comando `python cli.py normas`**: lista las correcciones legales configuradas en el sistema.
- **API REST (Fase 4)** (`api.py`, `api_auth.py`, `engine/jobs.py`): segunda puerta de entrada a `engine/` por HTTP, sin reemplazar el CLI. Endpoints para generar, subir a Drive, consultar estado de trabajos (persistido en SQLite, sobrevive reinicios) y listar normas. Autenticación mínima por API key (`API_AUTH_TOKEN`). Se extrajo `engine/validacion.py` y `engine/drive_upload.py` de `cli.py` para que ambas interfaces compartan la misma lógica sin duplicarla.
- **Flujo OAuth "web" para Google Drive** (`engine/drive_service.py::iniciar_auth_web`/`completar_auth_web`): permite autorizar el acceso a Drive de un cliente desde un navegador remoto (no en la máquina donde corre el servidor), necesario para que la API funcione fuera de una PC local. No modifica el flujo de escritorio existente (`autenticar_drive`), que sigue usando el CLI sin cambios.
- **Asistente de onboarding por módulos / portal del cliente (Fase 5)** (`wizard_api.py`, `engine/wizard_store.py`, `engine/extraccion.py`, `engine/pdf_utils.py`, `engine/wizard_defaults.py`, `static/wizard/`): el cliente final llena sus propios datos desde `http://localhost:8000/portal/`, en 7 módulos secuenciales bloqueados uno tras otro (`409` si se intenta saltar uno). El módulo 1 sube RUT/Cámara de Comercio/Cédula (PDF) + logo/firma, extrae los datos con Claude (texto de PDF vía `pdfplumber`, no OCR ni Gemini) y detecta conflictos entre documentos (ej. dirección distinta en RUT vs Cámara de Comercio) para que el usuario confirme en vez de que uno pise al otro en silencio. El módulo 7 (cargos para la Matriz IPVR) admite subir manuales de funciones, escribir la lista a mano, u omitirlo. Al finalizar, ensambla un YAML nuevo en `clientes_config/` y entrega el mismo `POST /clientes/{id}/generar` de la Fase 4 — no duplica la lógica de generación. Estado en `wizard_drafts.db` (SQLite), archivos temporales en `wizard_uploads/`.

### Corregido
- **Crash de codificación UTF-8 en salida sin terminal interactiva**: `rich.Console` truena con `UnicodeEncodeError` al imprimir caracteres fuera de cp1252 (ej. "→") cuando el proceso no corre en una terminal real — bloqueaba por completo la generación vía API REST. El `os.environ["PYTHONIOENCODING"]` que tenía `cli.py` no tenía ningún efecto real (esa variable no reconfigura un intérprete ya arrancado). Corregido con `sys.stdout.reconfigure(encoding="utf-8")` en `engine/config.py`, compartido por CLI y API.

---

## [8.5] — 2026-05

### Agregado
- **Pestañas por tipo en el listado maestro**: SST-MTZ-005 ahora genera hojas adicionales (PROGRAMA, PLAN, PROCEDIMIENTO, INSTRUCTIVO, REGLAMENTO, MANUAL, POLÍTICA, INDICADOR, MATRIZ, FORMATO) con todos los documentos de cada tipo y sus links.
- **Matriz de Peligros Viales** (`PESV-MTZ-002`): generación automática para clientes con PESV, detectando actores viales (conductor/motociclista/peatón/pasajero) por cargo.
- **Color durazno en columna CLASIFICACIÓN** de la matriz vial.
- **Negrillas selectivas** en matriz vial: PROCESO, ACTORES VIALES, CLASIFICACIÓN e INTERPRETACIÓN DEL NIVEL DE RIESGO.
- **Colores semáforo en columna R** de la matriz vial (rojo/naranja/verde según probabilidad × exposición).
- **Limpieza de hojas PESV** en todas las matrices Excel cuando el cliente no solicita PESV (`engine/pesv_cleanup.py`).
- **Verificador de links del listado maestro** (`verificar_links_listado.py`) con estados OK/WARN/FAIL/VACIO.
- **Script de limpieza de caché** (`limpiar_cache.py`).

### Corregido
- **Equivalencia FMT ↔ FTM** en matcher de links. Antes archivos con código FTM no matcheaban con filas FMT de la matriz.
- **Desambiguación por nombre** cuando hay códigos duplicados (ej: SST-RGL-001 usado para REGLAMENTO COPASST y para REGLAMENTO DE HIGIENE).
- **Tolerancia a typos** (`ADQUISICION` ↔ `ADQUISISION`, `INSPECCION` ↔ `INSPECION`).
- **Fallback por nombre dentro del prefijo** cuando el código numérico no coincide (ej: matriz pide PRG-005 pero archivo es PRG-006 con mismo nombre).
- **Filtro por carpeta/numeral**: cuando hay duplicados, prefiere el archivo en la carpeta del numeral indicado en la matriz (ej: `4.1.2`).
- **Validación umbral mínimo**: no asigna link si la similitud nombre-archivo es < 0.40 (evita asignaciones incorrectas).
- **Detección de "Auxiliar de Reparto Motorizado"**: ya no se clasifica como peatón duplicado (solo como motociclista).
- **Fuente Arial 9** forzada en matriz vial (antes salía Calibri 10 heredado).
- **Texto negro en celdas de la matriz vial** (antes salía blanco/invisible en algunas celdas heredadas).
- **Alturas uniformes** de filas en matriz vial (antes el FACTOR SOCIAL salía comprimido).
- **Colores uniformes en columna R** de la matriz vial (mismo rojo/naranja/verde sin variaciones).
- **SEGURIDAD VIAL se conserva** cuando se elimina PESV (aplica para todas las empresas por Resolución 1565/2014).

### Cambiado
- **Agrupación por proceso en matriz IPVR**: cargos del mismo proceso quedan juntos aunque vengan mezclados en el YAML.
- **Merge de columnas A y B** (PROCESO y ZONA) extendido a todos los cargos del mismo proceso, no solo al cargo individual.

---

## [8.0] — 2026-04

### Agregado
- **Matriz de Peligros IPVR** (`SST-MTZ-012`): generación automática según cargos del cliente (catálogo base GTC 45 + extras por tipo).
- **Lectura de manuales de funciones**: detecta el campo DENOMINACIÓN del manual `.docx` para identificar el cargo.
- **Negrillas en matriz IPVR**: PROCESO, CARGO, TAREA RUTINARIA, CLASIFICACIÓN, ACEPTABILIDAD DEL RIESGO.
- **Colores alternados azul/verde** por bloque de cargo, durazno fijo en columnas PROCESO/ZONA.
- **Humanización con Claude API** para formatos (no solo políticas/procedimientos), adaptando COPASST↔VIGÍA.

### Corregido
- **Reglas ortográficas** en la salida de Claude (~80 palabras SST con tildes correctas).
- **Detección de títulos** por estilo de párrafo y por patrón visual.
- **Bibliografía y documentos relacionados** ya no se humanizan.
- **Definiciones**: cada término en negrilla, separados por línea.

---

## [7.x] — 2026-03

### Agregado
- Subida automática a Google Drive (`drive_service.py`).
- Inyección de links `=HYPERLINK("url","texto")` en SST-MTZ-005.
- Adaptación automática COPASST ↔ Vigía según número de trabajadores.
- Colores corporativos aplicados a Word y Excel.
- Reporte de auditoría (`REPORTE_AUDITORIA_SST.txt`).

---

## Cómo registrar un cambio

Cuando se hace un cambio significativo, agregar línea en la sección **No publicado** bajo:
- **Agregado** — funcionalidad nueva
- **Cambiado** — comportamiento de algo existente cambia
- **Obsoleto** — marca a remover en versión futura
- **Removido** — funcionalidad eliminada
- **Corregido** — bug fix
- **Seguridad** — cambios relacionados con vulnerabilidades

Al lanzar una versión:
1. Mover lo de **No publicado** a una sección con la versión y fecha
2. Crear nueva sección **No publicado** vacía
