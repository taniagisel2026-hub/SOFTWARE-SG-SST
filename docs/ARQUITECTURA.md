# Arquitectura técnica

Documento para desarrolladores que necesiten **mantener, depurar o extender** el sistema.

> Si eres consultor SST sin formación técnica, esto NO es para ti. Lee `GUIA_CONSULTOR.md`.

---

## Diagrama general

```
┌─────────────────────────────────────────────────────────────────┐
│  USUARIO (Consultor SST)                                         │
│  • Edita: clientes_config/MI_CLIENTE.yaml                        │
│  • Ejecuta: python cli.py generar / subir-drive                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  cli.py (Click)                                                  │
│  Define comandos: generar, subir-drive, verificar, etc.          │
└─────────────────────────────────────────────────────────────────┘
                              │
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
| 5 | Aplicar correcciones legales al normograma | `normativa.py` |
| 5b | Adaptar COPASST ↔ Vigía según N° trabajadores | `generador.py` |
| 5c | Aplicar colores corporativos del cliente | `colores.py` |
| 6 | Si `generar_pesv: false` → eliminar carpeta PESV + limpiar hojas PESV de matrices | `pesv_cleanup.py` |
| 6b | Generar Matriz IPVR `SST-MTZ-012` | `matriz_peligros.py` |
| 6c | Generar Matriz Vial `PESV-MTZ-002` (si aplica) | `matriz_peligros_viales.py` |
| 7 | Generar reporte de auditoría | `auditoria.py` |
| 8 | Limpiar temporales | shutil |

Al terminar, en `clientes/<RAZON_SOCIAL>/` queda la estructura PHVA completa más el `REPORTE_AUDITORIA_SST.txt`.

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

---

## Archivos sensibles (no commitear)

```
credentials.json           # OAuth Google Cloud
.env                       # API keys
tokens/                    # Tokens de Drive de cada cliente
clientes/                  # SG-SST generados (datos sensibles de clientes)
.cache/                    # Respuestas cacheadas de Claude
```

Sugerencia para `.gitignore`:

```
credentials.json
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
- Principales: `python-docx`, `docxtpl`, `openpyxl`, `anthropic`, `click`, `rich`, `pyyaml`, `diskcache`, `google-api-python-client`

---

## Limitaciones conocidas

1. **OneDrive**: trabajar en una carpeta sincronizada con OneDrive puede causar conflictos cuando OneDrive intenta sincronizar archivos mientras Python los escribe. Solución: cerrar OneDrive temporalmente o trabajar en una carpeta local.

2. **Archivos abiertos en Excel/Word**: Windows bloquea el archivo. El generador falla con `PermissionError`. Solución: cerrar todos los Word/Excel antes de generar.

3. **Manuales con formato muy distinto al esperado**: si el manual no tiene tabla con DENOMINACIÓN/DEPENDENCIA, el sistema usa el nombre del archivo como denominación. Funciona pero pierde precisión en la inferencia de actor vial.

4. **Tildes en nombres de archivo Windows**: archivos como `INSTALACIÓN.docx` pueden dar problemas si la consola usa codificación cp1252. El sistema usa rutas largas (`\\?\C:\...`) en `engine/utils.py::get_long_path()` para mitigar, pero algunos prints en consola pueden salir con caracteres raros (solo cosmético).

---

## Roadmap (siguientes mejoras sugeridas)

- [ ] Tests adicionales: `docx_engine`, `xlsx_engine`, `cliente`, `drive_service` (mock)
- [ ] Modo `--dry-run` para `subir-drive` (simular sin tocar Drive)
- [ ] Migrar la subida a Drive a async para acelerar
- [ ] Endpoint REST (Fase 4 del plan original)
- [ ] Frontend web (Fase 5)
- [ ] OCR de RUT y Cámara de Comercio (Fase 6)
