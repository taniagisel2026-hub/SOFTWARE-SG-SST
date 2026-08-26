# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

CLI tool that generates personalized SG-SST (Occupational Health & Safety Management System) documents for Colombian companies, compliant with Decreto 1072/2015 and Resolución 0312/2019. Built for internal use by MEDINA SST EMPRESARIAL S.A.S.

---

## Commands

```cmd
# Install dependencies
pip install -r requirements.txt

# Generate a client's SG-SST
python cli.py generar clientes_config\MI_CLIENTE.yaml

# Upload generated SG-SST to client's Google Drive
python cli.py subir-drive clientes_config\MI_CLIENTE.yaml

# Validate a client YAML
python cli.py verificar clientes_config\MI_CLIENTE.yaml

# Process a single document (debug)
python cli.py probar-doc <documento.docx> clientes_config\MI_CLIENTE.yaml

# Run all tests
python -m pytest

# Run a single test file
python -m pytest tests/test_listado_maestro.py -v

# Run tests by marker (skip slow/integration)
python -m pytest -m "not slow and not integration"

# Clear Claude response cache
python limpiar_cache.py

# Verify Drive links in the master list
python verificar_links_listado.py "clientes/<RAZON_SOCIAL>"

# Run the web API + wizard portal locally
uvicorn api:app --reload --port 8000
# Interactive API docs: http://localhost:8000/docs
# Wizard portal (client-facing, 8-step onboarding): http://localhost:8000/portal/
```

---

## Architecture

### Web layer: `api.py` + `wizard_api.py` (second entry point, alongside the CLI)

`api.py` is a FastAPI REST API that wraps the same `engine/` functions `cli.py` uses — it does not replace the CLI, it complements it. Endpoints (all except `/health` and the Drive OAuth callback require `Authorization: Bearer <API_AUTH_TOKEN>`, checked by `api_auth.verify_api_key`):

- `POST /clientes/{cliente_id}/generar`, `POST /clientes/{cliente_id}/subir-drive` — async, backed by `engine/jobs.py` (SQLite-persisted job status, survives restarts; `GET /jobs/{job_id}` polls status).
- `POST /clientes/validar`, `POST /clientes/verificar` — wrap `engine/validacion.py`.
- `GET /auth/drive/{cliente_id}/start` + `GET /auth/drive/callback` — OAuth 2.0 web flow for Drive (the callback is unauthenticated by design, protected instead by Google's single-use `code`).

`api.py` also mounts `wizard_router` (defined in `wizard_api.py`, a separate router by design to avoid regression risk on the Phase-4 routes above) and serves the static wizard portal at `/portal/` (`static/wizard/index.html` → `modulo1.html`...`modulo8.html` → `finalizar.html`). The wizard is a guided, client-facing onboarding flow that builds up a "borrador" (draft, persisted via `engine/wizard_store.py` in its own SQLite db) module by module — uploads (RUT, Cámara de Comercio, cédula, logo, firma) are extracted with Claude via `engine/extraccion.py`, gated sequentially (`modulo_N` can't be confirmed before `modulo_{N-1}`), and assembled into a client YAML on `POST /wizard/borradores/{id}/finalizar`, which then feeds the same `generar_sgsst()` pipeline as the CLI.

### CLI pipeline

The pipeline is orchestrated by `engine/generador.py::generar_sgsst()` and driven by `cli.py` (Click).

```
cli.py → engine/generador.py → engine/cliente.py   (YAML → contexto de marcadores)
                             → engine/docx_engine.py  (Word: reemplaza {{}} + humaniza)
                             → engine/xlsx_engine.py  (Excel: reemplaza {{}} + logos)
                             → engine/claude_service.py (Anthropic API + diskcache)
                             → engine/normativa.py    (correcciones legales)
                             → engine/colores.py      (colores corporativos)
                             → engine/pie_pagina.py   (posición del número de página)
                             → engine/estilo_encabezado.py (reacomoda logo/título/código/versión/fecha, .docx y .xlsx)
                             → engine/tipo_letra.py   (fuente de todo el documento, .docx y .xlsx)
                             → engine/pesv_cleanup.py (elimina hojas PESV si no aplica)
                             → engine/matriz_peligros.py     (IPVR SST-MTZ-012)
                             → engine/matriz_peligros_viales.py (vial PESV-MTZ-002)
                             → engine/auditoria.py    (reporte final)

python cli.py subir-drive → engine/drive_service.py → engine/listado_maestro.py
```

**Key flow**: `cliente.py` reads the YAML and expands every field into multiple variants (uppercase, lowercase, no tildes, underscores) so that `{{RAZON_SOCIAL}}`, `{{razon_social}}`, and `{{RAZÓN_SOCIAL}}` all resolve correctly from the same YAML value.

### Critical module: `engine/listado_maestro.py`

Fills `SST-MTZ-005` with Drive hyperlinks using fuzzy matching. The algorithm normalizes codes (`FTM` → `FMT`), computes Jaccard + containment similarity between matrix row names and uploaded filenames, and requires ≥0.40 score to assign a link. Fallback by prefix+name requires ≥0.75. Each production bug must become a regression test in `tests/test_listado_maestro.py`.

### Humanization levels (defined in `engine/config.py::DOC_TYPES`)

- `"completa"` — policies, programs, manuals, guides (fully rewritten by Claude)
- `"media"` — procedures, formats, plans, regulations (adapted but structure preserved)
- `"ninguna"` — matrices (no Claude call, only marker replacement)

### COPASST ↔ Vigía adaptation

`engine/generador.py` automatically switches between COPASST (≥10 workers) and Vigía (<10 workers) based on the `num_trabajadores` field in the client YAML.

### Claude API cache

`diskcache` in `.cache/claude_responses/` with 30-day TTL. Key = hash(text + context + humanization level). Run `python limpiar_cache.py` to reset.

### Google Drive tokens

OAuth 2.0 tokens per client stored in `tokens/token_<NIT>.json`. First run opens browser for authorization.

### Document style options (optional YAML fields, applied late in the pipeline)

`pie_pagina_numero`, `estilo_encabezado`, `tipo_letra` — all default to `""` (template's original look) and only run if set. `estilo_encabezado` and `tipo_letra` apply to both `.docx` and `.xlsx`; Excel headers are located by **text search** (a cell starting with "CÓDIGO" in the first ~8 rows) rather than a fixed column, since the header block's column position varies per matrix type. Matrices without a detectable header are skipped without error. Configurable live in the client portal's Módulo 8, with a preview that renders the user's real uploaded logo, versión, and fecha.

---

## Extending the system

**New YAML field**: add to `clientes_config/ejemplo_cliente.yaml`, expand in `engine/cliente.py::generar_contexto()`, add to `engine/claude_service.py::_construir_contexto()` if relevant to Claude.

**New IPVR peligro**: append to `PELIGROS_BASE_ADMINISTRATIVO` or `EXTRAS_POR_TIPO` in `engine/matriz_peligros.py`. Add a regression test.

**New matrix document type**: add to `TIPOS_PESTANA` list in `engine/listado_maestro.py`.

**New prefix equivalence** (historical typo): add to `_PREFIJOS_EQUIVALENTES` in `engine/listado_maestro.py`.

---

## Environment variables (`.env` in project root)

| Variable | Default | Purpose |
|---|---|---|
| `ANTHROPIC_API_KEY` | _(required)_ | Claude API key |
| `CLAUDE_MODEL` | `claude-haiku-4-5-20251001` | Claude model to use |
| `PLANTILLA_PATH` | `Plantilla SG-SST 2026` | Master template folder |
| `OUTPUT_BASE_PATH` | `clientes` | Output folder for generated SG-SSTs |
| `CACHE_DIR` | `.cache` | Claude cache directory |
| `CACHE_EXPIRY_DAYS` | `30` | Cache TTL |
| `CLAUDE_PRECIO_INPUT` / `CLAUDE_PRECIO_OUTPUT` | `1` / `5` | $/MTok for cost tracking |
| `API_AUTH_TOKEN` | `""` | Bearer token required by `api.py`/`wizard_api.py` endpoints |
| `API_BASE_URL` | `http://localhost:8000` | Base URL used to build the Drive OAuth redirect |
| `JOBS_DB_PATH` | `jobs.db` | SQLite db for `engine/jobs.py` (async generate/subir-drive job status) |
| `WIZARD_DB_PATH` | `wizard_drafts.db` | SQLite db for `engine/wizard_store.py` (wizard portal drafts) |
| `WIZARD_UPLOADS_DIR` | `wizard_uploads` | Where the wizard portal stores files clients upload |

---

## Files not to commit

`credentials.json`, `.env`, `tokens/`, `clientes/`, `.cache/`, `jobs.db`, `wizard_drafts.db`, `wizard_uploads/`

---

## Known limitations

- **OneDrive sync conflicts**: close or pause OneDrive before running the generator.
- **Open files in Excel/Word**: Windows file locks cause `PermissionError`; close all Office apps first.
- **Accented filenames on Windows**: uses `engine/utils.py::get_long_path()` (`\\?\C:\...`) to handle long paths; console output may show garbled characters (cosmetic only).

---

## Test policy

- `tests/conftest.py` — shared fixtures (mock clients, temp template copies, Claude stub)
- Markers: `slow` (>5 s), `integration` (requires Drive/Claude API)
- Every production bug → regression test before fix
