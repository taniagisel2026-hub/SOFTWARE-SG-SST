"""API REST del generador de SG-SST (Fase 4).

Segunda puerta de entrada a las mismas funciones de engine/ que ya usa
cli.py -- no reemplaza al CLI, lo complementa. Pensada para eventualmente
correr en un servidor (de ahi el flujo OAuth "web" en engine/drive_service.py
y la persistencia de jobs en SQLite), pero por ahora se corre y se prueba
localmente:

    uvicorn api:app --reload --port 8000

Documentacion interactiva en http://localhost:8000/docs
"""
from __future__ import annotations

from pathlib import Path

from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from api_auth import verify_api_key
from engine import jobs
from engine.cliente import cargar_cliente
from engine.config import API_BASE_URL, OUTPUT_BASE_PATH
from engine.normativa import listar_correcciones
from engine.validacion import validar_cliente_dict, construir_reporte_verificacion
from wizard_api import wizard_router

__version__ = "8.5.0"

app = FastAPI(
    title="SG-SST Generator API",
    version=__version__,
    description="API REST para generar SG-SST, subirlos a Drive y consultar su estado.",
)
app.include_router(wizard_router)
app.mount(
    "/portal",
    StaticFiles(directory=Path(__file__).parent / "static" / "wizard", html=True),
    name="wizard-ui",
)

protegido = [Depends(verify_api_key)]


@app.on_event("startup")
def _al_arrancar() -> None:
    n = jobs.marcar_jobs_huerfanos_como_fallidos()
    if n:
        print(f"[api] {n} job(s) interrumpido(s) por un reinicio anterior, marcados como 'failed'.")


# ---------------------------------------------------------------------------
# Modelos de request
# ---------------------------------------------------------------------------

class ConfigPathRequest(BaseModel):
    config_path: str


class GenerarRequest(BaseModel):
    config_path: str
    output_dir: str | None = None
    humanizar: bool = True
    solo_marcadores: bool = False
    estricto: bool = False


class SubirDriveRequest(BaseModel):
    config_path: str
    directorio: str | None = None
    actualizar_listado: bool = True


# ---------------------------------------------------------------------------
# Salud y normas (sin estado, bajo riesgo)
# ---------------------------------------------------------------------------

@app.get("/health")
def health() -> dict:
    return {"status": "ok", "version": __version__}


@app.get("/normas", dependencies=protegido)
def normas() -> dict:
    correcciones = listar_correcciones()
    return {"correcciones": correcciones, "total": len(correcciones)}


# ---------------------------------------------------------------------------
# Validacion / verificacion de cliente
# ---------------------------------------------------------------------------

def _cargar_o_404(config_path: str) -> dict:
    path = Path(config_path)
    if not path.exists():
        raise HTTPException(404, f"No se encontro el archivo: {config_path}")
    return cargar_cliente(path)


@app.post("/clientes/validar", dependencies=protegido)
def clientes_validar(body: ConfigPathRequest) -> dict:
    cliente = _cargar_o_404(body.config_path)
    problemas = validar_cliente_dict(cliente, Path(body.config_path).name)
    return {
        "valido": not problemas,
        "problemas": problemas,
        "cliente": cliente if not problemas else None,
    }


@app.post("/clientes/verificar", dependencies=protegido)
def clientes_verificar(body: ConfigPathRequest) -> dict:
    cliente = _cargar_o_404(body.config_path)
    problemas = validar_cliente_dict(cliente, Path(body.config_path).name)
    if problemas:
        raise HTTPException(400, problemas[0])
    return construir_reporte_verificacion(cliente, body.config_path)


# ---------------------------------------------------------------------------
# Generacion (trabajo en segundo plano)
# ---------------------------------------------------------------------------

@app.post("/clientes/{cliente_id}/generar", status_code=202, dependencies=protegido)
def clientes_generar(cliente_id: str, body: GenerarRequest, background_tasks: BackgroundTasks) -> dict:
    cliente = _cargar_o_404(body.config_path)
    problemas = validar_cliente_dict(cliente, Path(body.config_path).name)
    if problemas:
        raise HTTPException(400, problemas[0])

    job_id = jobs.crear_job("generar", cliente_id, body.model_dump())
    background_tasks.add_task(
        jobs.ejecutar_generacion_en_background,
        job_id,
        body.config_path,
        body.output_dir,
        body.humanizar,
        body.solo_marcadores,
        body.estricto,
    )
    return {"job_id": job_id, "status": "queued"}


# ---------------------------------------------------------------------------
# Subida a Drive (trabajo en segundo plano)
# ---------------------------------------------------------------------------

@app.post("/clientes/{cliente_id}/subir-drive", status_code=202, dependencies=protegido)
def clientes_subir_drive(cliente_id: str, body: SubirDriveRequest, background_tasks: BackgroundTasks) -> dict:
    cliente = _cargar_o_404(body.config_path)
    problemas = validar_cliente_dict(cliente, Path(body.config_path).name)
    if problemas:
        raise HTTPException(400, problemas[0])

    if body.directorio:
        output_dir = Path(body.directorio)
    else:
        razon_social = cliente.get("razon_social", "CLIENTE")
        nombre_carpeta = razon_social.replace(" ", "_").replace(".", "")
        output_dir = OUTPUT_BASE_PATH / nombre_carpeta

    if not output_dir.exists():
        raise HTTPException(
            400,
            f"No se encontro el directorio {output_dir}. "
            "Primero genera el SG-SST con POST /clientes/{id}/generar.",
        )

    job_id = jobs.crear_job("subir-drive", cliente_id, body.model_dump())
    background_tasks.add_task(
        jobs.ejecutar_subida_en_background,
        job_id,
        cliente,
        str(output_dir),
        cliente_id,
        body.actualizar_listado,
    )
    return {"job_id": job_id, "status": "queued"}


# ---------------------------------------------------------------------------
# Estado de jobs
# ---------------------------------------------------------------------------

@app.get("/jobs/{job_id}", dependencies=protegido)
def jobs_estado(job_id: str) -> dict:
    job = jobs.obtener_job(job_id)
    if job is None:
        raise HTTPException(404, f"No existe el job: {job_id}")
    return job


@app.get("/clientes/{cliente_id}/jobs", dependencies=protegido)
def clientes_jobs(cliente_id: str, tipo: str | None = None) -> dict:
    return {"jobs": jobs.listar_jobs(cliente_id, tipo)}


# ---------------------------------------------------------------------------
# Login de Google Drive (flujo web, para cuando la API corre en un servidor)
# ---------------------------------------------------------------------------

def _redirect_uri() -> str:
    return f"{API_BASE_URL}/auth/drive/callback"


@app.get("/auth/drive/{cliente_id}/start", dependencies=protegido)
def auth_drive_start(cliente_id: str, redirect: bool = True):
    from engine.drive_service import iniciar_auth_web

    try:
        auth_url = iniciar_auth_web(cliente_id, _redirect_uri())
    except FileNotFoundError as e:
        raise HTTPException(500, str(e))

    if redirect:
        return RedirectResponse(auth_url)
    return {"auth_url": auth_url}


@app.get("/auth/drive/callback", response_class=HTMLResponse)
def auth_drive_callback(code: str, state: str) -> str:
    """Google redirige aqui tras el consentimiento. `state` trae el cliente_id,
    firmado por iniciar_auth_web() (ver completar_auth_web/_verificar_state).

    Sin autenticacion por API key -- la seguridad viene del `code` de Google,
    que es de un solo uso e irrepetible (practica estandar de OAuth), y de la
    firma del `state` (que solo pudo generar quien ya tenia la API key, en
    /auth/drive/{cliente_id}/start).
    """
    import html
    from engine.drive_service import completar_auth_web

    try:
        cliente_id, email = completar_auth_web(state, code, _redirect_uri())
    except Exception as e:
        return f"<h1>Error autorizando Drive</h1><p>{html.escape(str(e))}</p>"

    return (
        f"<h1>Autorizacion completada</h1>"
        f"<p>Cliente: {html.escape(cliente_id)}</p>"
        f"<p>Cuenta de Google Drive conectada: <b>{html.escape(email)}</b></p>"
        f"<p>Si esa NO es la cuenta correcta, vuelve a hacer clic en "
        f"\"Autorizar Google Drive\" y elige la cuenta correcta en el paso de Google.</p>"
        f"<p>Ya puedes cerrar esta pestana.</p>"
    )
