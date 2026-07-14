"""Asistente de onboarding por modulos (Fase 5), montado dentro de api.py.

Router separado (en vez de crecer api.py directamente) porque el numero de
endpoints del asistente (subidas + confirmar por modulo + finalizar) es
grande y las rutas de la Fase 4 (generar/subir-drive/jobs/normas) no deben
tener ningun riesgo de regresion por este cambio.

Reusa la misma autenticacion por API key que ya existe (api_auth.verify_api_key)
-- esto NO es un sistema de login por cliente, sigue siendo el mismo token
estatico de la Fase 4. Login real por cliente es un tema aparte, mas grande.
"""
from __future__ import annotations

import shutil
from pathlib import Path

import yaml
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from pydantic import BaseModel

from api_auth import verify_api_key
from engine import extraccion, pdf_utils, wizard_defaults, wizard_store
from engine.cliente import formatear_fecha_hoy
from engine.config import CLIENTES_CONFIG_PATH, WIZARD_UPLOADS_DIR
from engine.validacion import validar_cliente_dict

wizard_router = APIRouter(prefix="/wizard", tags=["wizard"])

protegido = [Depends(verify_api_key)]

# Roles de archivo del modulo 1 que se extraen con Claude (PDF -> texto -> JSON)
ROLES_EXTRAIBLES = {"rut": "rut", "camara_comercio": "camara_comercio", "cedula_rl": "cedula"}
# Roles de archivo del modulo 1 que solo se suben, sin extraccion
ROLES_IMAGEN_MODULO_1 = ("logo", "firma_rl")


def _obtener_o_404(borrador_id: str) -> dict:
    borrador = wizard_store.obtener_borrador(borrador_id)
    if borrador is None:
        raise HTTPException(404, f"No existe el borrador: {borrador_id}")
    return borrador


def _verificar_gate(borrador: dict, modulo_key: str) -> None:
    """Lanza 409 si el modulo anterior a `modulo_key` no esta confirmado todavia."""
    idx = wizard_store.MODULOS_ORDEN.index(modulo_key)
    if idx == 0:
        return
    anterior = wizard_store.MODULOS_ORDEN[idx - 1]
    if anterior not in borrador["modulos_confirmados"]:
        raise HTTPException(
            409, f"Debes confirmar {anterior} antes de continuar con {modulo_key}."
        )


def _guardar_archivo_subido(borrador_id: str, rol: str, upload: UploadFile) -> Path:
    carpeta = WIZARD_UPLOADS_DIR / borrador_id
    carpeta.mkdir(parents=True, exist_ok=True)
    ext = Path(upload.filename or "").suffix or ""
    destino = carpeta / f"{rol}{ext}"
    with open(destino, "wb") as f:
        shutil.copyfileobj(upload.file, f)
    return destino


# ---------------------------------------------------------------------------
# Crear / consultar / eliminar borrador
# ---------------------------------------------------------------------------

@wizard_router.post("/borradores", dependencies=protegido)
def crear_borrador() -> dict:
    borrador_id = wizard_store.crear_borrador()
    return wizard_store.obtener_borrador(borrador_id)


@wizard_router.get("/borradores/{borrador_id}", dependencies=protegido)
def ver_borrador(borrador_id: str) -> dict:
    return _obtener_o_404(borrador_id)


@wizard_router.delete("/borradores/{borrador_id}", dependencies=protegido)
def eliminar_borrador(borrador_id: str) -> dict:
    _obtener_o_404(borrador_id)
    carpeta = WIZARD_UPLOADS_DIR / borrador_id
    if carpeta.exists():
        shutil.rmtree(carpeta)
    wizard_store.eliminar_borrador(borrador_id)
    return {"eliminado": True}


# ---------------------------------------------------------------------------
# Modulo 1 -- Documentos legales (subida + extraccion + confirmacion)
# ---------------------------------------------------------------------------

@wizard_router.post("/borradores/{borrador_id}/modulo-1/archivos", dependencies=protegido)
async def modulo1_subir_archivos(
    borrador_id: str,
    rut: UploadFile | None = File(None),
    camara_comercio: UploadFile | None = File(None),
    cedula_rl: UploadFile | None = File(None),
    logo: UploadFile | None = File(None),
    firma_rl: UploadFile | None = File(None),
) -> dict:
    """Sube cualquier subconjunto de los 5 archivos del modulo 1. Los PDFs
    (rut/camara_comercio/cedula_rl) se extraen de inmediato con Claude; las
    imagenes (logo/firma_rl) solo se guardan."""
    _obtener_o_404(borrador_id)
    subidos = {
        "rut": rut, "camara_comercio": camara_comercio, "cedula_rl": cedula_rl,
        "logo": logo, "firma_rl": firma_rl,
    }

    for rol, upload in subidos.items():
        if upload is None:
            continue
        destino = _guardar_archivo_subido(borrador_id, rol, upload)
        wizard_store.guardar_archivo(borrador_id, rol, str(destino))
        if rol in ROLES_EXTRAIBLES:
            texto = pdf_utils.extraer_texto_pdf(destino)
            extraido = extraccion.extraer_datos_documento(ROLES_EXTRAIBLES[rol], texto)
            wizard_store.guardar_extraccion(borrador_id, rol, extraido)

    borrador = wizard_store.obtener_borrador(borrador_id)
    conflictos = extraccion.detectar_conflictos(borrador["extraccion"])
    return {
        "extraccion": borrador["extraccion"],
        "conflictos": conflictos,
        "archivos": borrador["archivos"],
    }


@wizard_router.get("/borradores/{borrador_id}/modulo-1/extraccion", dependencies=protegido)
def modulo1_ver_extraccion(borrador_id: str) -> dict:
    """Reconsulta la extraccion ya guardada, sin volver a subir archivos
    (util para recargar la pagina sin perder el progreso)."""
    borrador = _obtener_o_404(borrador_id)
    conflictos = extraccion.detectar_conflictos(borrador["extraccion"])
    return {
        "extraccion": borrador["extraccion"],
        "conflictos": conflictos,
        "archivos": borrador["archivos"],
    }


class Modulo1ConfirmRequest(BaseModel):
    razon_social: str
    nit: str
    direccion_empresa: str = ""
    ciudad_empresa: str = ""
    departamento_empresa: str = ""
    barrio_empresa: str = ""
    actividad_economica: str = ""
    representante_legal: str = ""
    cedula_rl: str = ""
    cargo_rl: str = ""
    ciudad_cedula_rl: str = ""
    email_empresa: str = ""
    sitio_web: str = ""
    telefono_empresa: str = ""


@wizard_router.post("/borradores/{borrador_id}/modulo-1/confirmar", dependencies=protegido)
def modulo1_confirmar(borrador_id: str, body: Modulo1ConfirmRequest) -> dict:
    """El usuario ya revisó/corrigió lo extraído -- esto es lo que queda
    guardado, sin importar lo que digan los documentos originales."""
    _obtener_o_404(borrador_id)
    datos = body.model_dump()
    problemas = validar_cliente_dict(datos, "modulo_1")
    if problemas:
        raise HTTPException(400, problemas[0])
    return wizard_store.guardar_modulo(borrador_id, "modulo_1", datos)


# ---------------------------------------------------------------------------
# Modulo 2 -- Riesgo y ARL
# ---------------------------------------------------------------------------

class Modulo2Request(BaseModel):
    nivel_riesgo: str
    arl: str
    numero_trabajadores_empresa: str
    color_corporativo: str


@wizard_router.post("/borradores/{borrador_id}/modulo-2/confirmar", dependencies=protegido)
def modulo2_confirmar(borrador_id: str, body: Modulo2Request) -> dict:
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_2")
    return wizard_store.guardar_modulo(borrador_id, "modulo_2", body.model_dump())


# ---------------------------------------------------------------------------
# Modulo 3 -- Sedes
# ---------------------------------------------------------------------------

class Sede(BaseModel):
    nombre: str
    direccion: str = ""
    ciudad: str = ""
    telefono: str = ""


class Modulo3Request(BaseModel):
    sedes: list[Sede]


@wizard_router.post("/borradores/{borrador_id}/modulo-3/confirmar", dependencies=protegido)
def modulo3_confirmar(borrador_id: str, body: Modulo3Request) -> dict:
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_3")
    datos = {"sedes": [s.model_dump() for s in body.sedes]}
    return wizard_store.guardar_modulo(borrador_id, "modulo_3", datos)


# ---------------------------------------------------------------------------
# Modulo 4 -- Responsable SST y consultora (+ firma)
# ---------------------------------------------------------------------------

class Modulo4Request(BaseModel):
    nombre_responsable_sst: str
    cedula_responsable_sst: str
    cargo_sst: str
    ciudad_cedula_sst: str
    licencia_sst: str
    empresa_consultora_sst: str
    nit_consultora_sst: str
    licencia_empresa_consultora: str


@wizard_router.post("/borradores/{borrador_id}/modulo-4/confirmar", dependencies=protegido)
def modulo4_confirmar(borrador_id: str, body: Modulo4Request) -> dict:
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_4")
    return wizard_store.guardar_modulo(borrador_id, "modulo_4", body.model_dump())


@wizard_router.post("/borradores/{borrador_id}/modulo-4/archivos", dependencies=protegido)
async def modulo4_subir_archivos(
    borrador_id: str,
    firma_sst: UploadFile | None = File(None),
    hoja_vida_sst: UploadFile | None = File(None),
    licencia_sst_pdf: UploadFile | None = File(None),
    cedula_sst_pdf: UploadFile | None = File(None),
    diploma_cursos_sst: UploadFile | None = File(None),
) -> dict:
    """Sube cualquier subconjunto de los archivos del modulo 4: firma del
    responsable (imagen, se inserta en los documentos) + hoja de vida,
    licencia, cedula y diploma de cursos (PDF, soporte documental -- se
    copian tal cual al SG-SST generado, no se procesan)."""
    _obtener_o_404(borrador_id)
    subidos = {
        "firma_sst": firma_sst,
        "hoja_vida_sst": hoja_vida_sst,
        "licencia_sst_pdf": licencia_sst_pdf,
        "cedula_sst_pdf": cedula_sst_pdf,
        "diploma_cursos_sst": diploma_cursos_sst,
    }
    destinos = {}
    for rol, upload in subidos.items():
        if upload is None:
            continue
        destino = _guardar_archivo_subido(borrador_id, rol, upload)
        wizard_store.guardar_archivo(borrador_id, rol, str(destino))
        destinos[rol] = str(destino)
    return {"archivos": destinos}


# ---------------------------------------------------------------------------
# Modulo 5 -- Fechas, version y opciones
# ---------------------------------------------------------------------------

class Modulo5Request(BaseModel):
    version: str = "1"
    fecha_emision: str = ""
    fecha: str = ""
    año: str = ""
    ciudad: str = ""
    mes_año: str = ""
    generar_pesv: bool = False


@wizard_router.post("/borradores/{borrador_id}/modulo-5/confirmar", dependencies=protegido)
def modulo5_confirmar(borrador_id: str, body: Modulo5Request) -> dict:
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_5")
    return wizard_store.guardar_modulo(borrador_id, "modulo_5", body.model_dump())


# ---------------------------------------------------------------------------
# Modulo 6 -- Descripcion de la empresa
# ---------------------------------------------------------------------------

class Modulo6Request(BaseModel):
    descripcion_empresa: str


@wizard_router.post("/borradores/{borrador_id}/modulo-6/confirmar", dependencies=protegido)
def modulo6_confirmar(borrador_id: str, body: Modulo6Request) -> dict:
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_6")
    return wizard_store.guardar_modulo(borrador_id, "modulo_6", body.model_dump())


# ---------------------------------------------------------------------------
# Modulo 7 -- Cargos (opcional, para la matriz de peligros IPVR)
# ---------------------------------------------------------------------------
# Dos formas de completarlo (igual que ya soporta generar_sgsst() por CLI):
#   1) Subir manuales de funciones (.docx/.xlsx/.pdf) -- el nombre del cargo
#      se detecta automaticamente del contenido del manual, no hay que
#      escribirlo. Termina siendo `manuales_cargos_dir` en el YAML final.
#   2) Escribir la lista de cargos a mano. Termina siendo `cargos` en el YAML.
# O simplemente omitirlo -- la matriz IPVR no se genera, sin error.

class Cargo(BaseModel):
    nombre: str
    proceso: str = "ADMINISTRATIVO"
    zona: str = ""
    tipo: str = "administrativo"  # administrativo | operativo | conductor | servicios_generales
    num_expuestos: int = 1
    actividades: str = ""
    tareas: str = ""


class Modulo7ListaRequest(BaseModel):
    cargos: list[Cargo]


@wizard_router.post("/borradores/{borrador_id}/modulo-7/manuales", dependencies=protegido)
async def modulo7_subir_manuales(borrador_id: str, archivos: list[UploadFile] = File(...)) -> dict:
    """Sube uno o mas manuales de funciones. Cada archivo se guarda en
    wizard_uploads/<borrador_id>/manuales/ -- esa carpeta se convierte en
    manuales_cargos_dir al finalizar (ver _ensamblar_cliente)."""
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_7")

    carpeta = WIZARD_UPLOADS_DIR / borrador_id / "manuales"
    carpeta.mkdir(parents=True, exist_ok=True)
    extensiones_validas = {".docx", ".xlsx", ".pdf"}
    guardados = []
    for archivo in archivos:
        ext = Path(archivo.filename or "").suffix.lower()
        if ext not in extensiones_validas:
            raise HTTPException(400, f"Formato no soportado: {archivo.filename} (usa .docx, .xlsx o .pdf)")
        destino = carpeta / Path(archivo.filename).name
        with open(destino, "wb") as f:
            shutil.copyfileobj(archivo.file, f)
        guardados.append(destino.name)

    return wizard_store.guardar_modulo(
        borrador_id, "modulo_7", {"modo": "manuales", "manuales_subidos": guardados}
    )


@wizard_router.post("/borradores/{borrador_id}/modulo-7/lista", dependencies=protegido)
def modulo7_lista_cargos(borrador_id: str, body: Modulo7ListaRequest) -> dict:
    """Alternativa a subir manuales: escribir los cargos a mano."""
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_7")
    if not body.cargos:
        raise HTTPException(400, "Agrega al menos un cargo, o usa /modulo-7/saltar si no aplica.")
    datos = {"modo": "lista", "cargos": [c.model_dump() for c in body.cargos]}
    return wizard_store.guardar_modulo(borrador_id, "modulo_7", datos)


@wizard_router.post("/borradores/{borrador_id}/modulo-7/saltar", dependencies=protegido)
def modulo7_saltar(borrador_id: str) -> dict:
    """generar_sgsst() ya soporta que ni 'cargos' ni 'manuales_cargos_dir'
    esten presentes -- simplemente omite la matriz de peligros IPVR."""
    borrador = _obtener_o_404(borrador_id)
    _verificar_gate(borrador, "modulo_7")
    datos = {"modo": "omitido"}
    return wizard_store.guardar_modulo(borrador_id, "modulo_7", datos)


# ---------------------------------------------------------------------------
# Valores por defecto (pre-llenado)
# ---------------------------------------------------------------------------

@wizard_router.get("/borradores/{borrador_id}/defaults/modulo-4", dependencies=protegido)
def defaults_modulo4(borrador_id: str) -> dict:
    _obtener_o_404(borrador_id)
    return wizard_defaults.DEFAULTS_CONSULTORA


@wizard_router.get("/borradores/{borrador_id}/defaults/modulo-5", dependencies=protegido)
def defaults_modulo5(borrador_id: str) -> dict:
    _obtener_o_404(borrador_id)
    hoy = formatear_fecha_hoy()
    return {
        "version": "001",
        "fecha_emision": hoy["fecha_emision"],
        "fecha": hoy["fecha"],
        "año": hoy["año"],
        "ciudad": "Bogotá D.C.",
        "mes_año": f"{hoy['mes_actual']}, {hoy['año']}",
        "generar_pesv": False,
    }


# ---------------------------------------------------------------------------
# Finalizar -- ensambla el YAML y entrega el handoff a /clientes/{id}/generar
# ---------------------------------------------------------------------------

def _ensamblar_cliente(borrador: dict) -> dict:
    datos = borrador["datos"]
    cliente: dict = {}
    cliente.update(datos.get("modulo_1", {}))
    cliente.update(datos.get("modulo_2", {}))
    if "modulo_3" in datos:
        cliente["sedes"] = datos["modulo_3"].get("sedes", [])
    cliente.update(datos.get("modulo_4", {}))
    cliente.update(datos.get("modulo_5", {}))
    cliente.update(datos.get("modulo_6", {}))
    # modulo_7 (cargos): modo "lista" -> cargos: [...] directo en el YAML.
    # modo "manuales" se resuelve en finalizar() (necesita copiar archivos
    # y conocer el slug del cliente primero). modo "omitido" no agrega nada.
    if datos.get("modulo_7", {}).get("modo") == "lista":
        cliente["cargos"] = datos["modulo_7"]["cargos"]
    return cliente


@wizard_router.post("/borradores/{borrador_id}/finalizar", dependencies=protegido)
def finalizar(borrador_id: str) -> dict:
    borrador = _obtener_o_404(borrador_id)

    # modulo_7 es opcional (se puede omitir), todos los demas son obligatorios
    requeridos = [m for m in wizard_store.MODULOS_ORDEN if m != "modulo_7"]
    faltantes = [m for m in requeridos if m not in borrador["modulos_confirmados"]]
    if faltantes:
        raise HTTPException(409, f"Faltan modulos por confirmar: {', '.join(faltantes)}")

    cliente = _ensamblar_cliente(borrador)
    problemas = validar_cliente_dict(cliente, "borrador")
    if problemas:
        raise HTTPException(400, problemas[0])

    slug = cliente["razon_social"].replace(" ", "_").replace(".", "")

    # Copiar logo/firmas del borrador (desechable) a clientes_config/ (permanente),
    # con rutas relativas al YAML -- exactamente lo que obtener_imagenes() espera.
    assets_dir = CLIENTES_CONFIG_PATH / f"{slug}_assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    for rol in (
        "logo", "firma_rl", "firma_sst",
        "hoja_vida_sst", "licencia_sst_pdf", "cedula_sst_pdf", "diploma_cursos_sst",
    ):
        origen_str = borrador["archivos"].get(rol)
        if not origen_str:
            continue
        origen = Path(origen_str)
        if not origen.exists():
            continue
        destino = assets_dir / origen.name
        shutil.copyfile(origen, destino)
        cliente[rol] = f"{slug}_assets/{origen.name}"

    # Modulo 7, modo "manuales": copiar la carpeta de manuales subida (desechable)
    # a clientes_config/<slug>_assets/manuales_cargos/ (permanente) y apuntar
    # manuales_cargos_dir ahi -- cargar_cargos_desde_manuales() la resuelve
    # relativa al propio YAML (igual que las imagenes).
    modulo_7 = borrador["datos"].get("modulo_7", {})
    if modulo_7.get("modo") == "manuales":
        origen_manuales = WIZARD_UPLOADS_DIR / borrador_id / "manuales"
        if origen_manuales.exists():
            destino_manuales = assets_dir / "manuales_cargos"
            if destino_manuales.exists():
                shutil.rmtree(destino_manuales)
            shutil.copytree(origen_manuales, destino_manuales)
            cliente["manuales_cargos_dir"] = f"{slug}_assets/manuales_cargos"

    config_path = CLIENTES_CONFIG_PATH / f"{slug}.yaml"
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(cliente, f, allow_unicode=True, sort_keys=False)

    wizard_store.marcar_finalizado(borrador_id, str(config_path))

    return {
        "config_path": str(config_path),
        "cliente_id": slug,
        "siguiente_paso": f"POST /clientes/{slug}/generar",
    }
