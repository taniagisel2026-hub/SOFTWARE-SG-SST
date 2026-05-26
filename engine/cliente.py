"""Carga y gestión de datos del cliente - FUSIONADO v7 + nuevo."""
import yaml
import openpyxl
from pathlib import Path
from datetime import datetime

from .utils import norm_key, key_variants, format_value, get_long_path


def cargar_cliente(yaml_path: str | Path) -> dict:
    """Carga la configuración del cliente desde un archivo YAML."""
    with open(yaml_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data


def leer_marcadores_excel(ruta_excel: str) -> dict:
    """Lee marcadores desde un archivo MARCADORES.xlsx (compatibilidad v7).

    Retorna un contexto expandido con todas las variantes.
    """
    ruta_excel = get_long_path(ruta_excel)
    wb = openpyxl.load_workbook(ruta_excel, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    wb.close()

    if not rows:
        return {}

    primera_fila = [str(c).strip() if c is not None else "" for c in rows[0]]
    idx_marcador = None
    idx_valor = None

    for i, h in enumerate(primera_fila):
        h_up = h.upper()
        if "MARCADOR" in h_up and idx_marcador is None:
            idx_marcador = i
        if "DATO" in h_up and idx_valor is None:
            idx_valor = i

    if idx_marcador is not None and idx_valor is not None:
        datos_rows = rows[1:]
    else:
        idx_marcador = 0
        idx_valor = 1
        datos_rows = rows

    contexto_base = {}
    for fila in datos_rows:
        raw_m = fila[idx_marcador] if len(fila) > idx_marcador else None
        raw_v = fila[idx_valor] if len(fila) > idx_valor else None

        if raw_m is None:
            continue

        raw_m_str = str(raw_m).strip()
        clave_original = raw_m_str.replace("{{", "").replace("}}", "").strip()
        marcador = norm_key(raw_m_str)
        if not marcador or marcador.lower() == "none":
            continue

        valor_formateado = format_value(raw_v)
        contexto_base[marcador] = valor_formateado

        if clave_original != marcador:
            contexto_base[clave_original] = valor_formateado
            contexto_base[clave_original.upper()] = valor_formateado

    # Expandir con variantes
    contexto = {}
    for k, v in contexto_base.items():
        if k not in contexto:
            contexto[k] = v
        for variante in key_variants(k):
            if variante not in contexto:
                contexto[variante] = v

    return contexto


def _parsear_numero_trabajadores(valor) -> int:
    """Convierte el número de trabajadores a entero de forma segura.

    Acepta int, str con dígitos, str con texto como '15 trabajadores', etc.
    """
    if isinstance(valor, int):
        return valor
    try:
        return int(str(valor).strip())
    except ValueError:
        # Intentar extraer el primer número del string
        import re
        match = re.search(r"\d+", str(valor))
        return int(match.group()) if match else 0


def generar_contexto(cliente: dict) -> dict:
    """Convierte los datos del cliente YAML al contexto expandido para docxtpl.

    Genera el mismo formato que leer_marcadores_excel pero desde YAML.
    Incluye todas las variantes de cada clave.
    """
    now = datetime.now()
    meses = [
        "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    mes_actual = meses[now.month].capitalize()
    año = str(cliente.get("año", now.year))

    # Determinar COPASST o Vigía según número de trabajadores
    # Decreto 1072 de 2015, Art. 2.2.4.6.12:
    #   < 10 trabajadores → Vigía de SST
    #   >= 10 trabajadores → COPASST (Comité Paritario de SST)
    num_trabajadores = _parsear_numero_trabajadores(
        cliente.get("numero_trabajadores_empresa")
        or cliente.get("num_trabajadores", 0)
    )
    es_copasst = num_trabajadores >= 10
    copasst_o_vigia = "COPASST" if es_copasst else "Vigía de SST"
    copasst_o_vigia_completo = (
        "Comité Paritario de Seguridad y Salud en el Trabajo (COPASST)"
        if es_copasst
        else "Vigía de Seguridad y Salud en el Trabajo"
    )

    # Mapeo base de marcadores
    base = {
        "VERSION": str(cliente.get("version", "1")),
        "FECHA_EMISION": cliente.get("fecha_emision", now.strftime("%d de %B de %Y")),
        "FECHA": cliente.get("fecha", now.strftime("%d de %B de %Y")),
        "CIUDAD_CEDULA_SST": cliente.get("ciudad_cedula_sst", ""),
        "CIUDAD_CEDULA_RL": cliente.get("ciudad_cedula_rl", ""),
        "CIUDAD": cliente.get("ciudad", ""),
        "MES_AÑO": f"{mes_actual}, {año}",
        "MES_ANO": f"{mes_actual}, {año}",
        "RAZON_SOCIAL": cliente.get("razon_social", ""),
        "NIT": str(cliente.get("nit", "")),
        "DIRECCION_EMPRESA": cliente.get("direccion_empresa", ""),
        "TELEFONO_EMPRESA": str(cliente.get("telefono_empresa", "")),
        "EMAIL_EMPRESA": cliente.get("email_empresa", ""),
        "SITIO_WEB": cliente.get("sitio_web", ""),
        "BARRIO_EMPRESA": cliente.get("barrio_empresa", ""),
        "CIUDAD_EMPRESA": cliente.get("ciudad_empresa", ""),
        "DEPARTAMENTO_EMPRESA": cliente.get("departamento_empresa", ""),
        "ACTIVIDAD_ECONOMICA": str(cliente.get("actividad_economica", "")).strip(),
        "NIVEL_RIESGO": str(cliente.get("nivel_riesgo", "")),
        "ARL": str(cliente.get("arl", "")),
        "NUMERO_TRABAJADORES_EMPRESA": str(cliente.get("numero_trabajadores_empresa", "")),
        "NOMBRE_RESPONSABLE_SST": cliente.get("nombre_responsable_sst", ""),
        "CEDULA_RESPONSABLE_SST": str(cliente.get("cedula_responsable_sst", "")),
        "LICENCIA_SST": cliente.get("licencia_sst", ""),
        "EMPRESA_CONSULTORA_SST": cliente.get("empresa_consultora_sst", ""),
        "NIT_CONSULTORA_SST": str(cliente.get("nit_consultora_sst", "")),
        "LICENCIA_EMPRESA_CONSULTORA": cliente.get("licencia_empresa_consultora", ""),
        "AÑO": año,
        "ANO": año,
        "REPRESENTANTE_LEGAL": cliente.get("representante_legal", ""),
        "CEDULA_RL": str(cliente.get("cedula_rl", "")),
        "CARGO_RL": cliente.get("cargo_rl", ""),
        "CARGO_SST": cliente.get("cargo_sst", ""),
        # COPASST o Vigía (derivado automático)
        "COPASST_O_VIGIA": copasst_o_vigia,
        "COPASST_VIGIA": copasst_o_vigia,
        "COPASST_O_VIGIA_COMPLETO": copasst_o_vigia_completo,
        "NOMBRE_COPASST_VIGIA": copasst_o_vigia_completo,
        # Imágenes - se pasan como rutas, docxtpl las resuelve
        "LOGO": cliente.get("logo", ""),
        "FIRMA_RL": cliente.get("firma_rl", ""),
        "FIRMA_SST": cliente.get("firma_sst", ""),
    }

    # Limpiar None
    base = {k: (v if v is not None else "") for k, v in base.items()}

    # Expandir con variantes (como v7)
    contexto = {}
    for k, v in base.items():
        if k not in contexto:
            contexto[k] = v
        for variante in key_variants(k):
            if variante not in contexto:
                contexto[variante] = v

    return contexto


def obtener_imagenes(cliente: dict, config_dir: Path) -> dict:
    """Retorna rutas absolutas a las imágenes del cliente."""
    imagenes = {}
    for key in ("logo", "firma_rl", "firma_sst"):
        ruta = cliente.get(key, "")
        if ruta and not ruta.startswith("ruta/"):
            p = Path(ruta)
            if not p.is_absolute():
                p = config_dir / p
            if p.exists():
                imagenes[key.upper()] = str(p)
    return imagenes


def clasificar_documento(filename: str) -> str | None:
    """Identifica el tipo de documento por su prefijo."""
    from .config import DOC_TYPES
    for code, info in DOC_TYPES.items():
        if info["prefijo"] in filename:
            return code
    return None


def nivel_humanizacion(filename: str) -> str:
    """Retorna el nivel de humanización requerido para un documento."""
    from .config import DOC_TYPES
    tipo = clasificar_documento(filename)
    if tipo and tipo in DOC_TYPES:
        return DOC_TYPES[tipo]["humanizar"]
    return "ninguna"
