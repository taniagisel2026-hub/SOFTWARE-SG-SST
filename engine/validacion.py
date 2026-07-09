"""Validacion y verificacion del YAML de cliente, compartida por cli.py y api.py.

Extraido de cli.py (_validar_yaml y el comando verificar) para que ambas
interfaces (CLI y API REST) usen exactamente la misma logica sin duplicarla.
"""
from __future__ import annotations

from pathlib import Path

from engine.cliente import cargar_cliente, generar_contexto, obtener_imagenes

CLAVES_REQUERIDAS = ("razon_social", "nit")


def validar_cliente_dict(cliente: dict | None, nombre_archivo: str = "el archivo") -> list[str]:
    """Valida un dict de cliente ya cargado. Retorna la lista de problemas (vacia si es valido)."""
    if not isinstance(cliente, dict):
        return [f"{nombre_archivo} no contiene un YAML valido (se esperaba un diccionario)."]

    faltantes = [k for k in CLAVES_REQUERIDAS if not cliente.get(k)]
    if faltantes:
        return [
            f"Faltan claves obligatorias en {nombre_archivo}: {', '.join(faltantes)}. "
            f"Revisa el archivo de configuracion."
        ]
    return []


def cargar_y_validar_cliente(config_path: str | Path) -> tuple[dict | None, list[str]]:
    """Carga el YAML del cliente y lo valida. Retorna (cliente_o_None, problemas)."""
    config_path = Path(config_path)
    cliente = cargar_cliente(config_path)
    problemas = validar_cliente_dict(cliente, config_path.name)
    return (cliente if not problemas else None, problemas)


def construir_reporte_verificacion(cliente: dict, config_path: str | Path) -> dict:
    """Arma el reporte de verificacion (marcadores, imagenes, opciones) como datos estructurados."""
    config_path = Path(config_path)
    contexto = generar_contexto(cliente)
    imagenes = obtener_imagenes(cliente, config_path.parent)

    campos = []
    vistos: set[str] = set()
    for marcador in sorted(contexto.keys()):
        upper_key = marcador.upper()
        if upper_key in vistos:
            continue
        vistos.add(upper_key)
        valor = contexto[marcador]
        estado = "ok" if valor and str(valor).strip() else "vacio"
        campos.append({"marcador": marcador, "valor": str(valor), "estado": estado})

    marcadores_vacios = [m for m, v in contexto.items() if not v or not str(v).strip()]
    marcadores_vacios_unicos = sorted(set(m.upper() for m in marcadores_vacios))

    return {
        "razon_social": cliente.get("razon_social", "N/A"),
        "campos": campos,
        "marcadores_vacios": marcadores_vacios_unicos,
        "imagenes": {k: str(v) for k, v in imagenes.items()},
        "opciones": {
            "pesv": bool(cliente.get("generar_pesv")),
            "normograma_actualizado": bool(cliente.get("generar_normograma_actualizado")),
            "sedes": len(cliente.get("sedes", [])),
        },
    }
