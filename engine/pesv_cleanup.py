"""Limpieza centralizada de hojas relacionadas con PESV.

Cuando el cliente NO solicita Plan Estrategico de Seguridad Vial (PESV),
hay que eliminar TODAS las hojas relacionadas en cualquier matriz Excel
del SG-SST generado.

Hojas a eliminar (case-insensitive, con o sin tilde, con espacios extra):
  - PESV
  - SEGURIDAD VIAL
  - PLAN DE SEGURIDAD VIAL

Matrices conocidas con hojas PESV:
  - SST-MTZ-005 MATRIZ DE GESTION DOCUMENTAL  -> hoja "PESV"
  - SST-MTZ-006 MATRIZ DE REQUISITOS LEGALES  -> hojas "SEGURIDAD VIAL " y "PESV"

Esto se aplica de forma general: el modulo recorre TODOS los .xlsx del SG-SST
generado y elimina las hojas que coincidan, asi futuras matrices PESV se
limpian automaticamente sin tocar codigo.
"""
from __future__ import annotations

import logging
import warnings
from pathlib import Path

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

# Patrones de nombres de hoja a eliminar cuando PESV no aplica.
# Solo se quita la hoja PESV propiamente dicha.
# La hoja "SEGURIDAD VIAL" SI se conserva (aplica para todas las empresas
# por la Resolucion 1565/2014 - Plan Vial Empresarial basico).
_PATRONES_PESV = (
    "pesv",
    "plan estrategico de seguridad vial",
    "plan estrategico seguridad vial",
)


def _normalizar(nombre: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFD", str(nombre or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.strip().lower()


def _es_hoja_pesv(nombre_hoja: str) -> bool:
    n = _normalizar(nombre_hoja)
    # Excluir explicitamente "SEGURIDAD VIAL" (sin PESV) - esa hoja se conserva
    if n == "seguridad vial":
        return False
    return any(p in n for p in _PATRONES_PESV)


def limpiar_hojas_pesv_en_matrices(output_base: Path | str) -> dict[str, list[str]]:
    """Elimina las hojas PESV / SEGURIDAD VIAL de todos los .xlsx del SG-SST.

    Args:
        output_base: Carpeta raiz del SG-SST generado del cliente.

    Returns:
        Dict {nombre_archivo: [hojas_eliminadas, ...]} con un resumen de
        que se quito y donde.
    """
    output_base = Path(output_base)
    if not output_base.exists():
        return {}

    resumen: dict[str, list[str]] = {}

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for xlsx in output_base.rglob("*.xlsx"):
            if xlsx.name.startswith("~$"):
                continue
            try:
                wb = load_workbook(str(xlsx))
            except Exception as e:
                logger.warning("No se pudo abrir %s: %s", xlsx, e)
                continue

            hojas_a_borrar = [s for s in wb.sheetnames if _es_hoja_pesv(s)]
            if not hojas_a_borrar:
                continue

            # Excel no permite quedarse sin hojas; saltar si seria el caso
            if len(hojas_a_borrar) >= len(wb.sheetnames):
                logger.warning(
                    "No se eliminan hojas PESV de %s porque dejaria el libro vacio.",
                    xlsx.name,
                )
                continue

            for hoja in hojas_a_borrar:
                del wb[hoja]
                logger.info("Hoja '%s' eliminada de %s", hoja, xlsx.name)

            try:
                wb.save(str(xlsx))
                resumen[xlsx.name] = hojas_a_borrar
            except Exception as e:
                logger.warning("No se pudo guardar %s tras limpiar PESV: %s", xlsx, e)

    return resumen
