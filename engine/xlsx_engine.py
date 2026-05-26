"""Motor de procesamiento de archivos .xlsx - FUSIONADO v7.6 + nuevo."""
import os
import re
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font

from .utils import (
    get_long_path, looks_like_image_path, download_image_if_url, resolve_image_path
)

COLOR_NEGRO_ARGB = "FF000000"
ANCHO_LOGO_EXCEL = 120
ALTO_LOGO_EXCEL = 60


def procesar_xlsx(
    plantilla_path: Path,
    output_path: Path,
    contexto: dict,
    temp_dir: str,
    logo_path: str = None,
) -> dict:
    """Procesa un archivo .xlsx reemplazando marcadores con lógica del v7.

    Args:
        plantilla_path: Ruta al archivo .xlsx plantilla
        output_path: Ruta de salida
        contexto: Dict con variantes de marcadores expandidas
        temp_dir: Directorio temporal para imágenes
        logo_path: No usado en xlsx, incluido por consistencia

    Returns:
        Dict con celdas_modificadas y residuos
    """
    ruta_origen = get_long_path(str(plantilla_path))
    ruta_destino = get_long_path(str(output_path))

    wb = openpyxl.load_workbook(ruta_origen)
    celdas_modificadas = 0
    residuos = []

    # Preparar marcadores con llaves {{KEY}}
    marcadores_con_llaves = {}
    for k, v in contexto.items():
        llave = "{{" + k + "}}"
        if isinstance(v, str):
            marcadores_con_llaves[llave] = v

    fallback_dirs = [os.path.dirname(ruta_origen), temp_dir]

    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        for fila in ws.iter_rows():
            for celda in fila:
                if not celda.value or not isinstance(celda.value, str):
                    continue

                valor_nuevo = celda.value
                hubo_reemplazo = False

                for marcador, valor in marcadores_con_llaves.items():
                    if marcador in valor_nuevo:
                        valor_nuevo = valor_nuevo.replace(marcador, valor)
                        hubo_reemplazo = True

                if hubo_reemplazo:
                    es_imagen = False
                    if looks_like_image_path(valor_nuevo):
                        img_path = download_image_if_url(valor_nuevo, temp_dir)
                        if not img_path:
                            img_path = resolve_image_path(valor_nuevo, fallback_dirs)
                        if img_path and os.path.exists(img_path):
                            es_imagen = True
                            celda.value = ""
                            try:
                                img_xl = XLImage(img_path)
                                img_xl.width = ANCHO_LOGO_EXCEL
                                img_xl.height = ALTO_LOGO_EXCEL
                                ws.add_image(img_xl, celda.coordinate)
                                celdas_modificadas += 1
                            except Exception:
                                celda.value = valor_nuevo

                    if not es_imagen:
                        celda.value = valor_nuevo
                        celdas_modificadas += 1
                        # Forzar color negro
                        f = celda.font
                        if f:
                            celda.font = Font(
                                name=f.name, size=f.size, bold=f.bold, italic=f.italic,
                                underline=f.underline, strike=f.strike, vertAlign=f.vertAlign,
                                color=COLOR_NEGRO_ARGB,
                            )

                # Detectar residuos
                if isinstance(celda.value, str) and ("{{" in celda.value or "}}" in celda.value):
                    residuos.append(
                        f"  [Hoja '{nombre_hoja}' celda {celda.coordinate}] -> '{str(celda.value)[:80]}'"
                    )

    os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)
    wb.save(ruta_destino)

    return {
        "celdas_modificadas": celdas_modificadas,
        "residuos": residuos,
    }


def listar_marcadores_en_xlsx(xlsx_path: Path) -> list:
    """Lista todos los marcadores {{...}} encontrados en un archivo Excel."""
    wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
    marcadores = set()
    patron = re.compile(r"\{\{[A-ZÑ_]+\}\}")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    marcadores.update(patron.findall(cell.value))

    wb.close()
    return sorted(marcadores)
