"""Reporte de auditoría del SG-SST generado - del v7.6."""
from datetime import datetime
from pathlib import Path


def generar_reporte(ruta_salida: Path, stats: dict, config_info: dict = None) -> Path:
    """Genera un reporte de auditoría detallado en formato TXT.

    Args:
        ruta_salida: Directorio donde se guardará el reporte
        stats: Estadísticas de la generación
        config_info: Info del cliente para el reporte

    Returns:
        Ruta al archivo de reporte generado
    """
    ruta_reporte = ruta_salida / "REPORTE_AUDITORIA_SST.txt"
    total = stats.get("docx_ok", 0) + stats.get("xlsx_ok", 0)
    sep = "=" * 70

    with open(str(ruta_reporte), "w", encoding="utf-8") as f:
        f.write(sep + "\n")
        f.write("      REPORTE DE AUDITORIA - GENERADOR SG-SST v8.5\n")
        f.write(sep + "\n")
        f.write(f"  Fecha y hora    : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        if config_info:
            f.write(f"  Empresa         : {config_info.get('razon_social', 'N/A')}\n")
            f.write(f"  NIT             : {config_info.get('nit', 'N/A')}\n")
            f.write(f"  Nivel riesgo    : {config_info.get('nivel_riesgo', 'N/A')}\n")

        f.write(f"  Ruta de destino : {ruta_salida}\n")
        f.write(sep + "\n\n")

        # Resumen
        f.write("  RESUMEN DE PROCESAMIENTO\n")
        f.write("  " + "-" * 46 + "\n")
        f.write(f"  Archivos Word  (.docx) procesados : {stats.get('docx_ok', 0):>4}\n")
        f.write(f"  Archivos Excel (.xlsx) procesados : {stats.get('xlsx_ok', 0):>4}\n")
        f.write(f"  Archivos copiados sin cambios     : {stats.get('copiados', 0):>4}\n")
        f.write(f"  Archivos humanizados con Claude   : {stats.get('humanizados', 0):>4}\n")
        f.write(f"  Archivos con errores              : {stats.get('errores', 0):>4}\n")
        f.write(f"  TOTAL PROCESADOS                  : {total + stats.get('copiados', 0):>4}\n\n")
        f.write("  " + "-" * 46 + "\n")

        # Alertas (residuos)
        alertas = stats.get("alertas", {})
        if alertas:
            f.write(f"\n  ALERTAS: {len(alertas)} archivo(s) con marcadores SIN reemplazar\n\n")
            for archivo, datos in alertas.items():
                f.write(f"  >> {archivo}\n")
                for linea in datos:
                    f.write(f"      {linea}\n")
                f.write("\n")
        else:
            f.write("\n  SIN ALERTAS: Todos los marcadores fueron reemplazados.\n\n")

        # Variables faltantes
        faltantes = stats.get("faltantes_global", set())
        if faltantes:
            f.write("  " + "-" * 46 + "\n")
            f.write(f"  VARIABLES EN WORD SIN DATO EN EXCEL ({len(faltantes)} unicos):\n\n")
            for var in sorted(faltantes):
                f.write(f"  -> {{{{{var}}}}}\n")
            f.write("\n")

        # Errores
        errores_lista = stats.get("errores_lista", [])
        if errores_lista:
            f.write("  " + "-" * 46 + "\n")
            f.write(f"  ERRORES DE PROCESAMIENTO ({len(errores_lista)}):\n\n")
            for archivo, error in errores_lista:
                f.write(f"  Archivo : {archivo}\n")
                f.write(f"  Error   : {error}\n\n")

        # Detalle por archivo
        detalle = stats.get("detalle", [])
        if detalle:
            f.write("  " + "-" * 46 + "\n")
            f.write("  DETALLE POR ARCHIVO:\n\n")
            for tipo, archivo, info in detalle:
                tiene_alerta = archivo in alertas
                estado = "!!" if tiene_alerta else "OK"
                f.write(f"  [{estado}] [{tipo}] {archivo}\n")
                f.write(f"         {info}\n")

        f.write("\n" + sep + "\n  Fin del reporte.\n" + sep + "\n")

    return ruta_reporte
