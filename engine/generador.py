"""Orquestador principal: genera el SG-SST completo - FUSIONADO v7.6 + nuevo."""
import shutil
import tempfile
import warnings
from pathlib import Path

# Silenciar warnings de openpyxl (fechas fuera de rango, validaciones no soportadas)
# Son inofensivos: los valores se conservan, solo no se interpretan como fecha
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn

from .config import PLANTILLA_PATH, OUTPUT_BASE_PATH, DOC_TYPES
from .cliente import (
    cargar_cliente, generar_contexto, obtener_imagenes,
    nivel_humanizacion, clasificar_documento
)
from .docx_engine import procesar_docx, extraer_secciones_contenido, extraer_texto_docx
from .xlsx_engine import procesar_xlsx
from .claude_service import humanizar_secciones, actualizar_normograma, obtener_uso_sesion, resetear_uso_sesion
from .normativa import aplicar_correcciones_docx
from .colores import aplicar_colores_corporativos
from .auditoria import generar_reporte
from .utils import encontrar_ruta_logo

console = Console(highlight=False)

IGNORAR_PATRONES = {"~$", ".tmp", "Thumbs.db", ".DS_Store"}
NORMOGRAMA_CARPETA = "NORMAS"


def debe_ignorar(filename: str) -> bool:
    return any(filename.startswith(p) or filename.endswith(p) for p in IGNORAR_PATRONES)


def generar_sgsst(
    config_path: str | Path,
    output_dir: str | Path = None,
    humanizar: bool = True,
    solo_marcadores: bool = False,
    estricto: bool = False,
):
    """Genera el SG-SST completo para un cliente."""
    resetear_uso_sesion()
    console.print("\n[bold blue]=== GENERADOR SG-SST v8.5 ===[/bold blue]\n")

    # 1. Cargar configuracion
    console.print("[yellow]1.[/yellow] Cargando configuracion del cliente...")
    cliente = cargar_cliente(config_path)
    config_dir = Path(config_path).parent
    razon_social = cliente.get("razon_social", "CLIENTE")
    console.print(f"   Empresa: [bold]{razon_social}[/bold]")
    console.print(f"   NIT: {cliente.get('nit', 'N/A')}")
    console.print(f"   Actividad: {str(cliente.get('actividad_economica', 'N/A'))[:80]}...")
    console.print(f"   Nivel riesgo: {cliente.get('nivel_riesgo', 'N/A')} | ARL: {cliente.get('arl', 'N/A')}")
    console.print(f"   Trabajadores: {cliente.get('numero_trabajadores_empresa', 'N/A')}")

    # 2. Generar contexto expandido (con variantes como v7)
    console.print("\n[yellow]2.[/yellow] Preparando contexto de marcadores...")
    contexto = generar_contexto(cliente)
    imagenes = obtener_imagenes(cliente, config_dir)

    # Inyectar rutas de imágenes al contexto
    for key_img, ruta_img in imagenes.items():
        contexto[key_img] = ruta_img
        contexto[key_img.lower()] = ruta_img

    console.print(f"   {len(contexto)} variantes de marcadores generadas")
    console.print(f"   {len(imagenes)} imagenes encontradas")

    if humanizar and not solo_marcadores:
        console.print("   [green]Humanizacion con Claude API: ACTIVADA[/green]")
    else:
        console.print("   [dim]Humanizacion: desactivada[/dim]")

    # 3. Crear estructura de salida
    nombre_carpeta = razon_social.replace(" ", "_").replace(".", "")
    if output_dir:
        output_base = Path(output_dir)
    else:
        output_base = OUTPUT_BASE_PATH / nombre_carpeta

    console.print(f"\n[yellow]3.[/yellow] Creando estructura en: [cyan]{output_base}[/cyan]")

    # 4. Crear directorio temporal para imágenes
    temp_dir = tempfile.mkdtemp(prefix="sgsst_imgs_")
    logo_path = encontrar_ruta_logo(contexto, [str(PLANTILLA_PATH), temp_dir], temp_dir)

    if logo_path:
        console.print(f"   Logo para marca de agua: [green]encontrado[/green]")
    else:
        console.print(f"   Logo para marca de agua: [dim]no encontrado (se omite)[/dim]")

    # 5. Procesar archivos
    console.print(f"\n[yellow]4.[/yellow] Procesando plantilla desde: [cyan]{PLANTILLA_PATH}[/cyan]")

    # Determinar si usar versión VIGÍA o COPASST según número de trabajadores
    num_trab = _obtener_num_trabajadores(cliente)
    usar_vigia = num_trab < 10

    if usar_vigia:
        console.print(f"   [yellow]Usando versión VIGÍA DE SST ({num_trab} trabajadores)[/yellow]")
    else:
        console.print(f"   [yellow]Usando versión COPASST ({num_trab} trabajadores)[/yellow]")

    from .matriz_documental import determinar_categoria, normalizar_nivel_riesgo
    _nivel_riesgo = normalizar_nivel_riesgo(cliente.get("nivel_riesgo", "I"))
    categoria_res0312 = determinar_categoria(_nivel_riesgo, num_trab)
    console.print(f"   [yellow]Categoría Res. 0312: {categoria_res0312} ({_nivel_riesgo}, {num_trab} trabajadores)[/yellow]")

    archivos = _listar_archivos_plantilla(usar_vigia=usar_vigia, categoria=categoria_res0312)
    total = len(archivos)

    # Stats para auditoría (formato v7)
    stats = {
        "docx_ok": 0,
        "xlsx_ok": 0,
        "copiados": 0,
        "humanizados": 0,
        "errores": 0,
        "errores_lista": [],
        "alertas": {},
        "faltantes_global": set(),
        "detalle": [],
    }

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        console=console,
    ) as progress:
        task = progress.add_task("Procesando documentos...", total=total)

        for archivo in archivos:
            ruta_relativa = archivo.relative_to(PLANTILLA_PATH)
            # Eliminar tags de versión dual y de categoría del nombre de salida
            nombre_salida = (
                ruta_relativa.name
                .replace(" [COPASST]", "")
                .replace(" [VIGÍA]", "")
                .replace(" [CAT-I]", "")
                .replace(" [CAT-II]", "")
                .replace(" [CAT-III]", "")
            )
            output_path = output_base / ruta_relativa.parent / nombre_salida
            desc = nombre_salida[:50]

            try:
                if archivo.suffix.lower() == ".docx":
                    resultado = _procesar_archivo_docx(
                        archivo, output_path, contexto, temp_dir, logo_path,
                        cliente, humanizar, solo_marcadores,
                    )
                    stats["docx_ok"] += 1
                    stats["faltantes_global"].update(resultado.get("faltantes", []))

                    residuos = resultado.get("residuos", [])
                    faltantes = resultado.get("faltantes", [])
                    tvars = resultado.get("template_vars", [])
                    info = f"{len(tvars)} vars | {len(faltantes)} sin dato | {len(residuos)} residuos"
                    stats["detalle"].append(("DOCX", str(ruta_relativa), info))

                    if residuos:
                        stats["alertas"][str(ruta_relativa)] = residuos

                    if resultado.get("humanizado"):
                        stats["humanizados"] += 1

                elif archivo.suffix.lower() == ".xlsx":
                    resultado = procesar_xlsx(
                        archivo, output_path, contexto, temp_dir, logo_path
                    )
                    stats["xlsx_ok"] += 1

                    celdas = resultado.get("celdas_modificadas", 0)
                    residuos = resultado.get("residuos", [])
                    info = f"{celdas} celdas | {len(residuos)} residuos"
                    stats["detalle"].append(("XLSX", str(ruta_relativa), info))

                    if residuos:
                        stats["alertas"][str(ruta_relativa)] = residuos

                elif archivo.suffix.lower() == ".pptx":
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(str(archivo), str(output_path))
                    stats["copiados"] += 1
                else:
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(str(archivo), str(output_path))
                    stats["copiados"] += 1

            except Exception as e:
                stats["errores"] += 1
                stats["errores_lista"].append((str(ruta_relativa), str(e)))

            progress.update(task, advance=1, description=f"[dim]{desc}[/dim]")

    # 5b. Reparacion proactiva de .docx con Content_Types mal armado
    # (bug heredado de la plantilla con watermarks). Evita errores en los
    # pasos siguientes (correcciones legales, COPASST/Vigia, normograma).
    reparados = _reparar_docx_de_salida(output_base)
    if reparados > 0:
        console.print(f"   [dim]{reparados} archivo(s) .docx reparado(s) (Content_Types)[/dim]")

    # 6. Correcciones legales (siempre) + normograma con Claude (opcional)
    console.print("\n[yellow]5.[/yellow] Actualizando normas legales...")
    normas_corregidas = _aplicar_correcciones_legales(output_base)
    if normas_corregidas > 0:
        console.print(f"   [green]{normas_corregidas} correcciones legales aplicadas[/green]")
    else:
        console.print(f"   [dim]Sin correcciones legales necesarias[/dim]")

    if humanizar and not solo_marcadores and cliente.get("generar_normograma_actualizado", False):
        console.print("   Revisando normograma con Claude...")
        _actualizar_normograma_carpeta(output_base, cliente)
    stats["normas_corregidas"] = normas_corregidas

    # 6b. Adaptar contenido COPASST → Vigía en los documentos generados
    # La plantilla dual elige la carpeta correcta, pero el contenido interno de
    # los .docx de la carpeta VIGÍA SST sigue diciendo "COPASST" (solo se
    # renombraron los archivos). Esta función reemplaza el texto en el cuerpo.
    if usar_vigia:
        console.print(f"\n[yellow]6.[/yellow] Adaptando contenido de documentos: Vigía de SST ({num_trab} trabajadores)...")
        cambios_cv = _adaptar_copasst_vigia(output_base, es_copasst=False)
        console.print(f"   [green]{cambios_cv} reemplazos de texto aplicados en documentos[/green]")
    else:
        console.print(f"\n[yellow]6.[/yellow] COPASST aplica ({num_trab} trabajadores) - sin adaptación necesaria")

    # 6c. Colores corporativos
    color_corp = cliente.get("color_corporativo", "")
    if color_corp:
        console.print(f"\n[yellow]7.[/yellow] Aplicando colores corporativos: [bold]{color_corp}[/bold]...")
        color_stats = aplicar_colores_corporativos(output_base, color_corp)
        if color_stats["archivos"] > 0:
            console.print(
                f"   [green]{color_stats['archivos']} archivos actualizados[/green] "
                f"({color_stats['docx_cambios']} cambios Word, {color_stats['xlsx_cambios']} cambios Excel)"
            )
            console.print(f"   Paleta: {color_stats['paleta']}")
        else:
            console.print(f"   [dim]Sin cambios de color necesarios[/dim]")
    else:
        console.print(f"\n[yellow]7.[/yellow] Colores corporativos: [dim]usando plantilla original[/dim]")

    # 7. PESV
    if not cliente.get("generar_pesv", False):
        for pesv_name in ["PLAN ESTRATEGICO DE SEGURIDAD VIAL", "PLAN ESTRATÉGICO DE SEGURIDAD VIAL"]:
            pesv_dir = output_base / pesv_name
            if pesv_dir.exists():
                shutil.rmtree(str(pesv_dir))
                console.print(f"\n[yellow]8.[/yellow] PESV no aplica - carpeta eliminada")
                break
        else:
            console.print(f"\n[yellow]8.[/yellow] PESV no aplica")
        # Tambien eliminar las hojas PESV / SEGURIDAD VIAL de TODAS las matrices
        # afectadas (SST-MTZ-005 Gestion Documental, SST-MTZ-006 Matriz Legal, etc.)
        try:
            from .pesv_cleanup import limpiar_hojas_pesv_en_matrices
            hojas_eliminadas = limpiar_hojas_pesv_en_matrices(output_base)
            if hojas_eliminadas:
                for archivo, hojas in hojas_eliminadas.items():
                    console.print(
                        f"   [dim]Hoja(s) {', '.join(hojas)} eliminada(s) de {archivo}[/dim]"
                    )
        except Exception as e:
            console.print(f"   [yellow]No se pudieron limpiar hojas PESV de las matrices: {e}[/yellow]")
    else:
        console.print(f"\n[yellow]8.[/yellow] PESV incluido en la generacion")
        # Replicar la estructura COMPLETA de carpetas PESV (incluyendo las
        # subcarpetas vacias de FASE 2, 3, 4) para que el consultor las vea
        # como recordatorio de lo que falta por llenar.
        carpetas_creadas = _replicar_estructura_pesv(output_base)
        if carpetas_creadas:
            console.print(f"   [dim]{carpetas_creadas} carpetas PESV vacias replicadas (recordatorio)[/dim]")

    # 7c. Matriz de Peligros Viales (PESV-MTZ-002) si el cliente tiene PESV
    if cliente.get("generar_pesv", False) and (cliente.get("cargos") or cliente.get("manuales_cargos_dir")):
        console.print(f"\n[yellow]9.[/yellow] Generando Matriz de Peligros Viales (PESV-MTZ-002)...")
        try:
            from .matriz_peligros_viales import generar_matriz_peligros_viales
            ruta_vial = generar_matriz_peligros_viales(cliente, output_base, config_dir=config_dir)
            if ruta_vial:
                console.print(f"   [green]Matriz PESV-MTZ-002 llenada[/green]")
            else:
                console.print(
                    f"   [dim]Matriz vial PESV no aplica (ningún cargo tiene "
                    f"desplazamientos viales detectables: conductor, motociclista, "
                    f"peatón en diligencias, pasajero en visitas comerciales).[/dim]"
                )
        except Exception as e:
            console.print(f"   [red]Error generando matriz vial: {e}[/red]")

    # 7b. Matriz de Peligros (IPVR) SST-MTZ-012 segun cargos del YAML
    if cliente.get("cargos") or cliente.get("manuales_cargos_dir"):
        console.print(f"\n[yellow]10.[/yellow] Generando Matriz de Peligros IPVR (GTC 45)...")
        try:
            from .matriz_peligros import generar_matriz_peligros
            ruta_ipvr = generar_matriz_peligros(cliente, output_base, config_dir=config_dir)
            if ruta_ipvr:
                n_cargos = cliente.get("_cargos_procesados") or len(cliente.get("cargos", []))
                console.print(
                    f"   [green]Matriz IPVR llenada para {n_cargos} cargo(s)[/green]"
                )
            else:
                console.print(f"   [yellow]Matriz IPVR no se pudo llenar (ver logs)[/yellow]")
        except Exception as e:
            console.print(f"   [red]Error generando matriz IPVR: {e}[/red]")
    else:
        console.print(
            f"\n[yellow]10.[/yellow] Matriz IPVR: [dim]sin 'cargos' en YAML, se omite[/dim]"
        )

    # 8. Adaptar texto COPASST→Vigía en SST-MTZ-005 si aplica
    if usar_vigia:
        try:
            from .matriz_documental import _adaptar_vigia_xlsx
            mtz005 = next(
                (f for f in output_base.rglob("*.xlsx")
                 if "MTZ-005" in f.name and not f.name.startswith("~$")),
                None,
            )
            if mtz005:
                _adaptar_vigia_xlsx(mtz005)
                console.print(f"\n[yellow]11.[/yellow] SST-MTZ-005 adaptado: COPASST → Vigía SST")
        except Exception as e:
            console.print(f"\n[yellow]11.[/yellow] [yellow]No se pudo adaptar MTZ-005 para Vigía: {e}[/yellow]")

    # 9. Generar reporte de auditoría
    console.print("\n[yellow]12.[/yellow] Generando reporte de auditoria...")
    ruta_reporte = generar_reporte(output_base, stats, config_info=cliente)
    console.print(f"   Reporte: {ruta_reporte}")

    # 9. Limpiar temporal
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    # 10. Modo estricto: fallar si hay marcadores faltantes
    if estricto and stats["faltantes_global"]:
        faltantes_list = sorted(stats["faltantes_global"])
        console.print(f"\n[red bold]=== MODO ESTRICTO: GENERACION ABORTADA ===[/red bold]")
        console.print(f"   {len(faltantes_list)} marcador(es) sin dato en el YAML del cliente:\n")
        for var in faltantes_list[:15]:
            console.print(f"   [red]>>[/red] {{{{{var}}}}}")
        if len(faltantes_list) > 15:
            console.print(f"   ... y {len(faltantes_list) - 15} mas")
        console.print(f"\n   [yellow]Agrega estos campos al YAML o usa --no-estricto[/yellow]")
        # Limpiar salida incompleta
        try:
            shutil.rmtree(str(output_base), ignore_errors=True)
        except Exception:
            pass
        raise ValueError(
            f"Modo estricto: {len(faltantes_list)} marcadores sin dato. "
            f"Documentos no generados."
        )

    # 11. Resumen
    total_proc = stats["docx_ok"] + stats["xlsx_ok"]
    console.print(f"\n[bold green]=== GENERACION COMPLETADA ===[/bold green]")
    console.print(f"   Word procesados:  [green]{stats['docx_ok']}[/green]")
    console.print(f"   Excel procesados: [green]{stats['xlsx_ok']}[/green]")
    console.print(f"   Copiados:         {stats['copiados']}")
    console.print(f"   Humanizados:      [cyan]{stats['humanizados']}[/cyan]")

    if stats["alertas"]:
        console.print(f"   Con alertas:      [yellow]{len(stats['alertas'])}[/yellow] (ver reporte)")
    if stats["errores"]:
        console.print(f"   Errores:          [red]{stats['errores']}[/red]")
        for nombre, error in stats["errores_lista"][:3]:
            console.print(f"     [red]X[/red] {nombre}: {error}")

    if not stats["alertas"] and not stats["errores"]:
        console.print("   [bold green]TODOS LOS MARCADORES REEMPLAZADOS EXITOSAMENTE[/bold green]")

    # 12. Reporte de uso de Claude API
    uso = obtener_uso_sesion()
    if uso["llamadas"] > 0:
        console.print(f"\n   [bold]Uso Claude API:[/bold]")
        console.print(f"   Llamadas API:     {uso['llamadas']}")
        console.print(f"   Tokens input:     {uso['tokens_input']:,}")
        console.print(f"   Tokens output:    {uso['tokens_output']:,}")
        console.print(f"   Desde cache:      {uso['desde_cache']}")
        console.print(f"   [bold]Costo total:     ${uso['costo_total_usd']:.4f} USD[/bold]")
    elif uso["desde_cache"] > 0:
        console.print(f"\n   [dim]Claude API: {uso['desde_cache']} secciones desde cache (sin costo)[/dim]")

    console.print(f"   Salida: [cyan]{output_base}[/cyan]\n")

    return {
        "output_dir": str(output_base),
        "procesados": total_proc + stats["copiados"],
        "humanizados": stats["humanizados"],
        "total": total,
        "errores": stats["errores_lista"],
        "alertas": len(stats["alertas"]),
        "stats": stats,
        "uso_claude": uso,
    }


def _listar_archivos_plantilla(usar_vigia: bool = False, categoria: str = "III") -> list[Path]:
    """Lista archivos de la plantilla filtrando por versión COPASST/Vigía y categoría Res. 0312.

    Args:
        usar_vigia:  Si True, usa carpeta VIGÍA; si False, carpeta COPASST.
        categoria:   'I', 'II' o 'III' — selecciona el MTZ-005 correcto.
    """
    otras_cats = {"I", "II", "III"} - {categoria}

    archivos = []
    for archivo in PLANTILLA_PATH.rglob("*"):
        if not archivo.is_file() or debe_ignorar(archivo.name):
            continue

        ruta_str = str(archivo)

        # ── Filtro COPASST / VIGÍA ────────────────────────────────────────
        if usar_vigia:
            if "1.1.6 CONFORMACIÓN DEL COPASST" in ruta_str:
                continue
            if "1.1.7 CAPACITACIÓN COPASST" in ruta_str:
                continue
            if "[COPASST]" in archivo.name:
                continue
        else:
            if "1.1.6 DESIGNACIÓN DEL VIGÍA SST" in ruta_str:
                continue
            if "[VIGÍA]" in archivo.name:
                continue

        # ── Filtro por categoría Res. 0312 ────────────────────────────────
        # Excluir los [CAT-X] que NO corresponden a esta categoría
        for otra in otras_cats:
            if f"[CAT-{otra}]" in archivo.name:
                break
        else:
            archivos.append(archivo)

    return sorted(archivos)


def _replicar_estructura_pesv(output_base: Path) -> int:
    """Replica TODAS las subcarpetas de PLAN ESTRATEGICO DE SEGURIDAD VIAL
    al output, incluso si estan vacias en la plantilla.

    Necesario porque FASE 2, 3 y 4 de la plantilla aun no tienen documentos
    elaborados, pero queremos que aparezcan en el cliente/Drive como
    recordatorio de la estructura oficial del PESV (Resolucion 40595 de 2022).

    Returns:
        Numero de carpetas creadas (no contadas las que ya existian).
    """
    # La plantilla tiene este nombre con acento
    nombre_pesv_plantilla = "PLAN ESTRATÉGICO DE SEGURIDAD VIAL"
    plantilla_pesv = PLANTILLA_PATH / nombre_pesv_plantilla
    if not plantilla_pesv.exists():
        # Fallback sin acento
        plantilla_pesv = PLANTILLA_PATH / "PLAN ESTRATEGICO DE SEGURIDAD VIAL"
        if not plantilla_pesv.exists():
            return 0
        nombre_pesv_plantilla = "PLAN ESTRATEGICO DE SEGURIDAD VIAL"

    # Determinar el nombre que uso el output (puede variar si hubo renombrado)
    output_pesv = output_base / nombre_pesv_plantilla
    if not output_pesv.exists():
        # Buscar cualquier carpeta que empiece con "PLAN ESTRAT"
        for sub in output_base.iterdir():
            if sub.is_dir() and sub.name.upper().startswith("PLAN ESTRAT"):
                output_pesv = sub
                break
        else:
            # No existe ninguna carpeta PESV en el output - crearla
            output_pesv.mkdir(parents=True, exist_ok=True)

    carpetas_creadas = 0
    for sub in plantilla_pesv.rglob("*"):
        if not sub.is_dir():
            continue
        rel = sub.relative_to(plantilla_pesv)
        destino = output_pesv / rel
        if not destino.exists():
            destino.mkdir(parents=True, exist_ok=True)
            carpetas_creadas += 1
    return carpetas_creadas


def _procesar_archivo_docx(
    plantilla_path, output_path, contexto, temp_dir, logo_path,
    cliente, humanizar, solo_marcadores,
) -> dict:
    """Procesa un .docx individual. Retorna resultado con info de auditoría."""
    secciones_humanizadas = None

    if humanizar and not solo_marcadores:
        nivel = nivel_humanizacion(plantilla_path.name)
        if nivel != "ninguna":
            tipo = clasificar_documento(plantilla_path.name)
            tipo_nombre = DOC_TYPES.get(tipo, {}).get("nombre", "documento")

            secciones = extraer_secciones_contenido(plantilla_path, contexto)
            if secciones:
                # Pasar ruta relativa desde la carpeta de plantillas para que
                # Claude conozca el contexto de la carpeta padre (ej.
                # "1.1.8 COMITE DE CONVIVENCIA LABORAL/...") y no confunda
                # CCL con COPASST.
                try:
                    ruta_contexto = str(plantilla_path.relative_to(PLANTILLA_PATH))
                except ValueError:
                    ruta_contexto = plantilla_path.name

                secciones_humanizadas = humanizar_secciones(
                    secciones=secciones,
                    cliente=cliente,
                    nivel=nivel,
                    tipo_documento=tipo_nombre,
                    nombre_archivo=ruta_contexto,
                )

    resultado = procesar_docx(
        plantilla_path=plantilla_path,
        output_path=output_path,
        contexto=contexto,
        temp_dir=temp_dir,
        logo_path=logo_path,
        secciones_humanizadas=secciones_humanizadas,
    )

    resultado["humanizado"] = secciones_humanizadas is not None and len(secciones_humanizadas or []) > 0
    return resultado


def _reparar_docx_de_salida(output_base: Path) -> int:
    """Repara SOLO los .docx que no se pueden abrir por bug de Content_Types
    mal armado. Archivos que ya abren correctamente NO se tocan.

    Returns:
        Numero de archivos reparados.
    """
    from .normativa import _reparar_content_types_si_falta_watermark
    from docx import Document as _DocumentCheck

    reparados = 0
    for archivo in output_base.rglob("*.docx"):
        if archivo.name.startswith("~$"):
            continue
        try:
            _DocumentCheck(str(archivo))
        except Exception as e:
            if "content type" not in str(e).lower() and "partname" not in str(e).lower():
                continue
            try:
                if _reparar_content_types_si_falta_watermark(str(archivo)):
                    reparados += 1
            except Exception:
                pass
    return reparados


def _abrir_docx(archivo: Path):
    """Abre un .docx, reparando Content_Types automáticamente si es necesario."""
    from docx import Document
    from .normativa import _reparar_content_types_si_falta_watermark
    try:
        return Document(str(archivo))
    except Exception as e:
        if "content type" in str(e).lower() or "partname" in str(e).lower():
            if _reparar_content_types_si_falta_watermark(str(archivo)):
                return Document(str(archivo))
        raise


def _aplicar_correcciones_legales(output_base: Path) -> int:
    """Aplica correcciones legales deterministas a TODOS los .docx generados.

    Reemplaza normas derogadas/modificadas usando el mapa de correcciones
    verificadas en normativa.py. Gratis, exacto, sin API.

    Muestra detalle de cada corrección: archivo, ubicación, qué se cambió.

    Returns:
        Total de correcciones aplicadas
    """
    total = 0
    for archivo in output_base.rglob("*.docx"):
        if archivo.name.startswith("~$"):
            continue
        try:
            reporte = aplicar_correcciones_docx(str(archivo), detallado=True)
            if reporte.total_cambios > 0:
                total += reporte.total_cambios
                console.print(f"   [cyan]{archivo.name}[/cyan]: {reporte.total_cambios} corrección(es)")
                for cambio in reporte.detalle:
                    console.print(
                        f"      [{cambio.ubicacion}] "
                        f"[red]{cambio.texto_original}[/red] -> [green]{cambio.texto_nuevo}[/green]"
                    )
        except Exception as e:
            console.print(f"   [red]Error corrigiendo normas en {archivo.name}: {e}[/red]")
    return total


def _obtener_num_trabajadores(cliente: dict) -> int:
    """Obtiene el número de trabajadores del cliente de forma segura."""
    valor = (
        cliente.get("numero_trabajadores_empresa")
        or cliente.get("num_trabajadores", 0)
    )
    try:
        return int(str(valor).strip())
    except ValueError:
        import re as _re
        match = _re.search(r"\d+", str(valor))
        return int(match.group()) if match else 0


def _adaptar_copasst_vigia(output_base: Path, es_copasst: bool) -> int:
    """Reemplaza COPASST ↔ Vigía en el TEXTO de todos los .docx generados.

    Si es_copasst=False (< 10 trabajadores), reemplaza toda mención de
    COPASST por Vigía de SST dentro del contenido de los documentos.

    Conserva el formato original (opera sobre runs de python-docx).

    Returns:
        Número total de reemplazos realizados
    """
    import re as _re
    from docx import Document

    # Mapas de reemplazo según dirección
    if not es_copasst:
        # Empresa pequeña: COPASST → Vigía
        # IMPORTANTE: las reglas se aplican EN ORDEN. Las primeras son las
        # más específicas para evitar duplicaciones.
        # Nota sobre regex: "(?:la\s+)?" tolera "Comité Paritario de la Seguridad"
        # o "Comité Paritario de Seguridad".
        reemplazos = {
            # ============================================================
            # 1) LIMPIEZAS DE REDUNDANCIAS GENERADAS POR REEMPLAZOS PREVIOS
            # (van primero para limpiar cualquier resto de versiones anteriores)
            # ============================================================
            # "Comité Paritario ... (Vigía SST)" -> "Vigía SST"
            r"Comit[eé]\s+Paritario\s+de\s+(?:la\s+)?Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo"
            r"\s*\(\s*Vig[ií]a\s+(?:de\s+)?SST\s*\)":
                "Vigía SST",
            # "Comité Paritario ... - Vigía SST" -> "Vigía SST"
            r"Comit[eé]\s+Paritario\s+de\s+(?:la\s+)?Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo"
            r"\s*[-—–]\s*"
            r"Vig[ií]a\s+(?:de\s+)?SST":
                "Vigía SST",
            # "Vigía ... - Vigía ..." -> "Vigía SST"
            r"Vig[ií]a\s+de\s+(?:Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo|SST)"
            r"\s*[-—–]\s*"
            r"Vig[ií]a\s+de\s+(?:Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo|SST)":
                "Vigía SST",
            # "Vigía de Seguridad y Salud en el Trabajo (COPASST)" → "Vigía SST"
            r"Vig[ií]a\s+de\s+Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo\s*\(\s*COPASST\s*\)":
                "Vigía SST",

            # ============================================================
            # 2) "COPASST / Vigía" (la plantilla a veces los pone juntos)
            # ============================================================
            r"COPASST\s*[/oó]\s*Vig[ií]a\s*(?:de\s*)?SST": "Vigía SST",
            r"COPASST\s*[/oó]\s*Vig[ií]a": "Vigía SST",

            # ============================================================
            # 3) Nombre completo "Comité Paritario de Seguridad y Salud..."
            # Se reemplaza por el nombre completo correcto del Vigía.
            # ============================================================
            # Versión con paréntesis (COPASST)
            r"Comit[eé]\s+Paritario\s+de\s+(?:la\s+)?Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo\s*\(\s*COPASST\s*\)":
                "Vigía de Seguridad y Salud en el Trabajo",
            r"COMIT[EÉ]\s+PARITARIO\s+DE\s+(?:LA\s+)?SEGURIDAD\s+Y\s+SALUD\s+EN\s+EL\s+TRABAJO\s*\(\s*COPASST\s*\)":
                "Vigía de Seguridad y Salud en el Trabajo",
            # Versión sin paréntesis
            r"Comit[eé]\s+Paritario\s+de\s+(?:la\s+)?Seguridad\s+y\s+Salud\s+en\s+el\s+Trabajo":
                "Vigía de Seguridad y Salud en el Trabajo",
            r"COMIT[EÉ]\s+PARITARIO\s+DE\s+(?:LA\s+)?SEGURIDAD\s+Y\s+SALUD\s+EN\s+EL\s+TRABAJO":
                "Vigía de Seguridad y Salud en el Trabajo",
            # Versión corta "Comité Paritario" solo
            r"\bComit[eé]\s+Paritario\b": "Vigía",
            r"\bCOMIT[EÉ]\s+PARITARIO\b": "Vigía",

            # ============================================================
            # 4) "COPASST" como sigla → siempre "Vigía SST"
            # Nota: IGNORECASE en el sub hace que el primer patrón capture
            # cualquier capitalización de COPASST; los dos siguientes son
            # fallback por si acaso se cambia la lógica.
            # ============================================================
            r"\bCOPASST\b": "Vigía SST",
            r"\bCopasst\b": "Vigía SST",
            r"\bcopasst\b": "Vigía SST",
        }
    else:
        return 0  # Plantillas ya están en COPASST, no hay que cambiar nada

    total = 0
    for archivo in output_base.rglob("*.docx"):
        if archivo.name.startswith("~$"):
            continue
        try:
            doc = _abrir_docx(archivo)

            cambios = 0
            # Procesar párrafos, tablas, encabezados y pies
            for contenedor in _iterar_contenedores_docx(doc):
                for p in contenedor:
                    for run in p.runs:
                        texto_original = run.text
                        texto_nuevo = texto_original
                        for patron, reemplazo in reemplazos.items():
                            texto_nuevo = _re.sub(patron, reemplazo, texto_nuevo, flags=_re.IGNORECASE)
                        if texto_nuevo != texto_original:
                            run.text = texto_nuevo
                            cambios += 1

            if cambios > 0:
                doc.save(str(archivo))
                total += cambios

        except Exception as e:
            console.print(f"   [red]Error adaptando COPASST/Vigía en {archivo.name}: {e}[/red]")

    return total


def _iterar_contenedores_docx(doc):
    """Genera todos los contenedores de párrafos de un Document."""
    yield doc.paragraphs
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                yield celda.paragraphs
    for section in doc.sections:
        for element in [section.header, section.footer]:
            yield element.paragraphs
            for tabla in element.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        yield celda.paragraphs


def _renombrar_carpetas_copasst(output_base: Path, es_copasst: bool) -> int:
    """Renombra carpetas que contengan 'COPASST' en su nombre.

    Returns:
        Número de carpetas renombradas
    """
    if es_copasst:
        return 0

    renombradas = 0
    # Procesar de más profundo a más superficial para evitar conflictos
    carpetas = sorted(
        [d for d in output_base.rglob("*") if d.is_dir() and "COPASST" in d.name],
        key=lambda p: len(p.parts),
        reverse=True,
    )

    for carpeta in carpetas:
        nuevo_nombre = carpeta.name.replace("COPASST", "VIGÍA DE SST")
        nueva_ruta = carpeta.parent / nuevo_nombre
        if not nueva_ruta.exists():
            try:
                carpeta.rename(nueva_ruta)
                renombradas += 1
            except Exception as e:
                console.print(f"   [red]Error renombrando carpeta {carpeta.name}: {e}[/red]")

    return renombradas


def _renombrar_archivos_copasst(output_base: Path, es_copasst: bool) -> int:
    """Renombra archivos que contengan 'COPASST' en su nombre.

    Returns:
        Número de archivos renombrados
    """
    if es_copasst:
        return 0

    renombrados = 0
    for archivo in output_base.rglob("*"):
        if not archivo.is_file() or "COPASST" not in archivo.name:
            continue
        if archivo.name.startswith("~$"):
            continue

        nuevo_nombre = archivo.name.replace("COPASST", "VIGÍA DE SST")
        nueva_ruta = archivo.parent / nuevo_nombre
        if not nueva_ruta.exists():
            try:
                archivo.rename(nueva_ruta)
                renombrados += 1
            except Exception as e:
                console.print(f"   [red]Error renombrando {archivo.name}: {e}[/red]")

    return renombrados


def _actualizar_normograma_carpeta(output_base: Path, cliente: dict):
    """Actualiza normograma con Claude API (sugerencias adicionales)."""
    normas_dir = output_base / NORMOGRAMA_CARPETA
    if not normas_dir.exists():
        console.print("   [dim]Carpeta NORMAS no encontrada[/dim]")
        return

    for archivo in normas_dir.rglob("*.docx"):
        if archivo.name.startswith("~$"):
            continue
        try:
            texto = extraer_texto_docx(archivo)
            if len(texto) > 100:
                texto_actualizado = actualizar_normograma(texto, cliente)
                if texto_actualizado != texto:
                    doc = _abrir_docx(archivo)
                    parrafos_con_texto = [
                        (i, p) for i, p in enumerate(doc.paragraphs) if p.text.strip()
                    ]
                    lineas_nuevas = [l for l in texto_actualizado.split("\n") if l.strip()]
                    for i, (idx, paragraph) in enumerate(parrafos_con_texto):
                        if i < len(lineas_nuevas):
                            if paragraph.runs:
                                paragraph.runs[0].text = lineas_nuevas[i]
                                for run in paragraph.runs[1:]:
                                    run.text = ""
                    doc.save(str(archivo))
                    console.print(f"   [green]Normograma actualizado con Claude:[/green] {archivo.name}")
        except Exception as e:
            console.print(f"   [red]Error en normograma {archivo.name}: {e}[/red]")
