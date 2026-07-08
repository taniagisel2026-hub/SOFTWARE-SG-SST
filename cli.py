"""CLI principal del generador de SG-SST."""
from __future__ import annotations

import functools
import os
import sys
import logging
import tempfile
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn

__version__ = "8.5.0"

# Constantes de estimacion de tokens
CHARS_POR_TOKEN = 4          # ~4 caracteres por token (aproximacion para espanol)
TOKEN_OVERHEAD_PROMPT = 1000  # tokens fijos por prompt del sistema + contexto del cliente

# Imports del engine al top-level con proteccion.
# Si falta algun modulo, el error se muestra claro al usuario.
try:
    from engine.config import (
        PLANTILLA_PATH, OUTPUT_BASE_PATH, CACHE_DIR,
        ANTHROPIC_API_KEY, CACHE_EXPIRY_DAYS, DOC_TYPES,
        CLAUDE_PRECIO_INPUT_MTOK, CLAUDE_PRECIO_OUTPUT_MTOK,
    )
    from engine.cliente import (
        cargar_cliente, generar_contexto, obtener_imagenes,
        nivel_humanizacion, clasificar_documento,
    )
    from engine.docx_engine import (
        procesar_docx, extraer_secciones_contenido,
        listar_marcadores_en_docx,
    )
    from engine.xlsx_engine import listar_marcadores_en_xlsx
    from engine.claude_service import (
        humanizar_secciones, estadisticas_cache,
        limpiar_cache as _limpiar_cache_engine,
        obtener_uso_sesion,
    )
    from engine.generador import generar_sgsst
    from engine.listado_maestro import actualizar_matriz_gestion_documental
    from engine.normativa import listar_correcciones, CORRECCIONES
    from engine.colores import listar_paletas, PALETAS
    from engine.utils import encontrar_ruta_logo
except ImportError as _import_err:
    print(f"[ERROR] No se pudo importar el motor del SG-SST: {_import_err}")
    print("Verifica que estas ejecutando desde la carpeta del proyecto")
    print("y que todas las dependencias estan instaladas (pip install -r requirements.txt)")
    sys.exit(1)


os.environ["PYTHONIOENCODING"] = "utf-8"
console = Console(highlight=False)

logger = logging.getLogger("sgsst")
logger.propagate = False  # evitar logs duplicados si otro modulo configura logging

# Claves minimas que debe tener un YAML de cliente para no reventar
_YAML_CLAVES_REQUERIDAS = ("razon_social", "nit")


# ---------------------------------------------------------------------------
# Utilidades internas
# ---------------------------------------------------------------------------

def _setup_logging(verbose: bool = False) -> None:
    """Configura logging global segun nivel de verbosidad."""
    level = logging.DEBUG if verbose else logging.WARNING
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(
        "[%(levelname)s] %(name)s: %(message)s"
    ))
    logger.setLevel(level)
    if not logger.handlers:
        logger.addHandler(handler)


def _fatal(msg: str, exc: Exception, verbose: bool = False) -> None:
    """Muestra error al usuario, traza si verbose, y termina con exit(1)."""
    console.print(f"\n[red bold]Error:[/red bold] {msg}")
    console.print(f"  [red]{exc}[/red]")
    if verbose:
        import traceback
        console.print(f"\n[dim]{traceback.format_exc()}[/dim]")
    logger.error(msg, exc_info=True)
    sys.exit(1)


def _get_verbose() -> bool:
    """Obtiene el flag verbose del contexto de Click de forma segura."""
    ctx = click.get_current_context(silent=True)
    if ctx and ctx.obj:
        return ctx.obj.get("verbose", False)
    return False


def cli_error_handler(f):
    """Decorador que captura excepciones en comandos CLI y llama a _fatal."""
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        verbose = _get_verbose()
        try:
            return f(*args, **kwargs)
        except SystemExit:
            raise
        except FileNotFoundError as e:
            _fatal("Archivo no encontrado", e, verbose)
        except Exception as e:
            ctx = click.get_current_context(silent=True)
            cmd = ctx.info_name if ctx else "desconocido"
            _fatal(f"Error en comando '{cmd}'", e, verbose)
    return wrapper


def _validar_yaml(config_path: Path) -> dict:
    """Carga y valida el YAML del cliente. Retorna el dict si es valido.

    Raises:
        click.BadParameter: si faltan claves obligatorias o el formato es invalido.
    """
    cliente = cargar_cliente(config_path)

    if not isinstance(cliente, dict):
        raise click.BadParameter(
            f"El archivo {config_path.name} no contiene un YAML valido (se esperaba un diccionario)."
        )

    faltantes = [k for k in _YAML_CLAVES_REQUERIDAS if not cliente.get(k)]
    if faltantes:
        raise click.BadParameter(
            f"Faltan claves obligatorias en {config_path.name}: {', '.join(faltantes)}. "
            f"Revisa el archivo de configuracion."
        )

    return cliente


def _mostrar_estimacion_costo(
    secciones_input: list[dict],
    secciones_output: list[dict] | None = None,
) -> None:
    """Muestra estimacion de tokens y costo.

    Args:
        secciones_input: secciones originales (para calcular tokens de entrada).
        secciones_output: secciones humanizadas (para calcular tokens de salida).
                          Si es None, se estima igual al input.
    """
    chars_input = sum(len(s.get("texto_completo", "")) for s in secciones_input)
    tokens_input = (chars_input // CHARS_POR_TOKEN) + TOKEN_OVERHEAD_PROMPT

    if secciones_output:
        chars_output = sum(
            len(s.get("texto_humanizado", s.get("texto_completo", "")))
            for s in secciones_output
        )
        tokens_output = chars_output // CHARS_POR_TOKEN
    else:
        tokens_output = chars_input // CHARS_POR_TOKEN

    costo_input = tokens_input * CLAUDE_PRECIO_INPUT_MTOK / 1_000_000
    costo_output = tokens_output * CLAUDE_PRECIO_OUTPUT_MTOK / 1_000_000
    costo_total = costo_input + costo_output

    from engine.config import CLAUDE_MODEL
    modelo_nombre = CLAUDE_MODEL.split("-")[1].capitalize() if "-" in CLAUDE_MODEL else CLAUDE_MODEL
    console.print(f"[bold]Estimacion de costos Claude API ({modelo_nombre}):[/bold]")
    console.print(f"  Caracteres de contenido: {chars_input:,}")
    console.print(f"  Tokens input estimados:  ~{tokens_input:,} (${costo_input:.4f})")
    console.print(f"  Tokens output estimados: ~{tokens_output:,} (${costo_output:.4f})")
    console.print(f"  [bold]Costo total estimado:    ~${costo_total:.4f} USD[/bold]")
    console.print(
        f"  [dim](Precios: ${CLAUDE_PRECIO_INPUT_MTOK}/MTok in, "
        f"${CLAUDE_PRECIO_OUTPUT_MTOK}/MTok out - configurables en .env)[/dim]\n"
    )


# ---------------------------------------------------------------------------
# Grupo principal
# ---------------------------------------------------------------------------

@click.group()
@click.version_option(version=__version__, prog_name="SG-SST Generator")
@click.option("--verbose", "-v", is_flag=True, default=False,
              help="Modo verbose: muestra trazas de error completas")
@click.pass_context
def cli(ctx: click.Context, verbose: bool) -> None:
    """Generador de Sistemas de Gestion de Seguridad y Salud en el Trabajo.

    Basado en Decreto 1072 de 2015 y Resolucion 0312 de 2019.
    Desarrollado por Medina SST Empresarial S.A.S
    """
    ctx.ensure_object(dict)
    ctx.obj["verbose"] = verbose
    _setup_logging(verbose)


# ---------------------------------------------------------------------------
# generar
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("config", type=click.Path(exists=True))
@click.option("--output", "-o", type=click.Path(), default=None,
              help="Directorio de salida (por defecto: clientes/<razon_social>)")
@click.option("--humanizar/--no-humanizar", default=True,
              help="Humanizar contenido con Claude API")
@click.option("--solo-marcadores", is_flag=True, default=False,
              help="Solo reemplazar marcadores, sin humanizacion")
@click.option("--listado-maestro/--no-listado-maestro", default=True,
              help="Generar listado maestro de documentos")
@click.option("--estricto", is_flag=True, default=False,
              help="Modo estricto: falla si hay marcadores sin dato en el YAML")
@click.pass_context
@cli_error_handler
def generar(
    ctx: click.Context,
    config: str,
    output: str | None,
    humanizar: bool,
    solo_marcadores: bool,
    listado_maestro: bool,
    estricto: bool,
) -> None:
    """Genera el SG-SST completo para un cliente.

    CONFIG: Ruta al archivo YAML de configuracion del cliente.

    Ejemplos:

        python cli.py generar clientes_config/mi_empresa.yaml

        python cli.py generar config.yaml --no-humanizar

        python cli.py generar config.yaml --estricto
    """
    config_path = Path(config)
    _validar_yaml(config_path)  # falla rapido si el YAML esta mal

    resultado = generar_sgsst(
        config_path=config,
        output_dir=output,
        humanizar=humanizar,
        solo_marcadores=solo_marcadores,
        estricto=estricto,
    )

    # Mostrar resumen de errores si hubo problemas internos
    errores = resultado.get("errores", [])
    if errores:
        console.print(f"\n[yellow]Advertencia: {len(errores)} archivo(s) con errores durante la generacion:[/yellow]")
        for nombre, error in errores[:5]:
            console.print(f"  [red]X[/red] {nombre}: {error}")
        if len(errores) > 5:
            console.print(f"  ... y {len(errores) - 5} mas (ver reporte de auditoria)")

    # NOTA: Ya no generamos un LISTADO_MAESTRO.xlsx suelto. El listado maestro
    # es la matriz SST-MTZ-005 (I. PLANEAR/2.5.1 ARCHIVO Y RETENCION DOCUMENTAL)
    # que se actualiza con links de Drive tras la subida con `subir-drive`.
    if listado_maestro and resultado["procesados"] > 0:
        console.print(
            "[cyan]El listado maestro es la matriz SST-MTZ-005 y se actualizara "
            "con los links al subir a Drive ('subir-drive').[/cyan]"
        )


# ---------------------------------------------------------------------------
# verificar
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("config", type=click.Path(exists=True))
@click.pass_context
@cli_error_handler
def verificar(ctx: click.Context, config: str) -> None:
    """Verifica la configuracion del cliente sin generar documentos.

    Muestra los marcadores que se van a reemplazar y valida que
    los datos esten completos.
    """
    config_path = Path(config)
    cliente = _validar_yaml(config_path)
    contexto = generar_contexto(cliente)
    imagenes = obtener_imagenes(cliente, config_path.parent)

    # Tabla de marcadores base (sin variantes duplicadas)
    table = Table(title=f"Marcadores para: {cliente.get('razon_social', 'N/A')}")
    table.add_column("Marcador", style="cyan")
    table.add_column("Valor", style="green")
    table.add_column("Estado", style="bold")

    vistos: set[str] = set()
    for marcador in sorted(contexto.keys()):
        upper_key = marcador.upper()
        if upper_key in vistos:
            continue
        vistos.add(upper_key)
        valor = contexto[marcador]
        estado = "[green]OK[/green]" if valor and str(valor).strip() else "[red]VACIO[/red]"
        display = str(valor)[:60] + ("..." if len(str(valor)) > 60 else "")
        table.add_row(marcador, display, estado)

    console.print(table)

    # Imagenes
    console.print(f"\n[bold]Imagenes encontradas:[/bold] {len(imagenes)}")
    for key, path in imagenes.items():
        console.print(f"  {key}: {path}")

    # Marcadores vacios
    marcadores_vacios = [m for m, v in contexto.items() if not v or not str(v).strip()]
    marcadores_vacios_unicos = sorted(set(m.upper() for m in marcadores_vacios))
    if marcadores_vacios_unicos:
        console.print(f"\n[yellow]Advertencia: {len(marcadores_vacios_unicos)} marcadores vacios[/yellow]")
        for m in marcadores_vacios_unicos:
            console.print(f"  {m}")

    # Opciones de generacion
    console.print(f"\n[bold]Opciones:[/bold]")
    console.print(f"  PESV: {'Si' if cliente.get('generar_pesv') else 'No'}")
    console.print(f"  Normograma actualizado: {'Si' if cliente.get('generar_normograma_actualizado') else 'No'}")
    console.print(f"  Sedes: {len(cliente.get('sedes', []))}")


# ---------------------------------------------------------------------------
# escanear
# ---------------------------------------------------------------------------

@cli.command()
@click.option("--tipo", type=click.Choice(["todos", "docx", "xlsx"]), default="todos",
              help="Filtrar por tipo de archivo")
@click.pass_context
@cli_error_handler
def escanear(ctx: click.Context, tipo: str) -> None:
    """Escanea la plantilla y muestra todos los marcadores encontrados."""
    console.print(f"[bold]Escaneando plantilla:[/bold] {PLANTILLA_PATH}")
    if tipo != "todos":
        console.print(f"[dim]Filtro: solo archivos .{tipo}[/dim]")
    console.print()

    todos_marcadores: set[str] = set()
    archivos_con_marcadores: dict[str, list[str]] = {}

    for archivo in sorted(PLANTILLA_PATH.rglob("*")):
        if archivo.name.startswith("~$"):
            continue

        ext = archivo.suffix.lower()
        marcadores: list[str] = []
        try:
            if ext == ".docx" and tipo in ("todos", "docx"):
                marcadores = listar_marcadores_en_docx(archivo)
            elif ext == ".xlsx" and tipo in ("todos", "xlsx"):
                marcadores = listar_marcadores_en_xlsx(archivo)
        except Exception as e:
            console.print(f"  [red]Error en {archivo.name}: {e}[/red]")
            logger.debug("Error escaneando %s", archivo, exc_info=True)
            continue

        if marcadores:
            ruta_rel = archivo.relative_to(PLANTILLA_PATH)
            archivos_con_marcadores[str(ruta_rel)] = marcadores
            todos_marcadores.update(marcadores)

    # Resumen
    table = Table(title="Marcadores encontrados en la plantilla")
    table.add_column("Marcador", style="cyan")
    table.add_column("Apariciones", style="green", justify="right")

    conteo: dict[str, int] = {}
    for marcadores in archivos_con_marcadores.values():
        for m in marcadores:
            conteo[m] = conteo.get(m, 0) + 1

    for marcador in sorted(conteo.keys()):
        table.add_row(marcador, str(conteo[marcador]))

    console.print(table)
    console.print(
        f"\n[bold]Total:[/bold] {len(todos_marcadores)} marcadores unicos "
        f"en {len(archivos_con_marcadores)} archivos"
    )


# ---------------------------------------------------------------------------
# cache / limpiar-cache
# ---------------------------------------------------------------------------

@cli.command()
@click.pass_context
@cli_error_handler
def cache(ctx: click.Context) -> None:
    """Muestra estadisticas de la cache de Claude API."""
    stats = estadisticas_cache()
    console.print("[bold]Cache de Claude API:[/bold]")
    console.print(f"  Entradas: {stats['entradas']}")
    console.print(f"  Tamano: {stats['tamaño_bytes'] / 1024:.1f} KB")
    console.print(f"  Directorio: {stats['directorio']}")


@cli.command(name="limpiar-cache")
@click.pass_context
@cli_error_handler
def limpiar_cache(ctx: click.Context) -> None:
    """Limpia la cache de respuestas de Claude API."""
    _limpiar_cache_engine()
    console.print("[green]Cache limpiada exitosamente.[/green]")


# ---------------------------------------------------------------------------
# probar-doc
# ---------------------------------------------------------------------------

@cli.command(name="probar-doc")
@click.argument("documento", type=click.Path(exists=True))
@click.argument("config", type=click.Path(exists=True))
@click.option("--output", "-o", type=click.Path(), default=None,
              help="Ruta de salida del documento procesado")
@click.option("--nivel", type=click.Choice(["completa", "media", "ninguna"]),
              default=None, help="Forzar nivel de humanizacion")
@click.pass_context
@cli_error_handler
def probar_doc(
    ctx: click.Context,
    documento: str,
    config: str,
    output: str | None,
    nivel: str | None,
) -> None:
    """Procesa UN SOLO documento con humanizacion. Ideal para pruebas.

    DOCUMENTO: Ruta al .docx de la plantilla a procesar.
    CONFIG: Ruta al YAML de configuracion del cliente.

    Ejemplos:

        python cli.py probar-doc "Plantilla SG-SST 2026/I. PLANEAR/.../SST-POL-001.docx" config.yaml

        python cli.py probar-doc documento.docx config.yaml --nivel completa -o salida.docx
    """
    doc_path = Path(documento)
    config_path = Path(config)
    cliente = _validar_yaml(config_path)
    contexto = generar_contexto(cliente)
    imagenes = obtener_imagenes(cliente, config_path.parent)

    # Inyectar imagenes al contexto
    for key_img, ruta_img in imagenes.items():
        contexto[key_img] = ruta_img
        contexto[key_img.lower()] = ruta_img

    # Determinar ruta de salida (fallback a temp si no hay permisos de escritura)
    if output:
        output_path = Path(output)
    else:
        output_candidate = doc_path.parent / f"_HUMANIZADO_{doc_path.name}"
        try:
            output_candidate.parent.mkdir(parents=True, exist_ok=True)
            output_path = output_candidate
        except OSError:
            output_path = Path(tempfile.gettempdir()) / f"_HUMANIZADO_{doc_path.name}"
            console.print(f"[yellow]Sin permisos en {doc_path.parent}, usando: {output_path}[/yellow]")

    console.print(f"\n[bold]Procesando:[/bold] {doc_path.name}")
    console.print(f"[bold]Cliente:[/bold] {cliente.get('razon_social', 'N/A')}")

    # Determinar nivel
    if nivel is None:
        nivel = nivel_humanizacion(doc_path.name)
    console.print(f"[bold]Nivel humanizacion:[/bold] {nivel}")

    # Extraer secciones
    secciones = extraer_secciones_contenido(doc_path, contexto)
    console.print(f"[bold]Secciones de contenido:[/bold] {len(secciones)}")
    total_parrafos = sum(len(s["parrafos"]) for s in secciones)
    console.print(f"[bold]Parrafos humanizables:[/bold] {total_parrafos}")

    # Preview de secciones
    for sec in secciones[:3]:
        preview_text = sec["texto_completo"][:100].replace("\n", " ")
        console.print(f"  Seccion {sec['id']}: [dim]{preview_text}...[/dim]")
    if len(secciones) > 3:
        console.print(f"  ... y {len(secciones) - 3} secciones mas")

    # Humanizar
    secciones_humanizadas = None
    if nivel != "ninguna" and secciones:
        console.print("\n[yellow]Enviando a Claude API...[/yellow]")
        tipo = clasificar_documento(doc_path.name)
        tipo_nombre = DOC_TYPES.get(tipo, {}).get("nombre", "documento")

        # Incluir nombres de carpetas padre (hasta 3 niveles) en el contexto,
        # para que Claude identifique si el documento pertenece a CCL, COPASST,
        # PESV, etc. Ej: "1.1.8 COMITE DE CONVIVENCIA LABORAL/3. PROCEDIMIENTOS/SST-PRO-003.docx"
        padres = list(doc_path.parents)
        ruta_contexto = doc_path.name
        for p in padres[:3]:
            if p.name and p.name not in ("", "/", "\\"):
                ruta_contexto = f"{p.name}/{ruta_contexto}"
            else:
                break

        secciones_humanizadas = humanizar_secciones(
            secciones=secciones,
            cliente=cliente,
            nivel=nivel,
            tipo_documento=tipo_nombre,
            nombre_archivo=ruta_contexto,
        )

        # Comparacion
        console.print("\n[bold]--- Comparacion (primeras 2 secciones) ---[/bold]")
        for sec in (secciones_humanizadas or [])[:2]:
            console.print(f"\n[red]ORIGINAL (seccion {sec['id']}):[/red]")
            console.print(f"  {sec['texto_completo'][:200]}...")
            console.print(f"[green]HUMANIZADO:[/green]")
            humanizado = sec.get("texto_humanizado", "")
            console.print(f"  {humanizado[:200]}...")

    # Procesar con temp_dir limpio (se auto-limpia al salir del with)
    with tempfile.TemporaryDirectory(prefix="sgsst_test_") as temp_dir:
        logo_path = encontrar_ruta_logo(
            contexto, [str(doc_path.parent), temp_dir], temp_dir
        )

        console.print(f"\n[yellow]Guardando en:[/yellow] {output_path}")
        procesar_docx(
            plantilla_path=doc_path,
            output_path=output_path,
            contexto=contexto,
            temp_dir=temp_dir,
            logo_path=logo_path,
            secciones_humanizadas=secciones_humanizadas,
        )

    console.print("[bold green]Documento procesado exitosamente.[/bold green]\n")

    # Estimacion de costo (usa output real si hay humanizacion)
    if secciones_humanizadas:
        _mostrar_estimacion_costo(secciones, secciones_humanizadas)
    elif secciones:
        _mostrar_estimacion_costo(secciones)


# ---------------------------------------------------------------------------
# preview
# ---------------------------------------------------------------------------

@cli.command(name="preview")
@click.argument("documento", type=click.Path(exists=True))
@click.argument("config", type=click.Path(exists=True))
@click.pass_context
@cli_error_handler
def preview(ctx: click.Context, documento: str, config: str) -> None:
    """Muestra que secciones se humanizarian sin llamar a Claude API.

    Util para verificar que el sistema identifica correctamente
    el contenido humanizable antes de gastar tokens.
    """
    doc_path = Path(documento)
    config_path = Path(config)
    cliente = _validar_yaml(config_path)
    contexto = generar_contexto(cliente)
    nivel = nivel_humanizacion(doc_path.name)

    console.print(f"\n[bold]Documento:[/bold] {doc_path.name}")
    console.print(f"[bold]Nivel:[/bold] {nivel}")

    secciones = extraer_secciones_contenido(doc_path, contexto)

    if not secciones:
        console.print("[dim]Sin contenido humanizable[/dim]")
        return

    console.print(f"[bold]Secciones:[/bold] {len(secciones)}")
    console.print("[dim](Texto sin reemplazar marcadores - muestra el contenido tal como lo recibira Claude)[/dim]\n")

    for sec in secciones:
        console.print(
            f"[cyan]--- Seccion {sec['id']} "
            f"({len(sec['parrafos'])} parrafos, "
            f"{len(sec['texto_completo'])} chars) ---[/cyan]"
        )
        console.print(f"{sec['texto_completo'][:300]}")
        if len(sec["texto_completo"]) > 300:
            console.print("[dim]...[/dim]")
        console.print()

    _mostrar_estimacion_costo(secciones)


# ---------------------------------------------------------------------------
# subir-drive  (logica pesada en _ejecutar_subida_drive)
# ---------------------------------------------------------------------------

def _ejecutar_subida_drive(
    cliente: dict,
    output_dir: Path,
    cliente_id: str,
    actualizar_listado: bool,
) -> dict:
    """Logica de negocio para subir a Drive, separada del CLI.

    Returns:
        dict con claves: exitosos, errores, total, link
    """
    try:
        from engine.drive_service import (
            autenticar_drive, subir_sgsst_completo,
            obtener_link_carpeta, subir_archivo,
        )
    except ImportError as e:
        raise RuntimeError(
            "No se pudo cargar el modulo de Google Drive. "
            "Verifica que estan instaladas las dependencias: "
            "pip install google-auth google-auth-oauthlib google-api-python-client"
        ) from e

    razon_social = cliente.get("razon_social", "CLIENTE")

    # Autenticar
    console.print("\n[yellow]Autenticando con Google Drive...[/yellow]")
    service = autenticar_drive(cliente_id)
    console.print("[green]Autenticacion exitosa[/green]")

    # Subir con barra de progreso
    console.print("\n[yellow]Subiendo documentos a Drive...[/yellow]")

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        console=console,
    ) as progress:
        task = progress.add_task("Subiendo...", total=1.0)

        def progreso_cb(nombre: str, pct: float) -> None:
            progress.update(
                task, completed=min(pct, 1.0),
                description=f"[dim]{nombre[:45]}[/dim]",
            )

        resultado = subir_sgsst_completo(
            service=service,
            output_dir=output_dir,
            razon_social=razon_social,
            año=cliente.get("año"),
            progreso_callback=progreso_cb,
        )

    # Resultados
    console.print(f"\n[bold green]=== SUBIDA COMPLETADA ===[/bold green]")
    console.print(f"   Archivos subidos: [green]{resultado['exitosos']}[/green] / {resultado['total']}")

    if resultado["errores"]:
        console.print(f"   Errores: [red]{len(resultado['errores'])}[/red]")
        for nombre, error in resultado["errores"][:3]:
            console.print(f"     [red]X[/red] {nombre}: {error}")

    # Link a carpeta raiz
    link = obtener_link_carpeta(service, resultado["raiz_id"])
    console.print(f"\n   [bold]Carpeta en Drive:[/bold] {link}")

    # Actualizar la matriz SST-MTZ-005 con los links de Drive y re-subirla
    if actualizar_listado:
        console.print(
            "\n[yellow]Actualizando matriz SST-MTZ-005 con links de Drive...[/yellow]"
        )
        from engine.listado_maestro import RUTA_MATRIZ_GESTION  # ruta relativa
        ruta_matriz = actualizar_matriz_gestion_documental(
            output_dir,
            drive_links=resultado["archivos_subidos"],
            generar_pesv=cliente.get("generar_pesv", False),
        )
        if ruta_matriz is None:
            console.print(
                "[yellow]No se encontro la matriz SST-MTZ-005 en el output; "
                "se omite la actualizacion del listado maestro.[/yellow]"
            )
        else:
            console.print(f"[green]Matriz actualizada localmente:[/green] {ruta_matriz}")
            # Re-subir la matriz a Drive sobre el mismo archivo ya existente
            try:
                matriz_rel = str(RUTA_MATRIZ_GESTION)
                info_drive = resultado["archivos_subidos"].get(matriz_rel)
                if info_drive and info_drive.get("id"):
                    from googleapiclient.http import MediaFileUpload
                    media = MediaFileUpload(
                        str(ruta_matriz),
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        resumable=True,
                    )
                    service.files().update(
                        fileId=info_drive["id"], media_body=media
                    ).execute()
                    console.print("[green]Matriz re-subida a Drive con los links actualizados.[/green]")
                else:
                    # Fallback: subir como archivo nuevo
                    subir_archivo(service, ruta_matriz, resultado["raiz_id"])
                    console.print("[green]Matriz subida a Drive (fallback).[/green]")
            except Exception as e:
                console.print(f"[red]Error re-subiendo la matriz: {e}[/red]")

    console.print()
    return {
        "exitosos": resultado["exitosos"],
        "errores": resultado["errores"],
        "total": resultado["total"],
        "link": link,
    }


@cli.command(name="subir-drive")
@click.argument("config", type=click.Path(exists=True))
@click.option("--directorio", "-d", type=click.Path(exists=True), default=None,
              help="Directorio del SG-SST generado a subir")
@click.option("--cliente-id", type=str, default=None,
              help="ID unico del cliente (para token OAuth)")
@click.option("--actualizar-listado/--no-actualizar-listado", default=True,
              help="Regenerar listado maestro con links de Drive")
@click.pass_context
@cli_error_handler
def subir_drive(
    ctx: click.Context,
    config: str,
    directorio: str | None,
    cliente_id: str | None,
    actualizar_listado: bool,
) -> None:
    """Sube el SG-SST generado a Google Drive del cliente.

    Replica la estructura de carpetas PHVA en Drive y genera
    el listado maestro con hipervinculos a cada documento.

    La primera vez que se ejecuta para un cliente, abre el navegador
    para autorizar el acceso a su Google Drive.

    Ejemplos:

        python cli.py subir-drive config.yaml

        python cli.py subir-drive config.yaml -d clientes/MI_EMPRESA
    """
    config_path = Path(config)
    cliente = _validar_yaml(config_path)
    razon_social = cliente.get("razon_social", "CLIENTE")

    # Determinar directorio de origen
    if directorio:
        output_dir = Path(directorio)
    else:
        nombre_carpeta = razon_social.replace(" ", "_").replace(".", "")
        output_dir = OUTPUT_BASE_PATH / nombre_carpeta

    if not output_dir.exists():
        console.print(f"[red]No se encontro el directorio: {output_dir}[/red]")
        console.print("Primero genera el SG-SST con: python cli.py generar config.yaml")
        sys.exit(1)

    # ID del cliente para token
    if not cliente_id:
        nit = str(cliente.get("nit", "")).replace(".", "").replace("-", "")
        cliente_id = nit or razon_social.replace(" ", "_")

    console.print(f"\n[bold blue]=== SUBIDA A GOOGLE DRIVE ===[/bold blue]\n")
    console.print(f"[bold]Empresa:[/bold] {razon_social}")
    console.print(f"[bold]Origen:[/bold] {output_dir}")
    console.print(f"[bold]Cliente ID:[/bold] {cliente_id}")

    resumen = _ejecutar_subida_drive(cliente, output_dir, cliente_id, actualizar_listado)

    # Usar el retorno para resumen final
    if resumen["errores"]:
        console.print(
            f"[yellow]Subida completada con {len(resumen['errores'])} error(es). "
            f"Revisar los archivos que no se subieron.[/yellow]"
        )


# ---------------------------------------------------------------------------
# auth-drive
# ---------------------------------------------------------------------------

@cli.command(name="auth-drive")
@click.option("--cliente-id", type=str, required=True,
              help="ID unico del cliente")
@click.pass_context
@cli_error_handler
def auth_drive(ctx: click.Context, cliente_id: str) -> None:
    """Autentica con Google Drive para un cliente (solo autenticacion).

    Abre el navegador para autorizar acceso. Util para pre-autorizar
    antes de ejecutar subir-drive.
    """
    try:
        from engine.drive_service import autenticar_drive
    except ImportError as e:
        raise RuntimeError(
            "No se pudo cargar el modulo de Google Drive. "
            "Verifica: pip install google-auth google-auth-oauthlib google-api-python-client"
        ) from e

    console.print(f"\n[bold]Autenticando cliente:[/bold] {cliente_id}")
    service = autenticar_drive(cliente_id)
    about = service.about().get(fields="user").execute()
    email = about.get("user", {}).get("emailAddress", "N/A")
    console.print(f"[green]Autenticacion exitosa[/green]")
    console.print(f"[bold]Cuenta:[/bold] {email}")
    console.print(f"[dim]Token guardado para futuras sesiones[/dim]\n")


# ---------------------------------------------------------------------------
# normas
# ---------------------------------------------------------------------------

@cli.command()
@click.pass_context
@cli_error_handler
def normas(ctx: click.Context) -> None:
    """Muestra las correcciones legales configuradas en el sistema.

    Lista todas las normas derogadas/modificadas que se corrigen
    automaticamente en los documentos generados del SG-SST.
    """
    correcciones = listar_correcciones()

    table = Table(title="Correcciones Legales Configuradas")
    table.add_column("#", style="dim", justify="right")
    table.add_column("Norma original (derogada/erronea)", style="red")
    table.add_column("Correccion aplicada", style="green")
    table.add_column("Vigente desde", style="yellow", justify="right")

    for i, c in enumerate(correcciones, 1):
        vigente = str(c["vigente_desde"]) if c["vigente_desde"] else "siempre"
        table.add_row(str(i), c["vieja"], c["nueva"], vigente)

    console.print(f"\n")
    console.print(table)
    console.print(
        f"\n   [bold]{len(correcciones)}[/bold] correcciones configuradas."
    )
    console.print(
        "   Se aplican automaticamente a [bold]todos[/bold] los .docx generados,\n"
        "   segun el campo 'año' del cliente (una correccion 'vigente desde' un año\n"
        "   posterior al del cliente NO se aplica, para no adelantar cambios legales\n"
        "   que en ese momento aun no existian).\n"
    )
    console.print(
        "   [dim]Para agregar una correccion, edita CORRECCIONES en engine/normativa.py[/dim]\n"
    )


# ---------------------------------------------------------------------------
# colores
# ---------------------------------------------------------------------------

@cli.command()
@click.argument("paleta", required=False, default=None)
@click.argument("archivo", required=False, default=None, type=click.Path(exists=True))
@click.pass_context
@cli_error_handler
def colores(ctx: click.Context, paleta: str | None, archivo: str | None) -> None:
    """Muestra las paletas disponibles o prueba una paleta en un archivo.

    Sin argumentos: muestra la tabla de paletas.

    Con argumentos: aplica la paleta a un archivo .docx o .xlsx.

    \b
    Ejemplos:
        python cli.py colores
        python cli.py colores azul "mi archivo.docx"
        python cli.py colores verde "archivo.xlsx"
        python cli.py colores "#1A5276" "archivo.docx"

    Para usarlo en produccion, agrega al YAML del cliente:
        color_corporativo: azul
    """
    from engine.colores import (
        aplicar_colores_docx, aplicar_colores_xlsx,
        resolver_paleta, _construir_mapa_colores,
    )

    # Si dan paleta + archivo: aplicar colores a ese archivo
    if paleta and archivo:
        archivo_path = Path(archivo)
        pal = resolver_paleta(paleta)
        if pal is None:
            console.print(f"[red]Paleta '{paleta}' no reconocida.[/red]")
            console.print("Usa: azul, verde, morado, naranja, gris o un hex (#2E75B6)")
            return

        ext = archivo_path.suffix.lower()
        if ext not in (".docx", ".xlsx"):
            console.print(f"[red]Solo se pueden cambiar colores en .docx y .xlsx[/red]")
            return

        mapa = _construir_mapa_colores(pal)
        console.print(f"\n[bold]Paleta:[/bold] {pal['nombre']}")
        console.print(f"[bold]Archivo:[/bold] {archivo_path.name}")

        if ext == ".docx":
            cambios = aplicar_colores_docx(str(archivo_path), mapa)
        else:
            cambios = aplicar_colores_xlsx(str(archivo_path), mapa)

        if cambios > 0:
            console.print(f"[green]{cambios} cambios de color aplicados[/green]")
        else:
            console.print("[dim]Sin colores vino tinto para cambiar[/dim]")
        console.print(f"Abre el archivo para verificar.\n")
        return

    # Si dan solo paleta sin archivo
    if paleta and not archivo:
        console.print(f"[yellow]Falta el archivo. Uso:[/yellow]")
        console.print(f'  python cli.py colores {paleta} "archivo.docx"\n')
        return

    # Sin argumentos: mostrar paletas
    paletas_list = listar_paletas()

    table = Table(title="Paletas de Colores Corporativos")
    table.add_column("Nombre", style="bold")
    table.add_column("Descripcion", style="cyan")
    table.add_column("Color medio", style="green")
    table.add_column("Uso en YAML", style="dim")

    for p in paletas_list:
        table.add_row(p["nombre"], p["descripcion"], f"#{p['color_medio']}", f'color_corporativo: "{p["nombre"]}"')

    console.print(f"\n")
    console.print(table)
    console.print(f"\n   [bold]{len(paletas_list)}[/bold] paletas predefinidas.")
    console.print("   Tambien puedes usar un color hex directo: [cyan]color_corporativo: \"#2E75B6\"[/cyan]\n")
    console.print('   [bold]Para probar:[/bold] python cli.py colores azul "archivo.docx"\n')


# ---------------------------------------------------------------------------
# info
# ---------------------------------------------------------------------------

@cli.command()
@click.pass_context
@cli_error_handler
def info(ctx: click.Context) -> None:
    """Muestra informacion del sistema y configuracion actual."""
    console.print(f"\n[bold blue]=== SG-SST Generator v{__version__} ===[/bold blue]\n")

    # Configuracion
    table = Table(title="Configuracion del Sistema")
    table.add_column("Parametro", style="cyan")
    table.add_column("Valor", style="green")
    table.add_column("Estado", style="bold")

    plantilla_ok = PLANTILLA_PATH.exists()
    table.add_row(
        "Plantilla",
        str(PLANTILLA_PATH),
        "[green]OK[/green]" if plantilla_ok else "[red]NO ENCONTRADA[/red]",
    )

    table.add_row("Salida", str(OUTPUT_BASE_PATH), "[green]OK[/green]")
    table.add_row("Cache", str(CACHE_DIR), f"{CACHE_EXPIRY_DAYS} dias expiracion")

    # Solo mostrar ultimos 4 caracteres de la API key por seguridad
    api_ok = bool(ANTHROPIC_API_KEY)
    api_display = f"***...{ANTHROPIC_API_KEY[-4:]}" if api_ok else "No configurada"
    table.add_row(
        "Claude API Key",
        api_display,
        "[green]OK[/green]" if api_ok else "[yellow]FALTA[/yellow]",
    )

    table.add_row(
        "Precios Claude",
        f"${CLAUDE_PRECIO_INPUT_MTOK}/MTok in, ${CLAUDE_PRECIO_OUTPUT_MTOK}/MTok out",
        "[dim]configurables en .env[/dim]",
    )

    cred_path = Path("credentials.json")
    table.add_row(
        "Google Drive",
        "credentials.json",
        "[green]OK[/green]" if cred_path.exists() else "[yellow]FALTA[/yellow]",
    )

    console.print(table)

    # Tipos de documento
    console.print(f"\n[bold]Tipos de documento configurados:[/bold] {len(DOC_TYPES)}")
    for code, info_doc in DOC_TYPES.items():
        h_color = {"completa": "green", "media": "yellow", "ninguna": "dim"}.get(
            info_doc["humanizar"], "white"
        )
        console.print(
            f"  {info_doc['prefijo']}xxx  {info_doc['nombre']:<14} "
            f"[{h_color}]{info_doc['humanizar']}[/{h_color}]"
        )

    # Contar archivos en plantilla
    if plantilla_ok:
        archivos = list(PLANTILLA_PATH.rglob("*"))
        files = [a for a in archivos if a.is_file() and not a.name.startswith("~$")]
        docx = len([f for f in files if f.suffix.lower() == ".docx"])
        xlsx = len([f for f in files if f.suffix.lower() == ".xlsx"])
        otros = len(files) - docx - xlsx
        console.print(
            f"\n[bold]Plantilla:[/bold] {len(files)} archivos "
            f"({docx} .docx, {xlsx} .xlsx, {otros} otros)"
        )

    console.print()


if __name__ == "__main__":
    cli()
