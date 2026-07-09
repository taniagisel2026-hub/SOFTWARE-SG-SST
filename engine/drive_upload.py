"""Logica de negocio para subir un SG-SST generado a Google Drive.

Extraido de cli.py (_ejecutar_subida_drive) para que tanto el CLI como la
API REST (api.py) llamen a la misma funcion sin duplicar logica.
"""
from __future__ import annotations

from pathlib import Path

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn

from engine.listado_maestro import actualizar_matriz_gestion_documental

console = Console(highlight=False)


def ejecutar_subida_drive(
    cliente: dict,
    output_dir: Path,
    cliente_id: str,
    actualizar_listado: bool,
) -> dict:
    """Sube un SG-SST generado a Google Drive y actualiza la matriz de gestion documental.

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
