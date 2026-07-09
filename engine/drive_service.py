"""Servicio de Google Drive - gestiona carpetas y archivos por cliente."""
import os
import mimetypes
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from .config import BASE_DIR

# Permisos necesarios
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Ruta a las credenciales OAuth (flujo de escritorio, usado por el CLI)
CREDENTIALS_PATH = BASE_DIR / "credentials.json"
# Ruta a las credenciales OAuth "Web application" (flujo web, usado por la API).
# Se crea aparte en Google Cloud Console -- ver iniciar_auth_web().
CREDENTIALS_WEB_PATH = BASE_DIR / "credentials_web.json"
TOKENS_DIR = BASE_DIR / "tokens"

# MIME types para carpetas y archivos de Google
FOLDER_MIME = "application/vnd.google-apps.folder"


def _escape_drive_query(value: str) -> str:
    """Escapa comillas simples en valores para queries de Google Drive API."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


# Mapeo de extensiones a MIME types
MIME_MAP = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
}


def _token_path(cliente_id: str) -> Path:
    """Ruta al token de autenticacion de un cliente."""
    TOKENS_DIR.mkdir(parents=True, exist_ok=True)
    return TOKENS_DIR / f"token_{cliente_id}.json"


def autenticar_drive(cliente_id: str) -> object:
    """Autentica con Google Drive usando OAuth 2.0 por cliente.

    La primera vez abre el navegador para autorizar.
    Despues usa el token guardado.

    Args:
        cliente_id: Identificador unico del cliente (usado para guardar token)

    Returns:
        Servicio de Google Drive autenticado
    """
    creds = None
    token_file = _token_path(cliente_id)

    # Cargar token existente
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), SCOPES)

    # Si no hay token o esta expirado, re-autenticar
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"No se encontro credentials.json en {CREDENTIALS_PATH}\n"
                    "Descargalo desde Google Cloud Console > APIs > Credenciales"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Guardar token
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def iniciar_auth_web(cliente_id: str, redirect_uri: str) -> str:
    """Genera la URL de consentimiento de Google para el flujo OAuth "web" (API REST).

    A diferencia de autenticar_drive() (flujo de escritorio, abre un navegador
    EN LA MISMA MAQUINA donde corre el proceso), este flujo esta pensado para
    correr en un servidor: quien se autentica abre esta URL en SU PROPIO
    navegador, y Google redirige a completar_auth_web() con el codigo.

    Requiere credentials_web.json (Client ID tipo "Web application" en
    Google Cloud Console, con `redirect_uri` registrado como Authorized
    redirect URI) -- ver docs/ARQUITECTURA.md.
    """
    if not CREDENTIALS_WEB_PATH.exists():
        raise FileNotFoundError(
            f"No se encontro {CREDENTIALS_WEB_PATH}. "
            "Crea un Client ID tipo 'Web application' en Google Cloud Console "
            "(distinto del credentials.json de escritorio) y descargalo con ese nombre."
        )
    flow = Flow.from_client_secrets_file(
        str(CREDENTIALS_WEB_PATH), SCOPES, redirect_uri=redirect_uri
    )
    auth_url, _ = flow.authorization_url(
        access_type="offline",  # necesario para recibir refresh_token
        prompt="consent",       # fuerza a que Google lo reenvie incluso en re-autorizaciones
        state=cliente_id,
    )
    return auth_url


def completar_auth_web(cliente_id: str, code: str, redirect_uri: str) -> None:
    """Intercambia el codigo de autorizacion por un token y lo guarda.

    Escribe en el MISMO formato/ruta que autenticar_drive() (tokens/token_<id>.json),
    asi que una vez guardado, autenticar_drive() lo usa sin ningun cambio --
    no le importa si el token se creo por el flujo de escritorio o el web.
    """
    flow = Flow.from_client_secrets_file(
        str(CREDENTIALS_WEB_PATH), SCOPES, redirect_uri=redirect_uri
    )
    flow.fetch_token(code=code)
    creds = flow.credentials

    token_file = _token_path(cliente_id)
    with open(token_file, "w") as f:
        f.write(creds.to_json())


def crear_carpeta(service, nombre: str, parent_id: str = None) -> str:
    """Crea una carpeta en Google Drive y retorna su ID.

    Si la carpeta ya existe en el parent, retorna el ID existente.
    """
    # Buscar si ya existe
    safe_nombre = _escape_drive_query(nombre)
    query = f"name='{safe_nombre}' and mimeType='{FOLDER_MIME}' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    else:
        # Sin parent explícito: restringir a la raíz de Mi Unidad para no
        # encontrar carpetas con el mismo nombre que fueron compartidas o movidas.
        query += " and 'root' in parents"

    results = service.files().list(
        q=query, fields="files(id, name)", pageSize=1
    ).execute()

    if results.get("files"):
        return results["files"][0]["id"]

    # Crear nueva
    metadata = {"name": nombre, "mimeType": FOLDER_MIME}
    if parent_id:
        metadata["parents"] = [parent_id]

    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def _obtener_o_crear_carpeta_ruta(service, ruta_relativa: Path, raiz_id: str,
                                    cache_carpetas: dict) -> str:
    """Obtiene o crea la jerarquia de carpetas para una ruta relativa.

    Args:
        service: Servicio de Drive
        ruta_relativa: Ruta como "I. PLANEAR/1.1.1 RESPONSABLE/1. DESIGNACION"
        raiz_id: ID de la carpeta raiz del cliente en Drive
        cache_carpetas: Dict de {ruta: folder_id} para evitar re-crear

    Returns:
        ID de la carpeta final
    """
    partes = ruta_relativa.parts
    parent_id = raiz_id
    ruta_acumulada = Path()

    for parte in partes:
        ruta_acumulada = ruta_acumulada / parte
        ruta_str = str(ruta_acumulada)

        if ruta_str in cache_carpetas:
            parent_id = cache_carpetas[ruta_str]
        else:
            folder_id = crear_carpeta(service, parte, parent_id)
            cache_carpetas[ruta_str] = folder_id
            parent_id = folder_id

    return parent_id


def subir_archivo(service, archivo_local: Path, parent_id: str) -> dict:
    """Sube un archivo a Google Drive.

    Args:
        archivo_local: Ruta local al archivo
        parent_id: ID de la carpeta destino en Drive

    Returns:
        Dict con id, name, webViewLink del archivo subido
    """
    # Determinar MIME type
    ext = archivo_local.suffix.lower()
    mime_type = MIME_MAP.get(ext, mimetypes.guess_type(str(archivo_local))[0] or "application/octet-stream")

    # Verificar si ya existe (para actualizar en vez de duplicar)
    safe_name = _escape_drive_query(archivo_local.name)
    query = (
        f"name='{safe_name}' and '{parent_id}' in parents and trashed=false"
    )
    existente = service.files().list(q=query, fields="files(id)").execute()

    if existente.get("files"):
        # Actualizar archivo existente
        file_id = existente["files"][0]["id"]
        media = MediaFileUpload(str(archivo_local), mimetype=mime_type, resumable=True)
        archivo = service.files().update(
            fileId=file_id, media_body=media, fields="id, name, webViewLink"
        ).execute()
    else:
        # Subir nuevo
        metadata = {"name": archivo_local.name, "parents": [parent_id]}
        media = MediaFileUpload(str(archivo_local), mimetype=mime_type, resumable=True)
        archivo = service.files().create(
            body=metadata, media_body=media, fields="id, name, webViewLink"
        ).execute()

    return archivo


def subir_sgsst_completo(
    service,
    output_dir: Path,
    razon_social: str,
    año: str = None,
    parent_id: str = None,
    progreso_callback=None,
) -> dict:
    """Sube todo el SG-SST generado a Google Drive replicando la estructura.

    Args:
        service: Servicio de Drive autenticado
        output_dir: Directorio local con el SG-SST generado
        razon_social: Nombre de la empresa
        año: Año del SG-SST (ej. "2026"). Se incluye en el nombre de la carpeta
             para evitar pisar versiones anteriores del mismo cliente.
        parent_id: ID carpeta padre en Drive (opcional)
        progreso_callback: Funcion(archivo_nombre, progreso) para reportar avance

    Returns:
        Dict con estadisticas y mapeo {ruta_relativa: {id, webViewLink}}
    """
    # Crear carpeta raiz — el año distingue versiones del mismo cliente
    nombre_carpeta = f"SG-SST {año} {razon_social}" if año else f"SG-SST {razon_social}"
    raiz_id = crear_carpeta(service, nombre_carpeta, parent_id)
    cache_carpetas = {"": raiz_id}
    archivos_subidos = {}
    errores = []

    # Listar archivos locales
    archivos = sorted([f for f in output_dir.rglob("*") if f.is_file()])
    total = len(archivos)

    for i, archivo in enumerate(archivos):
        ruta_relativa = archivo.relative_to(output_dir)

        try:
            # Crear carpetas intermedias
            carpeta_padre = ruta_relativa.parent
            if str(carpeta_padre) != ".":
                folder_id = _obtener_o_crear_carpeta_ruta(
                    service, carpeta_padre, raiz_id, cache_carpetas
                )
            else:
                folder_id = raiz_id

            # Subir archivo
            resultado = subir_archivo(service, archivo, folder_id)
            archivos_subidos[str(ruta_relativa)] = {
                "id": resultado.get("id"),
                "link": resultado.get("webViewLink", ""),
                "nombre": resultado.get("name", archivo.name),
            }

            if progreso_callback:
                progreso_callback(archivo.name, (i + 1) / total)

        except Exception as e:
            errores.append((str(ruta_relativa), str(e)))

    # Crear en Drive las carpetas VACIAS (sin archivos) que existen en el output.
    # Necesario para que se vea la estructura del PESV completa (FASE 2, 3, 4)
    # como recordatorio, aunque aun no tengan documentos elaborados.
    try:
        carpetas_locales = sorted(
            [d for d in output_dir.rglob("*") if d.is_dir()],
            key=lambda p: len(p.parts),  # crear primero las mas cercanas a la raiz
        )
        for carpeta in carpetas_locales:
            rel = carpeta.relative_to(output_dir)
            _obtener_o_crear_carpeta_ruta(service, rel, raiz_id, cache_carpetas)
    except Exception as e:
        errores.append(("<carpetas vacias>", str(e)))

    return {
        "raiz_id": raiz_id,
        "archivos_subidos": archivos_subidos,
        "total": total,
        "exitosos": len(archivos_subidos),
        "errores": errores,
    }


def obtener_link_carpeta(service, folder_id: str) -> str:
    """Obtiene el link web de una carpeta de Drive."""
    try:
        folder = service.files().get(
            fileId=folder_id, fields="webViewLink"
        ).execute()
        return folder.get("webViewLink", "")
    except Exception:
        return f"https://drive.google.com/drive/folders/{folder_id}"


def listar_archivos_carpeta(service, folder_id: str) -> list[dict]:
    """Lista todos los archivos en una carpeta de Drive."""
    results = []
    page_token = None

    while True:
        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name, mimeType, webViewLink, modifiedTime)",
            pageToken=page_token,
            pageSize=100,
        ).execute()

        results.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break

    return results
