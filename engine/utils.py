"""Utilidades compartidas - fusionadas del script v7.6."""
import os
import re
import unicodedata
import urllib.request
from datetime import date, datetime
from urllib.parse import urlparse


# ══════════════════════════════════════════════════════════════════════
#  RUTAS LARGAS WINDOWS (>260 caracteres)
# ══════════════════════════════════════════════════════════════════════

def get_long_path(path: str) -> str:
    """Agrega prefijo \\\\?\\ en Windows para soportar rutas >260 chars."""
    if os.name == "nt":
        abs_path = os.path.abspath(path)
        if not abs_path.startswith("\\\\?\\"):
            return "\\\\?\\" + abs_path
    return path


# ══════════════════════════════════════════════════════════════════════
#  NORMALIZACIÓN DE CLAVES DE MARCADORES
# ══════════════════════════════════════════════════════════════════════

def norm_key(s: str) -> str:
    """Normaliza una clave de marcador removiendo llaves, tildes y espacios."""
    s = "" if s is None else str(s)
    s = s.replace("{{", "").replace("}}", "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip()


def key_variants(key: str) -> list:
    """Genera todas las variantes posibles de una clave de marcador.

    Esto es clave porque Word a veces cambia mayúsculas, tildes, etc.
    Genera variantes con/sin tildes, con/sin guiones bajos, upper/lower.
    """
    k = norm_key(key)
    if not k:
        return []

    variants = {k, k.upper(), k.lower()}
    k_space = re.sub(r"\s+", " ", k)
    variants.update([
        k_space, k_space.replace(" ", "_"), k_space.replace(" ", ""),
        k_space.upper(), k_space.upper().replace(" ", "_"), k_space.upper().replace(" ", ""),
        k_space.lower(), k_space.lower().replace(" ", "_"), k_space.lower().replace(" ", ""),
    ])

    k_clean = re.sub(r"[^0-9A-Za-z_ ]+", "", k_space)
    variants.update([
        k_clean, k_clean.replace(" ", "_"), k_clean.replace(" ", ""),
        k_clean.upper(), k_clean.upper().replace(" ", "_"), k_clean.lower().replace(" ", "_"),
    ])

    return [v for v in variants if v]


def format_value(v) -> str:
    """Formatea un valor para insertar en documento."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, str):
        s = v.strip()
        s = re.sub(r"\s+00:00:00(\.0+)?$", "", s)
        return s
    return str(v)


# ══════════════════════════════════════════════════════════════════════
#  SOPORTE DE IMÁGENES (local + URL)
# ══════════════════════════════════════════════════════════════════════

def looks_like_image_path(s: str) -> bool:
    """Detecta si un string parece ser una ruta a imagen o URL de imagen."""
    s = s.strip().strip('"').strip("'")
    if s.lower().startswith(("http://", "https://")):
        return True
    return os.path.splitext(s)[1].lower() in {".png", ".jpg", ".jpeg"}


def download_image_if_url(value: str, temp_dir: str, timeout: int = 5):
    """Descarga imagen desde URL a un archivo temporal. Retorna path o None."""
    raw = value.strip().strip('"').strip("'")
    if not raw.lower().startswith(("http://", "https://")):
        return None
    parsed = urlparse(raw)
    ext = os.path.splitext(parsed.path)[1].lower()
    if ext not in {".png", ".jpg", ".jpeg"}:
        ext = ".jpg"
    local_path = os.path.join(temp_dir, f"img_{abs(hash(raw))}{ext}")
    try:
        req = urllib.request.Request(raw, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        with open(local_path, "wb") as f:
            f.write(data)
        return local_path
    except Exception:
        return None


def resolve_image_path(value: str, fallback_dirs: list):
    """Resuelve la ruta de una imagen probando múltiples directorios."""
    raw = value.strip().strip('"').strip("'")
    if raw.lower().startswith(("http://", "https://")):
        return None
    candidates = [raw] if os.path.isabs(raw) else [os.path.join(d, raw) for d in fallback_dirs]
    return next((p for p in candidates if os.path.exists(p)), None)


def encontrar_ruta_logo(contexto: dict, fallback_dirs: list, temp_dir: str,
                        ruta_logo_marca_agua: str = "") -> str:
    """Busca el logo en el contexto para usar como marca de agua."""
    if ruta_logo_marca_agua and os.path.exists(ruta_logo_marca_agua):
        return ruta_logo_marca_agua

    for k, v in contexto.items():
        if "logo" in k.lower() and isinstance(v, str) and looks_like_image_path(v):
            img_path = download_image_if_url(v, temp_dir)
            if not img_path:
                img_path = resolve_image_path(v, fallback_dirs)
            if img_path:
                return img_path
    return None
