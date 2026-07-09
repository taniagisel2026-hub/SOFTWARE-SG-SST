"""Configuración global del sistema."""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# Fuerza UTF-8 en stdout/stderr sin importar como se lance el proceso.
# os.environ["PYTHONIOENCODING"] (como hacia cli.py) no sirve: el interprete
# ya arranco con su codificacion, esa variable solo aplica a subprocesos
# nuevos. reconfigure() si cambia la codificacion del stream ya abierto.
# Sin esto, corridas sin terminal interactiva (ej. bajo uvicorn) truenan
# con UnicodeEncodeError en cualquier caracter fuera de cp1252 (ej. "→").
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).parent.parent
PLANTILLA_PATH = Path(os.getenv("PLANTILLA_PATH", BASE_DIR / "Plantilla SG-SST 2026"))
OUTPUT_BASE_PATH = Path(os.getenv("OUTPUT_BASE_PATH", BASE_DIR / "clientes"))
CLIENTES_CONFIG_PATH = BASE_DIR / "clientes_config"
CACHE_DIR = Path(os.getenv("CACHE_DIR", BASE_DIR / ".cache"))
CACHE_EXPIRY_DAYS = int(os.getenv("CACHE_EXPIRY_DAYS", "30"))
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-haiku-4-5-20251001")

# Precios Claude API (USD por millon de tokens) - actualizar si cambian
# https://docs.anthropic.com/en/docs/about-claude/pricing
CLAUDE_PRECIO_INPUT_MTOK = float(os.getenv("CLAUDE_PRECIO_INPUT", "1"))    # $/MTok input (Haiku 4.5)
CLAUDE_PRECIO_OUTPUT_MTOK = float(os.getenv("CLAUDE_PRECIO_OUTPUT", "5"))  # $/MTok output (Haiku 4.5)

# API REST (Fase 4)
API_AUTH_TOKEN = os.getenv("API_AUTH_TOKEN", "")
API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:8000")
JOBS_DB_PATH = Path(os.getenv("JOBS_DB_PATH", BASE_DIR / "jobs.db"))

# Asistente de onboarding por modulos (Fase 5)
WIZARD_DB_PATH = Path(os.getenv("WIZARD_DB_PATH", BASE_DIR / "wizard_drafts.db"))
WIZARD_UPLOADS_DIR = Path(os.getenv("WIZARD_UPLOADS_DIR", BASE_DIR / "wizard_uploads"))

# Tipos de documento y su nivel de humanización con Claude
DOC_TYPES = {
    "POL": {"nombre": "Política", "humanizar": "completa", "prefijo": "SST-POL-"},
    "PRO": {"nombre": "Procedimiento", "humanizar": "media", "prefijo": "SST-PRO-"},
    # Formatos: humanizacion "media" para adaptar correctamente COPASST/Vigia
    # segun el numero de trabajadores del cliente. Los campos con guiones bajos
    # (_____) y lineas cortas se filtran automaticamente, solo se humanizan
    # los parrafos de texto real (introducciones, objetivos, alcances).
    "FMT": {"nombre": "Formato", "humanizar": "media", "prefijo": "SST-FMT-"},
    "FTM": {"nombre": "Formato", "humanizar": "media", "prefijo": "SST-FTM-"},
    "MTZ": {"nombre": "Matriz", "humanizar": "ninguna", "prefijo": "SST-MTZ-"},
    "PRG": {"nombre": "Programa", "humanizar": "completa", "prefijo": "SST-PRG-"},
    "PLN": {"nombre": "Plan", "humanizar": "media", "prefijo": "SST-PLN-"},
    "MNL": {"nombre": "Manual", "humanizar": "completa", "prefijo": "SST-MNL-"},
    "RGL": {"nombre": "Reglamento", "humanizar": "media", "prefijo": "SST-RGL-"},
    "GUA": {"nombre": "Guía", "humanizar": "completa", "prefijo": "SST-GUA-"},
}
