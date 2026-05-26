"""Configuración global del sistema."""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

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
