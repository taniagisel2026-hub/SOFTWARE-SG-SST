"""Extraccion de texto de archivos PDF, compartida por matriz_peligros.py
(manuales de cargos) y por el asistente de onboarding (RUT/Camara de Comercio).
"""
from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def extraer_texto_pdf(ruta: Path) -> str:
    """Extrae texto util de un .pdf (parrafos + tablas).
    Usa pdfplumber por buena extraccion de tablas; cae a pypdf si falla.
    Nunca lanza excepcion: retorna "" si no se pudo leer nada.
    """
    # Primer intento: pdfplumber (mejor para tablas)
    try:
        import pdfplumber
        lineas = []
        with pdfplumber.open(str(ruta)) as pdf:
            for pagina in pdf.pages:
                # Texto plano
                texto = pagina.extract_text() or ""
                for ln in texto.split("\n"):
                    ln = ln.strip()
                    if ln:
                        lineas.append(ln)
                # Tablas (formateadas como "col1 | col2 | col3")
                try:
                    tablas = pagina.extract_tables() or []
                except Exception:
                    tablas = []
                for tbl in tablas:
                    for row in tbl:
                        vistos = set()
                        partes = []
                        for cell in row:
                            if cell is None:
                                continue
                            t = str(cell).strip().replace("\n", " ")
                            if t and t not in vistos:
                                vistos.add(t)
                                partes.append(t)
                        if partes:
                            lineas.append(" | ".join(partes))
        return "\n".join(lineas)
    except ImportError:
        logger.warning("pdfplumber no instalado, usando pypdf")
    except Exception as e:
        logger.warning("pdfplumber fallo en %s: %s — intentando pypdf", ruta, e)

    # Fallback: pypdf (mas basico pero suele funcionar)
    try:
        from pypdf import PdfReader
    except ImportError:
        logger.error("Ni pdfplumber ni pypdf instalados, no se puede leer %s", ruta)
        return ""
    try:
        reader = PdfReader(str(ruta))
        lineas = []
        for pagina in reader.pages:
            texto = pagina.extract_text() or ""
            for ln in texto.split("\n"):
                ln = ln.strip()
                if ln:
                    lineas.append(ln)
        return "\n".join(lineas)
    except Exception as e:
        logger.warning("No se pudo leer PDF %s: %s", ruta, e)
        return ""
