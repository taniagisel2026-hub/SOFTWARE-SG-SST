"""Extraccion estructurada de datos desde texto de RUT / Camara de Comercio /
Cedula, para el asistente de onboarding por modulos (Fase 5).

Sigue el mismo patron de JSON-por-prompt ya usado en
engine/matriz_peligros.py::_generar_datos_cargo(): prompt pidiendo JSON puro,
se le quitan los fences de markdown si los trae, json.loads, y ante
cualquier falla se retorna {} en vez de lanzar excepcion -- el usuario
siempre puede llenar los campos a mano si la extraccion no funciono
(ej. un PDF escaneado sin capa de texto).

Cada documento se extrae por separado (nunca se combinan aqui) para que la
capa de arriba (wizard_api.py) pueda detectar y mostrar conflictos --por
ejemplo, direcciones distintas en el RUT y en la Camara de Comercio-- en
vez de que uno pise silenciosamente al otro.
"""
from __future__ import annotations

import json
import logging
import re

from . import claude_service
from .config import ANTHROPIC_API_KEY, CLAUDE_MODEL

logger = logging.getLogger(__name__)

# Campos que se le pide a Claude que extraiga, por tipo de documento.
CAMPOS_POR_TIPO = {
    "rut": [
        "razon_social", "nit", "direccion_empresa", "ciudad_empresa",
        "departamento_empresa", "actividad_economica", "telefono_empresa",
        "email_empresa",
    ],
    "camara_comercio": [
        "razon_social", "nit", "direccion_empresa", "ciudad_empresa",
        "departamento_empresa", "actividad_economica", "representante_legal",
        "cedula_rl", "cargo_rl",
    ],
    "cedula": [
        "representante_legal", "cedula_rl",
    ],
}

_NOMBRES_DOCUMENTO = {
    "rut": "RUT (Registro Unico Tributario)",
    "camara_comercio": "Certificado de Camara de Comercio",
    "cedula": "Cedula de Ciudadania",
}


def _construir_prompt(tipo: str, texto: str) -> str:
    campos = CAMPOS_POR_TIPO[tipo]
    nombre_doc = _NOMBRES_DOCUMENTO[tipo]
    campos_lista = "\n".join(f'  "{c}": "..."' for c in campos)
    return f"""Eres un asistente que extrae datos estructurados de documentos legales colombianos.

A continuacion el texto extraido de un {nombre_doc}:

---
{texto[:8000]}
---

Extrae EXACTAMENTE estos campos (usa "" si el campo no aparece en el texto, no inventes datos):
{{
{campos_lista}
}}

Genera EXCLUSIVAMENTE un objeto JSON valido con esos campos. Devuelve SOLO el JSON, sin explicaciones ni texto adicional."""


def extraer_datos_documento(tipo: str, texto: str) -> dict:
    """Extrae campos estructurados del texto de un documento via Claude.

    Args:
        tipo: "rut" | "camara_comercio" | "cedula"
        texto: texto plano ya extraido del PDF (ver engine/pdf_utils.py)

    Returns:
        dict con los campos de CAMPOS_POR_TIPO[tipo] (puede venir incompleto
        o vacio). Nunca lanza excepcion -- retorna {} si Claude no esta
        configurado, si el texto viene vacio (PDF escaneado sin capa de
        texto), o si la llamada/parseo falla por cualquier razon.
    """
    if tipo not in CAMPOS_POR_TIPO:
        raise ValueError(f"Tipo de documento desconocido: {tipo}")

    if not texto or not texto.strip():
        logger.warning("Texto vacio para extraccion de %s (¿PDF escaneado sin texto?)", tipo)
        return {}

    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada, no se puede extraer %s", tipo)
        return {}

    try:
        client = claude_service._crear_client()
        prompt = _construir_prompt(tipo, texto)
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        claude_service._registrar_uso(response)
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        # Solo quedarse con los campos esperados, por si Claude agrega de mas
        return {c: data.get(c, "") for c in CAMPOS_POR_TIPO[tipo]}
    except Exception as e:
        logger.warning("No se pudo extraer datos de %s: %s", tipo, e)
        return {}


def detectar_conflictos(extracciones: dict[str, dict]) -> list[dict]:
    """Compara los campos comunes entre dos o mas documentos extraidos y
    reporta donde difieren (ej. direccion distinta en RUT vs Camara de
    Comercio), en vez de que un valor pise silenciosamente al otro.

    Args:
        extracciones: {"rut": {...}, "camara_comercio": {...}, ...}

    Returns:
        lista de {"campo": str, "valores": {"rut": "...", "camara_comercio": "..."}}
        para cada campo donde dos o mas documentos no coinciden (normalizado
        por espacios/mayusculas antes de comparar).
    """
    campos_comparables = [
        "razon_social", "nit", "direccion_empresa", "ciudad_empresa", "departamento_empresa",
    ]
    conflictos = []
    for campo in campos_comparables:
        valores_por_doc = {}
        for doc_rol, datos in extracciones.items():
            valor = (datos or {}).get(campo, "").strip()
            if valor:
                valores_por_doc[doc_rol] = valor

        if len(valores_por_doc) < 2:
            continue

        normalizados = {" ".join(v.upper().split()) for v in valores_por_doc.values()}
        if len(normalizados) > 1:
            conflictos.append({"campo": campo, "valores": valores_por_doc})

    return conflictos
