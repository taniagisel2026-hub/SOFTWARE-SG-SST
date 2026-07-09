"""Persistencia del asistente de onboarding por modulos (Fase 5).

Un "borrador" acumula los datos de cada modulo hasta quedar listo para
convertirse en un YAML de cliente real. Usa SQLite (libreria estandar),
mismo patron que engine/jobs.py: el bloqueo entre modulos (no se puede
confirmar el modulo N sin haber confirmado el N-1) se verifica en la capa
de la API (wizard_api.py) leyendo `modulos_confirmados`, no aqui.
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Iterator

from engine.config import WIZARD_DB_PATH

MODULOS_ORDEN = [
    "modulo_1", "modulo_2", "modulo_3", "modulo_4",
    "modulo_5", "modulo_6", "modulo_7",
]


def _ahora_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def _conexion() -> Iterator[sqlite3.Connection]:
    conn = sqlite3.connect(str(WIZARD_DB_PATH))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def inicializar_db() -> None:
    """Crea la tabla `wizard_drafts` si no existe. Se llama al importar este modulo."""
    with _conexion() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS wizard_drafts (
                borrador_id TEXT PRIMARY KEY,
                modulo_actual INTEGER NOT NULL DEFAULT 1,
                modulos_confirmados TEXT NOT NULL DEFAULT '[]',
                datos_json TEXT NOT NULL DEFAULT '{}',
                archivos_json TEXT NOT NULL DEFAULT '{}',
                extraccion_json TEXT NOT NULL DEFAULT '{}',
                estado TEXT NOT NULL DEFAULT 'en_progreso',
                config_path_final TEXT,
                creado_en TEXT NOT NULL,
                actualizado_en TEXT NOT NULL
            )
        """)


inicializar_db()


def _row_a_dict(row: sqlite3.Row) -> dict:
    return {
        "borrador_id": row["borrador_id"],
        "modulo_actual": row["modulo_actual"],
        "modulos_confirmados": json.loads(row["modulos_confirmados"]),
        "datos": json.loads(row["datos_json"]),
        "archivos": json.loads(row["archivos_json"]),
        "extraccion": json.loads(row["extraccion_json"]),
        "estado": row["estado"],
        "config_path_final": row["config_path_final"],
        "creado_en": row["creado_en"],
        "actualizado_en": row["actualizado_en"],
    }


def crear_borrador() -> str:
    """Crea un borrador nuevo en modulo_actual=1. Retorna el borrador_id."""
    borrador_id = uuid.uuid4().hex
    ahora = _ahora_iso()
    with _conexion() as conn:
        conn.execute(
            """INSERT INTO wizard_drafts
               (borrador_id, creado_en, actualizado_en)
               VALUES (?, ?, ?)""",
            (borrador_id, ahora, ahora),
        )
    return borrador_id


def obtener_borrador(borrador_id: str) -> dict | None:
    with _conexion() as conn:
        row = conn.execute(
            "SELECT * FROM wizard_drafts WHERE borrador_id = ?", (borrador_id,)
        ).fetchone()
    return _row_a_dict(row) if row else None


def guardar_extraccion(borrador_id: str, doc_rol: str, extraido: dict) -> None:
    """Guarda el resultado de extraer un documento (rut/camara_comercio/cedula_rl),
    sin combinarlo con los demas -- eso lo hace la capa de la API al detectar conflictos."""
    borrador = obtener_borrador(borrador_id)
    if borrador is None:
        raise ValueError(f"No existe el borrador: {borrador_id}")
    extraccion = borrador["extraccion"]
    extraccion[doc_rol] = extraido
    with _conexion() as conn:
        conn.execute(
            "UPDATE wizard_drafts SET extraccion_json = ?, actualizado_en = ? WHERE borrador_id = ?",
            (json.dumps(extraccion, ensure_ascii=False), _ahora_iso(), borrador_id),
        )


def guardar_archivo(borrador_id: str, rol: str, ruta_relativa: str) -> None:
    borrador = obtener_borrador(borrador_id)
    if borrador is None:
        raise ValueError(f"No existe el borrador: {borrador_id}")
    archivos = borrador["archivos"]
    archivos[rol] = ruta_relativa
    with _conexion() as conn:
        conn.execute(
            "UPDATE wizard_drafts SET archivos_json = ?, actualizado_en = ? WHERE borrador_id = ?",
            (json.dumps(archivos, ensure_ascii=False), _ahora_iso(), borrador_id),
        )


def guardar_modulo(borrador_id: str, modulo_key: str, datos: dict) -> dict:
    """Guarda los datos confirmados de un modulo y avanza modulo_actual.

    No valida el orden aqui -- ese chequeo (409 si se salta un modulo) vive
    en wizard_api.py, que llama a esta funcion solo despues de validarlo.
    """
    borrador = obtener_borrador(borrador_id)
    if borrador is None:
        raise ValueError(f"No existe el borrador: {borrador_id}")

    datos_totales = borrador["datos"]
    datos_totales[modulo_key] = datos

    confirmados = borrador["modulos_confirmados"]
    if modulo_key not in confirmados:
        confirmados.append(modulo_key)

    indice = MODULOS_ORDEN.index(modulo_key)
    modulo_actual = max(borrador["modulo_actual"], indice + 2)  # +2: 1-indexado y avanza al siguiente

    with _conexion() as conn:
        conn.execute(
            """UPDATE wizard_drafts
               SET datos_json = ?, modulos_confirmados = ?, modulo_actual = ?, actualizado_en = ?
               WHERE borrador_id = ?""",
            (
                json.dumps(datos_totales, ensure_ascii=False),
                json.dumps(confirmados, ensure_ascii=False),
                modulo_actual,
                _ahora_iso(),
                borrador_id,
            ),
        )
    return obtener_borrador(borrador_id)


def saltar_modulo(borrador_id: str, modulo_key: str) -> dict:
    """Marca un modulo opcional (ej. modulo_7, cargos) como confirmado sin datos."""
    return guardar_modulo(borrador_id, modulo_key, {"omitido": True})


def marcar_finalizado(borrador_id: str, config_path: str) -> None:
    with _conexion() as conn:
        conn.execute(
            """UPDATE wizard_drafts SET estado = 'finalizado', config_path_final = ?, actualizado_en = ?
               WHERE borrador_id = ?""",
            (config_path, _ahora_iso(), borrador_id),
        )


def listar_borradores(estado: str | None = None) -> list[dict]:
    with _conexion() as conn:
        if estado:
            rows = conn.execute(
                "SELECT * FROM wizard_drafts WHERE estado = ? ORDER BY creado_en DESC", (estado,)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM wizard_drafts ORDER BY creado_en DESC"
            ).fetchall()
    return [_row_a_dict(row) for row in rows]


def eliminar_borrador(borrador_id: str) -> None:
    with _conexion() as conn:
        conn.execute("DELETE FROM wizard_drafts WHERE borrador_id = ?", (borrador_id,))
