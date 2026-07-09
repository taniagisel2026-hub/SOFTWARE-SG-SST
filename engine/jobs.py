"""Persistencia de trabajos en segundo plano (generacion y subida a Drive) para la API REST.

Usa SQLite (libreria estandar) para que el estado de un trabajo sobreviva un
reinicio del proceso de la API -- mismo patron que ya usa `diskcache` (tambien
basado en SQLite) para el cache de Claude en engine/claude_service.py.
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from engine.config import JOBS_DB_PATH

ESTADOS_ACTIVOS = ("queued", "running")


def _ahora_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def _conexion() -> Iterator[sqlite3.Connection]:
    conn = sqlite3.connect(str(JOBS_DB_PATH))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def inicializar_db() -> None:
    """Crea la tabla `jobs` si no existe. Se llama al importar este modulo."""
    with _conexion() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY,
                tipo TEXT NOT NULL,
                cliente_id TEXT NOT NULL,
                status TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                actualizado_en TEXT NOT NULL,
                params_json TEXT,
                resultado_json TEXT,
                error TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_cliente ON jobs(cliente_id)")


inicializar_db()


def _row_a_dict(row: sqlite3.Row) -> dict:
    return {
        "job_id": row["job_id"],
        "tipo": row["tipo"],
        "cliente_id": row["cliente_id"],
        "status": row["status"],
        "creado_en": row["creado_en"],
        "actualizado_en": row["actualizado_en"],
        "params": json.loads(row["params_json"]) if row["params_json"] else None,
        "resultado": json.loads(row["resultado_json"]) if row["resultado_json"] else None,
        "error": row["error"],
    }


def crear_job(tipo: str, cliente_id: str, params: dict) -> str:
    """Crea un job en estado 'queued'. Retorna el job_id."""
    job_id = uuid.uuid4().hex
    ahora = _ahora_iso()
    with _conexion() as conn:
        conn.execute(
            """INSERT INTO jobs
               (job_id, tipo, cliente_id, status, creado_en, actualizado_en, params_json)
               VALUES (?, ?, ?, 'queued', ?, ?, ?)""",
            (job_id, tipo, cliente_id, ahora, ahora, json.dumps(params, ensure_ascii=False, default=str)),
        )
    return job_id


def actualizar_job(
    job_id: str,
    status: str,
    resultado: dict | None = None,
    error: str | None = None,
) -> None:
    """Actualiza el estado (y opcionalmente resultado/error) de un job existente."""
    with _conexion() as conn:
        conn.execute(
            """UPDATE jobs SET status = ?, actualizado_en = ?, resultado_json = ?, error = ?
               WHERE job_id = ?""",
            (
                status,
                _ahora_iso(),
                json.dumps(resultado, ensure_ascii=False, default=str) if resultado is not None else None,
                error,
                job_id,
            ),
        )


def obtener_job(job_id: str) -> dict | None:
    with _conexion() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE job_id = ?", (job_id,)).fetchone()
    return _row_a_dict(row) if row else None


def listar_jobs(cliente_id: str, tipo: str | None = None) -> list[dict]:
    with _conexion() as conn:
        if tipo:
            rows = conn.execute(
                "SELECT * FROM jobs WHERE cliente_id = ? AND tipo = ? ORDER BY creado_en DESC",
                (cliente_id, tipo),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM jobs WHERE cliente_id = ? ORDER BY creado_en DESC",
                (cliente_id,),
            ).fetchall()
    return [_row_a_dict(row) for row in rows]


def marcar_jobs_huerfanos_como_fallidos() -> int:
    """Marca como 'failed' cualquier job que quedo en 'queued'/'running' de una corrida
    anterior (el proceso en segundo plano que lo ejecutaba ya no existe tras un reinicio).

    Se llama al arrancar la API. Retorna cuantos jobs se marcaron.
    """
    with _conexion() as conn:
        cursor = conn.execute(
            f"""UPDATE jobs SET status = 'failed', actualizado_en = ?, error = ?
                WHERE status IN ({','.join('?' * len(ESTADOS_ACTIVOS))})""",
            (_ahora_iso(), "Interrumpido por reinicio del servidor", *ESTADOS_ACTIVOS),
        )
        return cursor.rowcount


def ejecutar_generacion_en_background(
    job_id: str,
    config_path: str,
    output_dir: str | None,
    humanizar: bool,
    solo_marcadores: bool,
    estricto: bool,
) -> None:
    """Wrapper que corre generar_sgsst() y persiste el resultado/error en el job."""
    from engine.generador import generar_sgsst  # import diferido: evita ciclos con cli.py

    actualizar_job(job_id, "running")
    try:
        resultado = generar_sgsst(
            config_path=config_path,
            output_dir=output_dir,
            humanizar=humanizar,
            solo_marcadores=solo_marcadores,
            estricto=estricto,
        )
        actualizar_job(job_id, "done", resultado=resultado)
    except Exception as e:
        actualizar_job(job_id, "failed", error=str(e))


def ejecutar_subida_en_background(
    job_id: str,
    cliente: dict,
    output_dir: str,
    cliente_id: str,
    actualizar_listado: bool,
) -> None:
    """Wrapper que corre ejecutar_subida_drive() y persiste el resultado/error en el job."""
    from engine.drive_upload import ejecutar_subida_drive  # import diferido: evita ciclos

    actualizar_job(job_id, "running")
    try:
        resultado = ejecutar_subida_drive(cliente, Path(output_dir), cliente_id, actualizar_listado)
        actualizar_job(job_id, "done", resultado=resultado)
    except Exception as e:
        actualizar_job(job_id, "failed", error=str(e))
