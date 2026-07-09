"""Autenticacion minima de la API REST: un API key estatico compartido.

No es un sistema de login (eso es Fase 5) -- solo evita dejar los endpoints
abiertos a cualquiera en internet mientras no exista autenticacion por usuario.
"""
from fastapi import Header, HTTPException

from engine.config import API_AUTH_TOKEN


def verify_api_key(authorization: str | None = Header(default=None)) -> None:
    if not API_AUTH_TOKEN:
        raise HTTPException(
            status_code=500,
            detail="API_AUTH_TOKEN no esta configurado en .env",
        )
    if authorization != f"Bearer {API_AUTH_TOKEN}":
        raise HTTPException(status_code=401, detail="No autorizado")
