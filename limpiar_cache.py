"""Limpia los caches del sistema SG-SST.

Uso:
    python limpiar_cache.py              # Limpia solo Claude (default)
    python limpiar_cache.py --todo        # Limpia Claude + cache pytest + pycache
    python limpiar_cache.py --solo-claude # Solo respuestas de Claude
    python limpiar_cache.py --tokens      # Tambien tokens OAuth de Drive (forzara re-login)
"""
import shutil
import sys
from pathlib import Path

BASE = Path(__file__).parent

CACHES = {
    "claude": BASE / ".cache" / "claude_responses",
    "pycache": list(BASE.rglob("__pycache__")),
    "pytest": BASE / ".pytest_cache",
    "tokens": BASE / "tokens",
}


def limpiar(que: list[str]):
    if "claude" in que:
        d = CACHES["claude"]
        if d.exists():
            shutil.rmtree(d)
            print(f"[OK] Limpiado: {d}")
        else:
            print(f"[SKIP] Ya limpio: {d}")

    if "pycache" in que:
        n = 0
        for d in CACHES["pycache"]:
            try:
                shutil.rmtree(d)
                n += 1
            except Exception:
                pass
        print(f"[OK] Eliminadas {n} carpetas __pycache__")

    if "pytest" in que:
        d = CACHES["pytest"]
        if d.exists():
            shutil.rmtree(d)
            print(f"[OK] Limpiado: {d}")
        else:
            print(f"[SKIP] Ya limpio: {d}")

    if "tokens" in que:
        d = CACHES["tokens"]
        if d.exists():
            for f in d.glob("token_*.json"):
                f.unlink()
                print(f"[OK] Eliminado token: {f.name}")
            print("    (la proxima subida pedira login en Drive)")


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--todo" in args:
        limpiar(["claude", "pycache", "pytest"])
    elif "--solo-claude" in args:
        limpiar(["claude"])
    elif "--tokens" in args:
        limpiar(["claude", "tokens"])
    else:
        limpiar(["claude"])

    print("\nListo.")
