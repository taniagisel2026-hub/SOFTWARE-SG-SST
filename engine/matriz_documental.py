"""
engine/matriz_documental.py

Filtra SST-MTZ-005 Matriz de Gestión Documental según la categoría
del cliente determinada por Res. 0312/2019 (Arts. 3, 9 y 16).

Categoría I  (Art. 3)  — ≤10 trab, riesgo I-III  →  7 estándares
Categoría II (Art. 9)  — 11-50 trab, riesgo I-III → 21 estándares
Categoría III (Art. 16) — ≥51 trab o riesgo IV-V  → 60 estándares (sin filtro)
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl


# ─── Estándares por categoría ─────────────────────────────────────────────────

# Artículo 3 — 7 ítems
# Nota: 1.1.8 (CCL) no es uno de los 7 estándares del Art. 3, pero el
# Comité de Convivencia Laboral es obligatorio para TODAS las empresas
# por Res. 652/2012, independientemente de la categoría Res. 0312.
ESTANDARES_CAT_I: frozenset[str] = frozenset({
    "1.1.1",                                    # Responsable / persona SG-SST
    "1.1.4",                                    # Afiliación SGRL
    "1.1.8",                                    # CCL — obligatorio Res. 652/2012
    "1.2.1",                                    # Capacitación en SST
    "2.4.1",                                    # Plan Anual de Trabajo
    "2.5.1",                                    # Archivo documental — siempre incluir MTZ-005
    "3.1.4",                                    # Evaluaciones médicas ocupacionales
    "4.1.1", "4.1.2",                           # Identificación de peligros
    "4.2.1", "4.2.2", "4.2.3", "4.2.4",        # Medidas de prevención y control
    "4.2.5", "4.2.6",                           # Mantenimiento / EPP
})

# Artículo 9 — 21 ítems
ESTANDARES_CAT_II: frozenset[str] = frozenset({
    "1.1.1", "1.1.3", "1.1.4", "1.1.6", "1.1.8",   # Recursos y comités
    "1.2.1",                                          # Capacitación
    "2.1.1", "2.4.1", "2.5.1",                       # Política, plan, archivo
    "3.1.1", "3.1.2", "3.1.4", "3.1.6",              # Condiciones de salud
    "3.2.1", "3.2.2",                                 # Reporte e investigación AT
    "4.1.1", "4.1.2",                                 # Identificación peligros
    "4.2.5", "4.2.6",                                 # Mantenimiento / EPP
    "5.1.1", "5.1.2",                                 # Plan y brigada emergencias
    "6.1.3",                                          # Revisión alta dirección
})

# Categoría III → sin filtro (todos los estándares)

_RE_ESTANDAR = re.compile(r"^(\d+\.\d+\.\d+)")
_RE_ROMANO = re.compile(r"\b(V|IV|III|II|I)\b")
_ORDEN_RIESGO = {"V": 5, "IV": 4, "III": 3, "II": 2, "I": 1}


def normalizar_nivel_riesgo(valor: str) -> str:
    """
    Extrae el nivel de riesgo más alto de una cadena.
    "I - IV" → "IV", "III" → "III", "V" → "V".
    """
    encontrados = _RE_ROMANO.findall(str(valor).upper())
    if not encontrados:
        return "I"
    return max(encontrados, key=lambda r: _ORDEN_RIESGO.get(r, 0))

# Reemplazos COPASST → Vigía para celdas del Excel
_REEMPLAZOS_VIGIA = [
    ("Conformación COPASST",        "Designación del Vigía SST"),
    ("CONFORMACIÓN COPASST",        "DESIGNACIÓN DEL VIGÍA SST"),
    ("Capacitación COPASST",        "Capacitación Vigía SST"),
    ("CAPACITACIÓN COPASST",        "CAPACITACIÓN VIGÍA SST"),
    ("COPASST",                     "VIGÍA SST"),
    ("Copasst",                     "Vigía SST"),
]


# ─── Funciones públicas ───────────────────────────────────────────────────────

def determinar_categoria(nivel_riesgo: str, num_trabajadores: int) -> str:
    """Retorna 'I', 'II' o 'III' según Res. 0312/2019."""
    if nivel_riesgo in ("IV", "V"):
        return "III"
    if num_trabajadores > 50:
        return "III"
    if 11 <= num_trabajadores <= 50 and nivel_riesgo in ("I", "II", "III"):
        return "II"
    if num_trabajadores <= 10 and nivel_riesgo in ("I", "II", "III"):
        return "I"
    raise ValueError(
        f"Combinación no reconocida: {num_trabajadores} trabajadores, riesgo {nivel_riesgo}"
    )


def filtrar_matriz_documental(
    output_base: Path,
    nivel_riesgo: str,
    num_trabajadores: int,
    usar_vigia: bool = False,
) -> dict:
    """
    Localiza SST-MTZ-005 en output_base, determina la categoría del cliente
    y elimina las filas de estándares que no aplican.

    Retorna un dict con:
        categoria       — 'I', 'II' o 'III'
        filas_eliminadas — número de filas borradas
        articulo        — referencia legal aplicada
        error           — mensaje si algo falló (opcional)
    """
    mtz005 = next(
        (f for f in output_base.rglob("*.xlsx")
         if "MTZ-005" in f.name and not f.name.startswith("~$")),
        None,
    )
    if not mtz005:
        return {"categoria": None, "filas_eliminadas": 0,
                "error": "SST-MTZ-005 no encontrado en el output"}

    try:
        categoria = determinar_categoria(nivel_riesgo, num_trabajadores)
    except ValueError as e:
        return {"categoria": None, "filas_eliminadas": 0, "error": str(e)}

    _ARTICULOS = {"I": "Artículo 3", "II": "Artículo 9", "III": "Artículo 16"}
    articulo = _ARTICULOS[categoria]

    if categoria == "III":
        # Sin filtro — se adapta COPASST→Vigía si aplica y listo
        if usar_vigia:
            _adaptar_vigia_xlsx(mtz005)
        return {"categoria": "III", "filas_eliminadas": 0, "articulo": articulo}

    estandares_validos = (
        ESTANDARES_CAT_I if categoria == "I" else ESTANDARES_CAT_II
    )

    wb = openpyxl.load_workbook(str(mtz005))
    ws = wb["LISTADO MAESTRO PHVA"]

    filas_eliminar = []
    estandar_actual: str | None = None

    for fila_num in range(7, ws.max_row + 1):
        val_c = ws.cell(row=fila_num, column=3).value
        codigo = _extraer_codigo(val_c)
        if codigo:
            estandar_actual = codigo
        if estandar_actual and estandar_actual not in estandares_validos:
            filas_eliminar.append(fila_num)

    for fila_num in reversed(filas_eliminar):
        ws.delete_rows(fila_num)

    if usar_vigia:
        _adaptar_vigia_en_hoja(ws)

    wb.save(str(mtz005))
    return {
        "categoria": categoria,
        "filas_eliminadas": len(filas_eliminar),
        "articulo": articulo,
    }


# ─── Utilidades internas ──────────────────────────────────────────────────────

def _extraer_codigo(valor) -> str | None:
    """Extrae el código numérico 'X.X.X' del inicio de una celda col C."""
    if not valor:
        return None
    m = _RE_ESTANDAR.match(str(valor).strip())
    return m.group(1) if m else None


def _reemplazar_en_celda(ws, fila: int, col: int) -> None:
    """Aplica los reemplazos COPASST→Vigía en una celda específica."""
    celda = ws.cell(row=fila, column=col)
    if not celda.value or not isinstance(celda.value, str):
        return
    texto = celda.value
    for viejo, nuevo in _REEMPLAZOS_VIGIA:
        texto = texto.replace(viejo, nuevo)
    celda.value = texto


def _adaptar_vigia_en_hoja(ws) -> None:
    """Reemplaza COPASST→Vigía en columnas B, C y F de la hoja PHVA."""
    for fila_num in range(1, ws.max_row + 1):
        _reemplazar_en_celda(ws, fila_num, 2)   # col B — Estándar padre
        _reemplazar_en_celda(ws, fila_num, 3)   # col C — Sub-estándar
        _reemplazar_en_celda(ws, fila_num, 6)   # col F — Nombre del documento


def _adaptar_vigia_xlsx(ruta: Path) -> None:
    """Versión para Cat III (sin filtro de filas): solo adapta el texto."""
    wb = openpyxl.load_workbook(str(ruta))
    ws = wb["LISTADO MAESTRO PHVA"]
    _adaptar_vigia_en_hoja(ws)
    wb.save(str(ruta))
