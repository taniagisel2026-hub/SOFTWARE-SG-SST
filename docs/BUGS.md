# Registro de Bugs

Documenta aquí cada bug encontrado mientras pruebas con clientes reales. Sirve para:

1. Que el equipo técnico pueda reproducirlos y arreglarlos
2. Convertir cada bug en un **regression test** para que no vuelva a aparecer
3. Llevar historial de problemas resueltos

---

## Plantilla para reportar bug nuevo

Copia esta plantilla y llena los campos cada vez que encuentres un bug:

```markdown
## BUG-XXX: [Título corto descriptivo]

**Fecha**: ____ / ____ / ______
**Reportado por**: _____________________
**Cliente**: _____________________
**Versión del software**: _____________________
**Severidad**: 🟢 Baja / 🟡 Media / 🟠 Alta / 🔴 Crítica
**Estado**: 🆕 Nuevo / 🔧 En revisión / ✅ Resuelto / ❌ No reproducible

### Comando ejecutado
\`\`\`cmd
python cli.py ...
\`\`\`

### Lo que esperaba
_(qué debería haber pasado)_

### Lo que pasó
_(qué pasó realmente)_

### Mensaje de error (si aplica)
\`\`\`
[pegar mensaje completo]
\`\`\`

### Pasos para reproducir
1. ...
2. ...
3. ...

### Archivos involucrados
- _(rutas relativas a la carpeta del proyecto)_

### Captura de pantalla / evidencia
_(opcional, pegar imagen o describir)_

### Notas adicionales
_(cualquier dato extra)_

### Resolución
_(se llena cuando el bug se arregla)_
- **Fecha resolución**: ____ / ____ / ______
- **Commit / cambio**: _____________________
- **Test agregado**: _____________________
```

---

## Categorías de severidad

| Severidad | Definición | Ejemplo |
|---|---|---|
| 🔴 **Crítica** | Bloquea totalmente la entrega al cliente. Datos corruptos o pérdida de información. | El generador crashea con un cliente válido. Los datos del cliente A aparecen en el SG-SST del cliente B. |
| 🟠 **Alta** | Impide entregar correctamente. Un componente importante falla. | La matriz de peligros sale vacía. Los links del listado maestro apuntan al documento incorrecto. |
| 🟡 **Media** | Funcionalidad parcial. Requiere workaround. | Una columna del Excel sale con formato incorrecto. Un texto humanizado tiene errores. |
| 🟢 **Baja** | Cosmético o de bajo impacto. | Una celda con texto blanco. Falta una tilde. Una banda con altura incorrecta. |

---

## Bugs reportados

### BUG-001: Codigos FMT en matriz no matchean con archivos FTM en Drive (RESUELTO)

**Fecha**: 2026-05-04
**Reportado por**: bts19
**Cliente**: FUNDACIÓN CRISTOVISIÓN
**Severidad**: 🟠 Alta
**Estado**: ✅ Resuelto

**Lo que pasaba**: Los códigos `SST-FMT-010` al `SST-FMT-014` de la matriz quedaban sin link, aunque los archivos `SST-FTM-010` a `SST-FTM-014` estaban en Drive.

**Resolución**:
- Agregada equivalencia `FMT ↔ FTM` en `engine/listado_maestro.py::_PREFIJOS_EQUIVALENTES`
- Test agregado: `test_listado_maestro.py::TestNormalizarCodigo::test_ftm_se_convierte_a_fmt`

---

### BUG-002: Links incorrectos por código duplicado SST-RGL-001 (RESUELTO)

**Fecha**: 2026-05-03
**Severidad**: 🟠 Alta
**Estado**: ✅ Resuelto

**Lo que pasaba**: El código `SST-RGL-001` se usa para 2 documentos distintos (Reglamento COPASST y Reglamento de Higiene). Ambas filas de la matriz apuntaban al mismo archivo.

**Resolución**:
- `_indexar_drive_links_por_codigo()` ahora devuelve una LISTA de candidatos por código
- `_elegir_mejor_match()` desambigua por nombre usando similitud Jaccard
- Test: `test_listado_maestro.py::TestElegirMejorMatch::test_codigos_duplicados_desambigua_por_nombre`

---

### BUG-003: PROGRAMA DE ORDEN Y ASEO no matchea (números distintos) (RESUELTO)

**Fecha**: 2026-05-07
**Severidad**: 🟡 Media
**Estado**: ✅ Resuelto

**Lo que pasaba**: La matriz pedía `SST-PRG-005 PROGRAMA DE ORDEN Y ASEO`, pero el archivo en Drive era `SST-PRG-006 PROGRAMA DE ORDEN Y ASEO`. Mismo nombre, distinto número.

**Resolución**:
- Agregado fallback por prefijo + nombre cuando el código exacto no coincide
- Si la similitud por nombre supera 0.75 con un archivo del mismo prefijo, se asigna
- Filtro adicional por carpeta/numeral de la sección

---

### BUG-004: Typo "ADQUISICION" vs "ADQUISISION" (RESUELTO)

**Fecha**: 2026-05-07
**Severidad**: 🟡 Media
**Estado**: ✅ Resuelto

**Lo que pasaba**: Archivos en Drive con typos en su nombre no matcheaban con la matriz (ej: archivo `POLITICA DE ADQUISISION` vs matriz `POLITICA DE ADQUISICION`).

**Resolución**:
- Implementada función `_similitud_palabras()` usando `SequenceMatcher` con umbral 0.85
- Tolera typos de 1-3 caracteres en palabras de la misma longitud

---

### BUG-005: Link asignado a documento incorrecto (sin validar nombre) (RESUELTO)

**Fecha**: 2026-05-07
**Severidad**: 🟠 Alta
**Estado**: ✅ Resuelto

**Lo que pasaba**: La matriz pedía `SST-FMT-053 FORMATO DE CONTROL DE PLAGAS` pero el único archivo SST-FMT-053 en Drive era `MANTENIMIENTO DE TANQUES`. El sistema lo asignaba igual.

**Resolución**:
- Aún con un solo candidato, ahora se valida que `similitud(esperado, candidato) >= 0.40`
- Si no supera el umbral, no se asigna link (mejor vacío que incorrecto)

---

### BUG-006: Hoja SEGURIDAD VIAL se eliminaba con PESV (RESUELTO)

**Fecha**: 2026-05-08
**Severidad**: 🟡 Media
**Estado**: ✅ Resuelto

**Lo que pasaba**: Al activar `generar_pesv: false`, se eliminaba también la hoja "SEGURIDAD VIAL" de SST-MTZ-006, que aplica para TODAS las empresas (Resolución 1565/2014).

**Resolución**:
- `engine/pesv_cleanup.py::_es_hoja_pesv()` excluye explícitamente "SEGURIDAD VIAL"
- Solo elimina hojas con "PESV" en el nombre

---

### BUG-008: [Content_Types].xml mal armado bloquea correcciones legales (RESUELTO)

**Fecha**: 2026-05-11
**Reportado por**: bts19 (cliente URBANISMO Y CONSTRUCCIONES MUNAR S.A.S)
**Severidad**: 🟡 Media
**Estado**: ✅ Resuelto

**Lo que pasaba**: Algunos archivos `.docx` de la plantilla tenían referencias a imágenes de marca de agua (`/word/media/watermark_*.png`) pero el `[Content_Types].xml` interno del Word no declaraba las extensiones. Al intentar abrirlos con `python-docx` para aplicar correcciones legales, salía:

```
"no content type for partname '/word/media/watermark_rIdWM1.png' in [Content_Types].xml"
```

Estos archivos no recibían las correcciones legales (Decreto viejo → Decreto nuevo, Resolución 652 → 3461, etc.).

**Archivos afectados** (al menos):
- SST-PRO-009 PROCEDIMIENTO DE ADQUISICIÓN DE BIENES Y SERVICIOS
- SST-FMT-015 ACTA DE REUNIÓN
- SST-FMT-016 INFORME TRIMESTRAL COPASST
- SST-MNL-001 MANUAL COMITE DE CONVIVENCIA LABORAL
- SST-FMT-017 CONVOCATORIA PARA EL CCL
- SST-FMT-018 INSCRIPCIONES CCL
- SST-FMT-021 CONFORMACIÓN DEL CCL
- SST-FMT-023 ACTA DE REUNIÓN CCL
- SST-FMT-024 PRESENTACIÓN DE SOLICITUDES O QUEJAS CCL

**Resolución**:
- Función `_reparar_content_types_si_falta_watermark()` en `engine/normativa.py`
- Detecta extensiones de imágenes en `word/media/` no declaradas
- Inyecta las declaraciones `<Default Extension="..." ContentType="..."/>` faltantes
- Se invoca automáticamente cuando python-docx falla con error de content type

---

### BUG-007: "Auxiliar de Reparto Motorizado" se clasificaba como peatón duplicado (RESUELTO)

**Fecha**: 2026-05-08
**Severidad**: 🟡 Media
**Estado**: ✅ Resuelto

**Lo que pasaba**: En la matriz vial, un cargo "Auxiliar de Reparto Motorizado" aparecía como motociclista Y como peatón (porque la heurística capturaba la palabra "auxiliar").

**Resolución**:
- Reescrita la lógica de `_inferir_actores_viales()` en `engine/matriz_peligros_viales.py`
- Cargos motorizados se excluyen explícitamente de la clasificación peatón
- Peatón solo se detecta con señales claras (diligencias, trámites bancarios, mensajería física)

---

## Cómo se convierte un bug en test

Cuando resuelvas un bug:

1. Identifica el caso que falló
2. Escribe un test que **reproduzca el bug** (debe fallar antes del fix)
3. Aplica el fix
4. Verifica que el test ahora pasa
5. Agrega el test al archivo correspondiente en `tests/`
6. Marca el bug como resuelto en este archivo con referencia al test

Ejemplo:

```python
# tests/test_listado_maestro.py

def test_ftm_se_convierte_a_fmt(self):
    """BUG-001: Archivos con FTM no matcheaban con codigos FMT.

    Solucion: agregar FTM como equivalente de FMT.
    """
    assert _normalizar_codigo("SST-FTM-010") == "SST-FMT-010"
```

Esto garantiza que el bug **no vuelva a aparecer** silenciosamente.

---

## Histórico de cierre de bugs

| Bug | Reportado | Resuelto | Días abierto |
|---|---|---|---|
| BUG-001 FMT/FTM | 2026-05-04 | 2026-05-04 | < 1 |
| BUG-002 Códigos duplicados | 2026-05-03 | 2026-05-03 | < 1 |
| BUG-003 Números distintos | 2026-05-07 | 2026-05-07 | < 1 |
| BUG-004 Typos ortográficos | 2026-05-07 | 2026-05-07 | < 1 |
| BUG-005 Validación de nombre | 2026-05-07 | 2026-05-07 | < 1 |
| BUG-006 SEGURIDAD VIAL borrada | 2026-05-08 | 2026-05-08 | < 1 |
| BUG-007 Auxiliar Motorizado | 2026-05-08 | 2026-05-08 | < 1 |
