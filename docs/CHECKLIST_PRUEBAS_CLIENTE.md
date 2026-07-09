# Checklist de Pruebas con Cliente Real

Usa esta lista cada vez que generes un SG-SST para un cliente nuevo, para verificar que **todo quedó bien antes de entregar**.

Imprime o copia este checklist por cada cliente y ve marcando.

---

## Cliente: ________________________________  Fecha: ____ / ____ / ______

---

### A. Preparación

- [ ] Tengo el RUT y Cámara de Comercio del cliente
- [ ] Tengo el logo (PNG/JPG, máx. 1000px ancho)
- [ ] Tengo firmas del RL y Profesional SST (PNG con fondo transparente)
- [ ] Sé si el cliente necesita PESV (sí/no)
- [ ] Tengo los manuales de funciones (`.docx`) o la lista de cargos
- [ ] Tengo acceso a la cuenta Google Drive del cliente para autorizar OAuth

### B. YAML del cliente

- [ ] Creé `clientes_config/MI_CLIENTE.yaml` copiando `ejemplo_cliente.yaml`
- [ ] `razon_social` coincide EXACTAMENTE con el RUT
- [ ] `nit` con guion del dígito de verificación (ej: `900123456-7`)
- [ ] `actividad_economica` tiene código(s) CIIU del RUT
- [ ] `nivel_riesgo` correcto (I/II/III/IV/V según ARL)
- [ ] `arl` y `numero_trabajadores_empresa` llenos
- [ ] `sedes` con al menos una sede
- [ ] Datos del Representante Legal completos
- [ ] Datos del Profesional SST completos (puede ser tu info)
- [ ] Rutas de `logo`, `firma_rl`, `firma_sst` apuntan a archivos que existen
- [ ] `generar_pesv` configurado correctamente (true/false)
- [ ] `color_corporativo` definido
- [ ] `descripcion_empresa` con texto que describa al cliente (Claude lo usa)
- [ ] (Opcional) `manuales_cargos_dir` apunta a carpeta con `.docx` de manuales
- [ ] (Opcional) Bloque `cargos:` con la lista si no hay manuales

### C. Generación

```cmd
python cli.py generar clientes_config\MI_CLIENTE.yaml
```

- [ ] El comando termina sin errores fatales
- [ ] En consola veo `=== GENERACIÓN COMPLETADA ===`
- [ ] La carpeta `clientes/<RAZON_SOCIAL>/` existe
- [ ] Estructura PHVA completa: `I. PLANEAR`, `II. HACER`, `III. VERIFICAR`, `IV. ACTUAR`, `NORMAS`
- [ ] Si PESV: carpeta `PLAN ESTRATÉGICO DE SEGURIDAD VIAL/` presente
- [ ] Si NO PESV: carpeta PESV NO existe

### D. Revisión del reporte de auditoría

Abrir `clientes/<RAZON_SOCIAL>/REPORTE_AUDITORIA_SST.txt`.

- [ ] No hay errores marcados como `[ERROR]`
- [ ] Si hay marcadores faltantes (`[FALTANTE]`), los agregué al YAML y re-generé
- [ ] Documentos procesados: número similar al total de la plantilla (~200)

### E. Revisión visual de documentos (muestreo)

Abre **al menos 3 documentos al azar** de tipos distintos:

- [ ] **Una política** (`I. PLANEAR/.../SST-POL-001.docx`): logo aparece, datos del cliente aparecen, no hay marcadores `{{...}}` sin reemplazar, color corporativo aplicado
- [ ] **Un procedimiento** (`SST-PRO-...docx`): texto humanizado coherente, COPASST/Vigía consistente con el N° de trabajadores
- [ ] **Un programa** (`SST-PRG-...docx`): se ve profesional, sin frases robóticas
- [ ] **Un manual** (`SST-MNL-...docx`): tablas con bordes, texto legible
- [ ] **Una matriz Excel** (cualquier `.xlsx`): colores corporativos, datos del cliente, sin marcadores

### F. Matriz de Peligros IPVR

Abrir `clientes/<RAZON_SOCIAL>/II. HACER/4.1.2 MATRIZ DE IPVR/SST-MTZ-012 MATRIZ DE IPVR.xlsx`

- [ ] Cargos del cliente aparecen
- [ ] Bloques agrupados por PROCESO (Administrativo juntos, Operativo juntos, etc.)
- [ ] Cada cargo tiene ~11 peligros base + extras según su tipo
- [ ] Columnas A-B (Proceso, Zona) en color **durazno**
- [ ] Columnas C-L con colores alternados azul/verde por cargo
- [ ] **PROCESO** y **CARGO** en negrilla en la primera fila de cada bloque
- [ ] **CLASIFICACIÓN**, **TAREA RUTINARIA**, **ACEPTABILIDAD DEL RIESGO** en negrilla
- [ ] Fórmulas calculadas en columnas NP, NR, Aceptabilidad
- [ ] No hay celdas con texto blanco/invisible

### G. Matriz Vial PESV (solo si `generar_pesv: true`)

Abrir `clientes/<RAZON_SOCIAL>/PLAN ESTRATÉGICO DE SEGURIDAD VIAL/FASE 1/6. CARACTERIZACIÓN.../PESV-MTZ-002....xlsx`

- [ ] Los 5 factores aparecen (HUMANO, MECÁNICO, INFRAESTRUCTURA VIAL, AMBIENTAL, SOCIAL)
- [ ] Cada factor tiene una **banda roja** con texto blanco
- [ ] Bloques por actor (CONDUCTOR/MOTOCICLISTA/PEATÓN/PASAJERO) según los cargos
- [ ] Cargos administrativos sin desplazamiento NO aparecen en la matriz vial
- [ ] Columna **CLASIFICACIÓN** (H) en color durazno
- [ ] Columna **INTERPRETACIÓN DEL NIVEL DE RIESGO** (R) con colores semáforo (rojo/naranja/verde) uniformes
- [ ] Texto en columna R en **negro** (incluso sobre rojo)
- [ ] Fuente **Arial 9** en todas las celdas
- [ ] Alturas uniformes (no hay filas comprimidas)
- [ ] **PROCESO, ACTORES VIALES, CLASIFICACIÓN, INTERPRETACIÓN** en negrilla
- [ ] TIPO DE DESPLAZAMIENTO **sin negrilla**

### H. Subida a Drive

```cmd
python cli.py subir-drive clientes_config\MI_CLIENTE.yaml
```

- [ ] Primera vez: autorización en navegador completada
- [ ] El comando termina con `Subida completada`
- [ ] En consola se ve número alto de archivos subidos
- [ ] En Drive del cliente apareció la carpeta `<RAZON_SOCIAL>/`
- [ ] Estructura PHVA completa replicada en Drive

### I. Verificación de links del listado maestro

```cmd
python verificar_links_listado.py "clientes/<RAZON_SOCIAL>"
```

- [ ] La mayoría de filas salen como `OK`
- [ ] Los `WARN` los revisé manualmente — son aceptables
- [ ] Los `FAIL` los investigué y son justificados (o los corregí)
- [ ] Los `VACIO` son certificados/documentos externos que aporta el cliente

### J. Validación del listado maestro completo

Abrir `clientes/<RAZON_SOCIAL>/I. PLANEAR/2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL/SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx`

- [ ] Hoja `LISTADO MAESTRO PHVA`: cada fila con código tiene link clicable
- [ ] Pestañas adicionales: `PROGRAMA`, `PLAN`, `PROCEDIMIENTO`, `REGLAMENTO`, `MANUAL`, `POLITICA`, `INDICADOR`, `MATRIZ`, `FORMATO`
- [ ] Si NO PESV: hoja `PESV` eliminada
- [ ] Si SÍ PESV: hoja `PESV` con datos
- [ ] Click en algunos links abre el documento correcto en Drive

### K. Verificación del normograma

- [ ] Abrir `clientes/<RAZON_SOCIAL>/NORMAS/` y revisar 2-3 archivos
- [ ] Normas derogadas reemplazadas o eliminadas
- [ ] Texto coherente y profesional
- [ ] El campo `año` del YAML es el año real del SG-SST (no el año actual por defecto) — las correcciones legales se filtran por ese campo
- [ ] Si el SG-SST es de un año anterior al actual: correr `python cli.py normas` y confirmar que las normas con `vigente_desde` posterior al `año` del cliente NO se aplicaron

### L. Backup y entrega

- [ ] Hice copia del YAML del cliente (lo guardé en `clientes_config/`)
- [ ] El cliente recibió: link a su carpeta Drive + listado maestro descargado
- [ ] Le expliqué cómo navegar el listado maestro (pestañas por tipo)
- [ ] Le indiqué qué documentos faltan que él debe aportar (certificados, exámenes médicos)

---

## ⚠️ Si algo falla durante las pruebas

Documenta el bug usando esta plantilla:

```markdown
### Bug encontrado

**Cliente**: ___________________________________
**Fecha**: ____ / ____ / ______
**Versión del software**: ___________

**Comando ejecutado**:
```cmd
python cli.py ...
```

**Lo que esperaba**:
_(describe brevemente)_

**Lo que pasó**:
_(describe lo que viste)_

**Mensaje de error** (si aplica):
```
[pegar mensaje completo aquí]
```

**Archivos involucrados**:
- `clientes/MI_CLIENTE/.../archivo.docx`
- _(otros)_

**Reproducible**: sí / no
**Severidad**: baja / media / alta / crítica
```

Guarda los bugs en `docs/BUGS.md` o repórtalos al equipo técnico.

---

## Tiempo estimado total

- Preparación: 30 min
- Generación: 5-40 min (depende de N° documentos + humanización)
- Subida a Drive: 5-15 min
- Revisión y validación: 30-60 min

**Total**: 1.5 - 3 horas por cliente nuevo (la primera vez). Con experiencia: 45 min - 1.5 h.
