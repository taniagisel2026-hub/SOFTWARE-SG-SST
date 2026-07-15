# Guía del Consultor SST

Esta guía te explica **paso a paso** cómo generar el SG-SST completo de un cliente con la herramienta, ya sea usando el portal web o editando el YAML a mano. No necesitas saber programar.

---

## Tabla de contenido

- [¿Portal web o edición manual del YAML?](#portal-web-o-edicion-manual-del-yaml)

1. [Antes de empezar — qué necesitas](#1-antes-de-empezar)
2. [Preparar el YAML del cliente](#2-preparar-el-yaml-del-cliente)
3. [Preparar los manuales de funciones (opcional pero recomendado)](#3-preparar-los-manuales-de-funciones)
4. [Generar el SG-SST](#4-generar-el-sg-sst)
5. [Subir a Google Drive del cliente](#5-subir-a-google-drive-del-cliente)
6. [Verificar que todo quedó bien](#6-verificar-que-todo-quedo-bien)
7. [Casos especiales: PESV, COPASST vs Vigía](#7-casos-especiales)
8. [Solución de problemas comunes](#8-solucion-de-problemas-comunes)

---

## ¿Portal web o edición manual del YAML?

Desde la Fase 5 hay **dos caminos** para llegar al mismo archivo de configuración del cliente — el resto de esta guía (secciones 2 a 8) sirve para ambos, porque los dos terminan en el mismo YAML.

**Opción A — Portal web (más rápido, menos errores de transcripción)**

Si el administrador ya tiene el sistema corriendo, te dará un link parecido a `http://<servidor>/portal/`. Ahí:

1. Creas un borrador nuevo.
2. Módulo por módulo, subes los documentos y llenas los datos que pide cada uno (8 en total: documentos legales, riesgo/ARL, sedes, responsable SST y consultora, fechas/versión/opciones, descripción de la empresa, cargos, y estilo del documento).
3. En el módulo 1 subes el RUT, la Cámara de Comercio y la Cédula del representante legal (PDF) + el logo y la firma (imágenes) — el sistema **extrae los datos solo** y te los muestra para que confirmes o corrijas. Si el RUT y la Cámara de Comercio traen una dirección distinta, te lo señala como conflicto para que elijas cuál es la correcta en vez de adivinar.
4. No puedes pasar al módulo siguiente sin confirmar el actual, así no se te olvida ningún dato.
5. Al terminar los 8 módulos, el botón **"Finalizar"** te entrega el YAML ya armado (el mismo resultado que llenar la sección 2 a mano) — con eso puedes saltar directo a la [sección 4 (Generar el SG-SST)](#4-generar-el-sg-sst).

**Opción B — Editar el YAML a mano**

Si no tienes acceso al portal, o prefieres el control manual campo por campo, sigue con la sección 2 de esta guía tal como está.

> El portal solo arma el archivo de configuración. La generación real (`python cli.py generar ...`), la subida a Drive y la verificación de links siguen siendo exactamente los mismos pasos sin importar cuál camino usaste para llegar al YAML.

---

## 1. Antes de empezar

Necesitas tener en mano del cliente:

- ✅ **RUT** o **Cámara de Comercio** (para obtener: razón social, NIT, dirección, actividad económica CIIU, representante legal)
- ✅ **Email y teléfono de contacto**
- ✅ **Número de trabajadores actual**
- ✅ **ARL** del cliente y **nivel de riesgo** (I, II, III, IV o V)
- ✅ **Logo de la empresa** en PNG o JPG (preferible fondo transparente)
- ✅ **Firma del Representante Legal** y **firma del Profesional SST** (PNG con fondo transparente)
- ✅ **Color corporativo** preferido (azul, verde, morado, naranja, gris, o cualquier otro en código hex)
- ✅ **Manuales de funciones** de cada cargo en `.docx` (opcional pero acelera mucho el proceso)
- ✅ Saber si el cliente **necesita PESV** (Plan Estratégico de Seguridad Vial — aplica si tiene 10+ vehículos o usa transporte como actividad principal)

---

## 2. Preparar el YAML del cliente

### Paso 2.1 — Copiar la plantilla

Abre la carpeta del software y entra a `clientes_config`. Verás un archivo `ejemplo_cliente.yaml`. **Copia y pega ese archivo** y renómbralo con el nombre corto del cliente:

```
clientes_config/CRISTOVISION.yaml
clientes_config/METALICAS_DEL_VALLE.yaml
```

> **Tip**: Usa solo letras mayúsculas, números y guion bajo. Sin tildes ni espacios en el nombre del archivo.

### Paso 2.2 — Editar el YAML

Abre el archivo con el Bloc de Notas o cualquier editor de texto y llena cada campo. Aquí está el detalle de cada sección:

#### Información de la empresa

```yaml
razon_social: "FUNDACIÓN CRISTOVISIÓN"        # Exacto como aparece en el RUT
nit: "830114869-4"                             # Con guion del dígito de verificación
direccion_empresa: "Calle 18 # 97-61"
telefono_empresa: "747 0710"
email_empresa: "talentohumano@cristovision.org"
sitio_web: "www.cristovision.org"              # Opcional, deja "" si no tiene
barrio_empresa: "Fontibón"
ciudad_empresa: "Bogotá D.C"
departamento_empresa: "Bogotá D.C"
```

#### Actividad económica

```yaml
actividad_economica: >
  Actividad Principal: 9491 (Actividades de organizaciones religiosas)
  Actividad Secundaria: 5911 (Producción de películas)
```

> Toma los códigos **CIIU** directamente del RUT del cliente. Si tiene varias actividades, sepáralas en líneas.

#### Seguridad social y trabajadores

```yaml
nivel_riesgo: "I"                              # I, II, III, IV o V según ARL
arl: "AXA COLPATRIA"
numero_trabajadores_empresa: "75"
color_corporativo: "azul"                      # azul, verde, morado, naranja, gris -- o un hex "#2E75B6"
```

#### Sedes

Una sola sede:
```yaml
sedes:
  - nombre: "Sede Principal"
    direccion: "Calle 18 # 97-61"
    ciudad: "Bogotá D.C"
    telefono: "747 0710"
```

Múltiples sedes:
```yaml
sedes:
  - nombre: "Sede Administrativa"
    direccion: "Calle 18 # 97-61"
    ciudad: "Bogotá D.C"
    telefono: "747 0710"
  - nombre: "Sede Producción"
    direccion: "Km 5 Vía Siberia"
    ciudad: "Cota, Cundinamarca"
    telefono: "601 876 5432"
```

#### Representante Legal y Responsable SST

```yaml
representante_legal: "Rolfi Rozo Villamizar"
cedula_rl: "79.123.456"
cargo_rl: "Gerente General"
ciudad_cedula_rl: "Bogotá D.C"

nombre_responsable_sst: "Pablo Alexander Medina Pedraza"
cedula_responsable_sst: "1.014.179.739"
cargo_sst: "Profesional SST"
ciudad_cedula_sst: "Bogotá D.C"
licencia_sst: "8866 del 30 de abril de 2021"
```

#### Imágenes (rutas a archivos)

```yaml
logo: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/Cristovisión.png"
firma_rl: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/firma_rl.png"
firma_sst: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/firma sst.jpg"
```

> **Importante**: usa **diagonales hacia adelante** `/` aunque sea Windows. Y comillas dobles `"..."`.

#### Documentos de soporte del responsable SST (PDF, opcionales)

```yaml
hoja_vida_sst: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/hoja_vida.pdf"
licencia_sst_pdf: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/licencia_sst.pdf"
cedula_sst_pdf: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/cedula_sst.pdf"
diploma_cursos_sst: "C:/Users/bts19/OneDrive/Desktop/Medina industrial/diploma.pdf"
```

Estos 4 se copian tal cual (sin procesar) a la carpeta `1.1.1 RESPONSABLE DEL SG-SST/2. DOCUMENTOS` del SG-SST generado, como soporte de que el responsable SST cumple los requisitos. En el asistente por módulos son obligatorios (Módulo 4); si se escribe el YAML a mano, son opcionales.

#### Fechas y versión

```yaml
version: "7"
fecha_emision: "20/04/2026"
fecha: "20 de abril del 2026"
año: "2026"
ciudad: "Bogotá D.C"
```

> **Importante**: el campo `año` no es solo decorativo. El sistema lo usa para decidir **qué correcciones legales aplicar** en el normograma (`engine/normativa.py`). Si el SG-SST es de un año anterior (por ejemplo, para reconstruir el sistema histórico de un cliente), solo se aplican las correcciones de normas que ya estaban vigentes en ese año — no se le inyectan normas que salieron después. Usa siempre el año real del SG-SST que estás generando, no el año actual.

#### Opciones de generación

```yaml
generar_pesv: false                            # true si aplica PESV
```

> Existe también `generar_normograma_actualizado` (revisa el normograma con Claude y sugiere normas adicionales según la actividad económica), pero el asistente y la plantilla `ejemplo_cliente.yaml` ya no lo usan por defecto — las correcciones legales de `engine/normativa.py` se mantienen actualizadas aparte y se aplican siempre, sin necesitar saldo de la API. Si quieres esa capa extra de todos modos, puedes agregar `generar_normograma_actualizado: true` al YAML a mano.

#### Estilo del documento (opcional)

```yaml
pie_pagina_numero: "izquierda"                 # izquierda | centro | derecha | ninguno
estilo_encabezado: "2"                         # "" (original) | "1" | "2" | "3" | "4"
tipo_letra: "Arial"                            # "" (Calibri original) | "Arial" | "Calibri" | "Times New Roman"
```

`pie_pagina_numero` cambia la posición del número de página en el pie de **todos** los documentos Word generados (o lo quita, con `"ninguno"`). Si se deja vacío, usa la posición que ya trae la plantilla (centrado).

`estilo_encabezado` reacomoda el logo, el título, el código, la versión y la fecha del encabezado en **todos** los documentos Word generados, en una de 4 grillas alternativas (nada se agrega ni se quita, solo cambia la ubicación). Si se deja vacío, usa el encabezado original de la plantilla.

`tipo_letra` cambia la fuente de **todos** los documentos Word generados (títulos, subtítulos y contenido a la vez, ya que la plantilla no usa Estilos de Word separados por nivel). Por ahora no afecta las matrices de Excel. Los tres viven en el Módulo 8 del asistente por módulos, con vista previa antes de confirmar.

#### Descripción para humanización con IA

```yaml
descripcion_empresa: >
  Empresa dedicada a la producción, realización y difusión de contenidos
  audiovisuales de carácter religioso, orientados a la evangelización,
  promoción de valores cristianos y fortalecimiento espiritual de la
  comunidad a través de medios de comunicación.
```

> Este texto se le pasa a Claude para que adapte los documentos al contexto real del cliente. Mientras más específico, mejor la humanización.

#### (Opcional) Cargos para Matriz de Peligros

**Opción A — Manuales de funciones (RECOMENDADO)**:

```yaml
manuales_cargos_dir: "C:/Users/bts19/Downloads/MANUALES CRISTOVISION"
```

**Opción B — Lista manual** (si no tienes los manuales):

```yaml
cargos:
  - nombre: "Director General"
    proceso: "ADMINISTRATIVO"
    zona: "Oficinas Administrativas"
    tipo: "administrativo"
    num_expuestos: 1
  - nombre: "Camarógrafo"
    proceso: "PRODUCCION Y CONTENIDO"
    zona: "Estudios de Grabación"
    tipo: "operativo"
    num_expuestos: 4
  - nombre: "Conductor"
    proceso: "LOGISTICA"
    tipo: "conductor"
    num_expuestos: 1
```

Tipos válidos: `administrativo`, `operativo`, `conductor`, `servicios_generales`.

---

## 3. Preparar los manuales de funciones

Si tienes los manuales en **Word, Excel o PDF**, ponlos todos en **una sola carpeta** (pueden estar mezclados):

```
C:/Users/bts19/Downloads/MANUALES CRISTOVISION/
├── 1. DIRECTOR GENERAL.docx
├── 2. SUBDIRECCIÓN DE PRODUCCIÓN.xlsx
├── 3. CAMARÓGRAFO.pdf
├── 4. EDITOR AUDIOVISUAL.docx
├── 5. CONDUCTOR.pdf
└── ...
```

### Formatos soportados

| Formato | Soportado | Notas |
|---|---|---|
| **`.docx`** (Word) | ✅ | Recomendado, lectura más confiable |
| **`.xlsx`** (Excel) | ✅ | Lee todas las hojas |
| **`.pdf`** (PDF) | ✅ | Funciona con PDFs de texto. PDFs escaneados requieren OCR previo |

### Reglas

- Un archivo = un cargo
- El sistema lee automáticamente el campo **DENOMINACIÓN** dentro del manual para identificar el cargo
- Si el manual no tiene DENOMINACIÓN, usa el nombre del archivo (sin la numeración inicial)
- Puedes mezclar formatos en la misma carpeta — el sistema los procesa todos
- Pueden ser los manuales que ya tienes — no requieren formato especial

### Lo que extrae el sistema

Por cada manual, lee:
- Nombre del cargo (DENOMINACIÓN)
- Dependencia
- Procesos
- Objetivo del cargo
- Funciones y responsabilidades específicas
- Toda la demás información

Y con eso, Claude genera:
- Las actividades resumidas para la matriz IPVR
- Las tareas en bullets
- Los peligros específicos que aplican a ese cargo
- La valoración GTC 45 (ND, NE, NC) según la exposición real

---

## 4. Generar el SG-SST

Abre **CMD** (Símbolo del sistema de Windows) y ejecuta:

```cmd
cd "C:\Users\bts19\OneDrive\Desktop\Medina industrial\SOFTWARE SG-SST"
python cli.py generar clientes_config\MI_CLIENTE.yaml
```

### Lo que verás en pantalla

```
=== GENERADOR SG-SST v8.5 ===

1. Cargando configuración del cliente...
   Empresa: FUNDACIÓN CRISTOVISIÓN
   NIT: 830114869-4
   Actividad: 9491 (Actividades de organizaciones religiosas)...
   Nivel riesgo: I | ARL: AXA COLPATRIA
   Trabajadores: 75

2. Preparando contexto de marcadores...
   125 variantes de marcadores generadas
   3 imagenes encontradas
   Humanizacion con Claude API: ACTIVADA

3. Procesando documentos Word...
   [████████░░░] 152/200 documentos

5. Revisando normograma con Claude...
6b. Generando Matriz de Peligros IPVR (GTC 45)...
   Matriz IPVR llenada para 8 cargo(s)

7. Generando reporte de auditoria...

=== GENERACIÓN COMPLETADA ===
   Word procesados:  187
   Excel procesados: 24
   Humanizados:      89
```

### Dónde queda el resultado

```
clientes/FUNDACIÓN_CRISTOVISIÓN/
├── I. PLANEAR/
├── II. HACER/
├── III. VERIFICAR/
├── IV. ACTUAR/
├── NORMAS/
├── PLAN ESTRATÉGICO DE SEGURIDAD VIAL/   (solo si generar_pesv: true)
└── REPORTE_AUDITORIA_SST.txt
```

### Tiempo estimado

- **Sin humanización**: 2–5 minutos
- **Con humanización completa**: 15–40 minutos (depende del número de documentos y cargos)

> Si solo quieres probar marcadores sin Claude:
> ```cmd
> python cli.py generar clientes_config\MI_CLIENTE.yaml --no-humanizar
> ```

---

## 5. Subir a Google Drive del cliente

```cmd
python cli.py subir-drive clientes_config\MI_CLIENTE.yaml
```

### Primera vez con un cliente

Se abrirá el navegador pidiendo permisos:

1. Inicia sesión con la cuenta del **cliente** (no la tuya)
2. Da permisos a la aplicación
3. La ventana se cerrará automáticamente

> **Importante**: cada cliente debe autorizar UNA VEZ desde su propia cuenta de Google. El token queda guardado en `tokens/token_<NIT>.json`.

### Si necesitas re-autorizar

```cmd
python cli.py auth-drive --cliente-id <NIT_sin_puntos_ni_guiones>
```

### Lo que pasa al subir

1. Crea la carpeta `<RAZON_SOCIAL>/` en el Drive del cliente
2. Replica toda la estructura PHVA
3. Sube cada documento
4. Inyecta `=HYPERLINK("url","texto")` en cada fila del listado maestro `SST-MTZ-005`
5. Genera pestañas adicionales por tipo de documento (PROGRAMA, FORMATO, MATRIZ, etc.)

---

## 6. Verificar que todo quedó bien

### Verificar los links del listado maestro

```cmd
python verificar_links_listado.py "clientes/FUNDACIÓN_CRISTOVISIÓN"
```

Te muestra fila por fila si:

| Estado | Significado |
|---|---|
| **OK** ✅ | El link apunta al documento correcto |
| **WARN** ⚠️ | Coincidencia parcial — revísalo manualmente |
| **FAIL** ❌ | El link apunta al documento incorrecto |
| **VACIO** | No se asignó link (puede ser certificado externo) |

### Revisar el reporte de auditoría

Abre `clientes/<RAZON_SOCIAL>/REPORTE_AUDITORIA_SST.txt`. Te dice:
- Documentos generados correctamente
- Marcadores que no se encontraron (campos faltantes en el YAML)
- Documentos con errores

### Abrir el listado maestro

`clientes/<RAZON_SOCIAL>/I. PLANEAR/2.5.1 ARCHIVO Y RETENCIÓN DOCUMENTAL/SST-MTZ-005 MATRIZ DE GESTIÓN DOCUMENTAL.xlsx`

Verifica que:
- ✅ Cada fila tenga el link a Drive en la columna LINK
- ✅ Las pestañas por tipo (PROGRAMA, FORMATO, MATRIZ...) tengan datos
- ✅ Si NO solicitaste PESV, la hoja PESV haya sido eliminada

---

## 7. Casos especiales

### PESV (Plan Estratégico de Seguridad Vial)

**Aplica si el cliente:**
- Tiene 10+ vehículos propios
- Su actividad principal involucra transporte
- Lo solicita explícitamente

**Cómo activar:**

```yaml
generar_pesv: true
```

Esto hará que:
- ✅ Se genere toda la carpeta `PLAN ESTRATÉGICO DE SEGURIDAD VIAL/`
- ✅ Se llene la matriz `PESV-MTZ-002 MATRIZ DE PELIGROS Y VALORACIÓN DE RIESGOS VIALES` según los cargos viales del cliente
- ✅ Se conserve la hoja `PESV` en el listado maestro

**Cómo desactivar:**

```yaml
generar_pesv: false
```

Esto hará que:
- ❌ Se elimine la carpeta PESV completa
- ❌ Se elimine la hoja PESV del listado maestro
- ❌ Se elimine la hoja PESV de la matriz de requisitos legales (SST-MTZ-006)
- ✅ La hoja **SEGURIDAD VIAL** sí se conserva (aplica para todas las empresas por Resolución 1565/2014)

### COPASST vs Vigía

El sistema **decide automáticamente** según el número de trabajadores:

| Trabajadores | Resultado |
|---|---|
| **≥ 10** | COPASST (Comité Paritario) |
| **< 10** | Vigía de Seguridad y Salud en el Trabajo |

Esto afecta:
- ✅ Textos dentro de los documentos (Claude los reescribe)
- ✅ Nombres de carpetas
- ✅ Nombres de archivos
- ✅ Humanización con IA

> Solo asegúrate de poner correctamente `numero_trabajadores_empresa` en el YAML.

### Color corporativo

```yaml
color_corporativo: "azul"   # azul, verde, morado, naranja, gris -- o un hex "#2E75B6"
```

Las **5 paletas predefinidas** son azul, verde, morado, naranja y gris (`engine/colores.py::PALETAS`). Si el color del cliente no es ninguna de esas, escribe su código **hex** directamente (ej. `"#C0392B"` para un rojo específico) — el sistema genera automáticamente las variantes claras/oscuras a partir de ese color. En el portal (Fase 5), el módulo 2 ya muestra estas 5 opciones como círculos de color para elegir con un clic, más un campo de hex para cualquier otro color.

El sistema cambia los colores de:
- Tablas y bordes en Word
- Encabezados y fondos en Excel
- Elementos de marca

Si no quieres cambiar nada, deja `color_corporativo: ""`.

---

## 8. Solución de problemas comunes

### "Archivo está siendo usado por otro proceso"

**Causa**: Excel/Word tiene abierto un archivo del cliente.

**Solución**: Cierra todos los Word y Excel y vuelve a correr.

### "ANTHROPIC_API_KEY no configurada"

**Causa**: Falta el archivo `.env` o la variable.

**Solución**: Crea `.env` en la raíz con:
```
ANTHROPIC_API_KEY=sk-ant-...
```

### "No se encontró credentials.json"

**Causa**: Falta el archivo de credenciales OAuth de Google.

**Solución**: Solicita el archivo al administrador y guárdalo en la raíz del proyecto.

### Algunos links del listado maestro están vacíos

**Causas posibles**:
1. El documento aún no se ha subido a Drive
2. El nombre del archivo en Drive es muy distinto al esperado en la matriz
3. El código fue modificado manualmente

**Solución**: Corre el verificador para diagnosticar:
```cmd
python verificar_links_listado.py "clientes/<RAZON_SOCIAL>"
```

Los `VACIO` legítimos son: certificados externos (afiliaciones, exámenes médicos), documentos que aporta el cliente.

### Un link apunta al documento incorrecto

**Causa**: Probablemente hay códigos duplicados o nombres muy similares.

**Solución**: El sistema usa el campo NOMBRE de la matriz para desambiguar. Si sigue saliendo mal:
1. Verifica que el nombre en la matriz coincida con el nombre del archivo
2. Verifica el numeral de la sección (ej: `4.1.2`) — el sistema busca el archivo dentro de la carpeta con ese numeral

### El logo del cliente no aparece

**Causas**:
1. La ruta en `logo:` está mal (verifica que el archivo exista)
2. La imagen es muy grande o pesada (>5MB)
3. El formato no es soportado (usar PNG o JPG)

**Solución**: Convierte el logo a PNG de máximo 1000px de ancho.

### Claude tarda demasiado o falla

**Causa**: Rate limit de Anthropic o problema de red.

**Solución**:
- Espera unos minutos y reintenta
- El sistema usa **caché**: lo ya humanizado no se vuelve a procesar
- Si quieres saltarte la humanización, usa `--no-humanizar`

### La matriz de peligros sale vacía

**Causa**: No definiste `cargos` ni `manuales_cargos_dir` en el YAML.

**Solución**: Agrega al YAML una de las dos opciones (ver sección 2).

### Marcadores sin reemplazar (aparecen `{{ALGO}}` en los documentos)

**Causa**: Falta un campo en el YAML del cliente.

**Solución**:
1. Revisa `REPORTE_AUDITORIA_SST.txt` para ver qué marcador faltó
2. Agrega ese campo al YAML
3. Vuelve a generar

Si quieres que falle inmediatamente si hay marcadores sin dato:
```cmd
python cli.py generar clientes_config\MI_CLIENTE.yaml --estricto
```

---

## Flujo completo resumido

```
1. Recibir RUT + Cámara Comercio + datos del cliente
                  ↓
2. Copiar ejemplo_cliente.yaml → MI_CLIENTE.yaml
                  ↓
3. Llenar todos los campos del YAML
                  ↓
4. (Opcional) Poner manuales de funciones en carpeta
                  ↓
5. python cli.py generar clientes_config\MI_CLIENTE.yaml
                  ↓
6. Revisar REPORTE_AUDITORIA_SST.txt
                  ↓
7. python cli.py subir-drive clientes_config\MI_CLIENTE.yaml
                  ↓
8. python verificar_links_listado.py "clientes/<RAZON_SOCIAL>"
                  ↓
9. Entregar al cliente: link a su carpeta de Drive
```

---

## ¿Y si necesito modificar manualmente alguna parte?

Sí puedes:

- **Editar un documento Word** después de generado: ábrelo, modifica, guarda. La próxima vez que generes para ese cliente, se sobrescribirá (a menos que cambies el nombre).
- **Agregar peligros adicionales a la matriz IPVR**: copia el formato de una fila existente y pégalo abajo.
- **Cambiar valores ND/NE/NC** en la matriz: edita las columnas correspondientes. Las fórmulas de Nivel de Riesgo e Interpretación se recalcularán automáticamente.

> **Nota**: Si re-generas el SG-SST después de modificar archivos manualmente, **se sobrescribirán todos**. Si quieres conservar cambios manuales, haz una copia de la carpeta antes de re-generar.

---

## Buenas prácticas

✅ Usa un YAML por cliente (no edites el mismo para varios)
✅ Guarda copia del YAML cuando termines, por si el cliente pide cambios después
✅ Verifica siempre el reporte de auditoría antes de entregar
✅ Verifica los links después de subir a Drive
✅ Pide al cliente que revise el documento maestro de SG-SST antes de la entrega final

❌ No edites la **Plantilla SG-SST 2026/** — es la base maestra
❌ No subas `credentials.json` ni `.env` a internet ni a repositorios públicos
❌ No borres la carpeta `tokens/` (perderías la autorización de Drive de los clientes)
