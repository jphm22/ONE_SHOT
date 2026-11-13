---
title: "OneShot AutoCharting V6"
subtitle: "Manual técnico · Arquitectura, bibliotecas, módulos, motor de datos y diagnóstico"
author: "Documentación del proyecto"
date: "10 de noviembre de 2026"
lang: es-PE
---

> **Documento de referencia del código.** Elaborado mediante revisión de `Generador_PPT_V6_Loops_Paleta_S26002825.zip`. Describe **exactamente la V6 empaquetada**. Las modificaciones visuales realizadas después en la instalación del usuario pueden diferir. Las pruebas disponibles son principalmente *offline*: no equivalen a una certificación del proveedor COM/ADO con todos los MDD/DDF reales.

# 1. Visión general y límites

**Objetivo:** convertir definiciones de variables de un `.xlsx` (fila 5+) en dos gráficos editables por lámina, empleando el archivo `.mdd` para metadatos y el `.ddf` para resultados. La extracción se ejecuta con Dimensions **TOM/MDM** y con **ADODB** como mecanismo complementario, en especial para loops. Se conservan trazas intermedias auditables y se desacopla la renderización de las consultas.

**Principios de diseño identificados en el código:**

1. **Una fuente de configuración:** el Excel incluye tipo de gráfico y cruce en la columna S; no hay catálogo adicional obligatorio.
2. **Datos nativos:** el MDD y el DDF no se transforman a TXT. TOM puede exportar tablas calculadas temporalmente a HTML.
3. **Separación de responsabilidades:** CLI, lectura de Excel, metadata, acceso DDF, lógica de loops, resultados y PowerPoint son módulos distintos.
4. **Trazabilidad estadística:** no inferir conteos desde porcentajes redondeados; marcar bases no verificadas.
5. **No falsear distribuciones:** convertir apilados o donas a barras si las condiciones de representación no se cumplen.
6. **Reutilización:** serializar resultados a JSON para producir variantes visuales sin consultar el estudio de nuevo.

La salida `ppt.py` no clona la plantilla `S26002825`: **extrae y aplica la paleta de series**. El diseño generado es propio de V6 (relación 16:9, cabecera y dos paneles), no una copia íntegra del PPTX de referencia.

# 2. Arquitectura física

```text
Generador_PPT_V6/
├── generar_ppt.py                       [entrada principal]
├── generar_ppt_dimensions_v3.py         [wrapper heredado]
├── auditar_ppt.py                       [auditoría de PPTX sin Dimensions]
├── LEEME.md, CAMBIOS_V6.md             [documentación histórica del ZIP]
├── Auditoria_V4_Recibido.json           [evidencia histórica; no código]
├── oneshot/
│   ├── __init__.py                      [paquete Python]
│   ├── cli.py                           [argumentos, modos, logging, orquestación]
│   ├── excel_config.py                  [lectura OOXML XLSX, reglas de gráficos]
│   ├── metadata.py                      [lectura MDD y rutas indexadas]
│   ├── ado.py                           [conexión y consultas directas al DDF]
│   ├── loops.py                         [estadísticas por atributo de loop]
│   ├── extraccion.py                    [tabulación TOM, conciliación ADO]
│   ├── modelos.py                       [dataclasses Variable y Resultado]
│   ├── salidas.py                       [escritura/lectura JSON y CSV]
│   ├── paleta.py                        [colores S26002825 / V5]
│   ├── ppt.py                           [gráficos editables y paginación]
│   └── vendor/
│       ├── __init__.py                  [paquete interno]
│       └── motor_dimensions_tom.py      [motor TOM reutilizado]
└── pruebas/
    ├── test_offline.py                 [seis pruebas V5]
    └── test_loops_paleta.py            [cinco pruebas de V6]
```

# 3. Flujo de ejecución y datos

```text
[Excel .xlsx] → excel_config.cargar() → Variable[]
                                         │
[MDD .mdd] → metadata.MetadataMDD() ─────┤
[DDF .ddf] ─→ TOM.Document / mrDataFileDsc
                                         │
                        extraccion.extraer()
                            ├─ TOM: ejes + tablas + Populate
                            ├─ HTML semántico temporal
                            ├─ adaptar(): % / media / base TOM
                            ├─ ADO: comprobar n/N en variables simples
                            └─ loops.extraer_loop(): N por atributo
                                         │
                                   Resultado[]
                            ├─ salidas → JSON / CSV / log
                            └─ ppt.crear() → gráficos PowerPoint
```

**Orden relevante:** el código exporta primero el conjunto de tablas TOM a HTML; después abre ADO para verificar conteos y consultar loops, evitando consultas simultáneas durante `Populate/Export`. Si el acceso ADODB falla, se mantiene la salida TOM representable cuando existe.

**Modo `presentar`:** omite por completo MDM, TOM, ADODB y Excel; consume `Resultado[]` desde el JSON (`salidas.leer_json()`) y llama a `ppt.crear()`.

# 4. Inventario de scripts: utilidad, funciones y dependencias

Las bibliotecas marcadas como **estándar** vienen con Python. El import de `openpyxl` en el motor reutilizado hace que sea una dependencia instalada, aunque el lector principal del Excel use XML directamente.

| Archivo | Responsabilidad y funciones clave | Bibliotecas involucradas y motivo |
|:--|:--|:--|
| `generar_ppt.py` | Punto de entrada. Importa y ejecuta `oneshot.cli.main()`. | `sys` (estándar): código de salida del proceso. |
| `generar_ppt_dimensions_v3.py` | Alias de compatibilidad con comandos antiguos; delega al principal. | Import interno de `generar_ppt`. No implementa extracción independiente. |
| `oneshot/cli.py` | `argumentos()`, `main()`; valida rutas, inicia logging, coordina cuatro modos, guarda archivos. | `argparse` parámetros; `logging` terminal/log; `sys` stdout; `pathlib` rutas; `Counter` resumen de estados. |
| `oneshot/excel_config.py` | `leer_hoja()`, `cargar()`, `inferir()`. Lee celdas E/F/G/H/J/S, reconoce alias y detecta gráficos faltantes. | `zipfile`, `xml.etree.ElementTree`: leer XLSX OOXML sin Excel COM; `re` validar expresiones; `unicodedata` normalizar texto; `logging`, `pathlib`. |
| `oneshot/metadata.py` | `MetadataMDD`; indexa recursivamente `Variables`, `Fields`, `ClassFields`, elementos y rutas. `campos_grid()`, `categorias_campo()`, `etiqueta_item()`. | `pythoncom` y `win32com.client` (pywin32) para `MDM.Document`; `re` para códigos/rutas; `pathlib` archivos. |
| `oneshot/ado.py` | `ConexionADO` abre DDF vía proveedor; `consultar()` consulta `VDATA`; `respuestas()` normaliza códigos; `verificar_conteos()` cruza N con % TOM. | `win32com.client` para `ADODB.Connection`/`ADODB.Recordset`; `Counter` conteos; `math`, `re` validación estadística y sintáctica; `logging`, `pathlib`. |
| `oneshot/loops.py` | `es_loop()` detecta candidatos; `extraer_loop()` consulta cada ruta indexada y calcula n, N, %, series y errores por atributo. | `collections.Counter` frecuencias; `logging` incidencias; imports internos `ado` y `modelos`. |
| `oneshot/extraccion.py` | `extraer()` coordina TOM+MDM+ADO; `construir_solicitud()`, `adaptar()`, `reconciliar_ado()`, `base_verificada()` y `conteos_semanticos()`. | `logging` observabilidad; `math`, `re` procesamiento; `shutil`, `tempfile` HTML; `dataclasses` solicitud; `pathlib`; motor TOM interno. |
| `oneshot/vendor/motor_dimensions_tom.py` | Motor reutilizado: define banners, tablas, estadísticas, `Populate`, exportación HTML y parsing semántico. | `pywin32` acceso TOM; `openpyxl` (importado por el motor heredado); `html.parser`, `re`, `json`, `math`, `itertools`, `tempfile`, `shutil`, `logging` y otras estándar. |
| `oneshot/modelos.py` | `Variable`, `Resultado`, `renderizable()`, `a_dict()` y `desde_dict()`. Contrato de datos sin dependencias de PowerPoint/COM. | `dataclasses` (estándar): tipos y serialización estable. |
| `oneshot/salidas.py` | `guardar_json()`, `leer_json()`, `guardar_csv()`, `guardar_detalle()`. | `json` persistencia; `csv` auditoría; `pathlib` archivos. |
| `oneshot/paleta.py` | Constantes RGB y `obtener('s26002825'/'v5')`; paleta por defecto de S26002825. | Sin librerías externas: tuplas RGB puras. |
| `oneshot/ppt.py` | `grafico()`, `panel()`, `paginar_loops()`, `crear()`, `tabla()`, auxiliares de texto y formas. | `python-pptx` gráficos/OOXML editables; `copy.deepcopy` paginación sin modificar el original; `pathlib` salida. |
| `auditar_ppt.py` | Inspecciona una presentación generada: diapositivas, gráficos, tablas, `N=0` y literales `None`. | `python-pptx` lectura; `argparse` CLI; `json` salida; `re` búsqueda; `pathlib` rutas. |
| `pruebas/test_offline.py` | Pruebas de porcentajes, ausencia de bases falsas, gráficos vacíos y JSON. | `unittest`/`tempfile` estándar; `python-pptx` inspección estructural. |
| `pruebas/test_loops_paleta.py` | Pruebas de rutas indexadas, loops RM, bases diferentes, paginación y colores exactos. | `unittest`/`tempfile` estándar; `python-pptx`; módulos internos simulados. |
| `oneshot/__init__.py`, `oneshot/vendor/__init__.py` | Identifican paquetes importables. | Sin lógica de negocio. |

**Dependencias externas realmente necesarias para la ejecución de la V6 empaquetada:**

```powershell
python -m pip install pywin32 python-pptx openpyxl
```

`openpyxl` se importa al cargar `oneshot/vendor/motor_dimensions_tom.py`, aunque el módulo `oneshot/excel_config.py` no lo requiera. **No es necesario instalar manualmente** `argparse`, `csv`, `json`, `logging`, `pathlib`, `xml.etree`, `zipfile`, `collections`, `dataclasses`, `unittest` o `tempfile`: son parte de la biblioteca estándar de Python.

**Software externo no instalable solo con pip:** Dimensions con las interfaces COM `TOM.Document` y `MDM.Document`, y proveedor ADO `mrOleDB.Provider.2` asociado a `mrDataFileDsc`. El registro COM y la compatibilidad de arquitectura deben resolverse en Windows. `python-pptx` no ejecuta Microsoft PowerPoint para generar el archivo.

# 5. Configuración, CLI y reglas de selección

## 5.1. Modos y opciones

`cli.py` reconoce `todo`, `validar`, `extraer` y `presentar`; `todo` se usa si no se especifica ninguna acción. `--solo-validar-excel` es alias de `validar`; `--exigir-catalogo` se acepta como compatibilidad pero **se ignora**.

| Opción | Tipo / valor por defecto | Uso exacto en V6 |
|:--|:--|:--|
| `--excel` | Path requerido excepto `presentar` | Archivo XLSX de configuración. |
| `--mdd`, `--ddf` | Paths nativos requeridos para `todo`/`extraer` | Datos de Dimensions. Se exige extensión `.mdd`/`.ddf`. |
| `--salida` | `Presentacion_Automatica.pptx` | Ruta del PPTX y raíz de archivos derivados. |
| `--datos` | Derivado de `--salida` | JSON de entrada al modo `presentar`. |
| `--hoja` | Primera hoja | Nombre exacto de hoja Excel. |
| `--desde` | `5` | Número de fila inicial Excel. |
| `--hasta` | Sin límite | Número de fila final Excel. |
| `--idioma` | `AUTO` | No sobrescribir `TOM.Document.Language`, salvo idioma explícito válido. |
| `--libreria` | `mrDataFileDsc` | Nombre de la conexión de datos Dimensions. |
| `--variable-total` | `TOTAL_BASE` | Variable de eje TOP cuando no hay cruce. |
| `--codigo-total` | `_1` | Categoría del TOP total, no un N fijo. |
| `--ponderador` | Vacío | Variable de ponderación TOM; ADO se omite cuando hay ponderador. |
| `--decimales` | `3` | Precisión configurada al tabular TOM. |
| `--max-categorias` | `12` | Recorta el top visual de algunas variables simples categóricas. |
| `--max-atributos` | `0` | 0=consultar todos los atributos de loop; positivo=limitar extracción. |
| `--atributos-por-grafico` | `8` | Cuántos atributos caben en un panel; pagina sin truncar la extracción. |
| `--paleta` | `s26002825` | `s26002825` o `v5`. |
| `--max-cruce` | `26` | Máximo de categorías en banner de cruce no-grid. |
| `--guardar-html` | `False` | Guardar informe de tablas TOM en `.dimensions.htm`. |
| `--sin-ado` | `False` | Desactivar respaldo directo DDF (incluyendo recuperación de loops). |
| `--incluir-errores` | `False` | Dibujar paneles para variables con estado no renderizable. |

## 5.2. Formato y límites de Excel

`excel_config.cargar()` lee E=tipo, F=estructura, G=variable, H=etiqueta, J=escala y S=instrucción; omite filas sin G. El tipo se toma de S, por ejemplo `barras | F1`; si está vacío, aplica inferencia automática. Si S contiene texto no reconocido, se registra advertencia y el gráfico queda inválido para extracción.

La heurística especial `PE26_*` se activa cuando en el Excel **completo** existen `CUOTA_CIUDAD`, `B8a` y `TOTAL_BASE`; no es una regla universal para otros estudios. Para estudios nuevos, ajustar `inferir()` en `excel_config.py` o completar la columna S.

`--desde` y `--hasta` seleccionan **filas del Excel**, no los números de diapositiva. `--hoja` sin indicar usa la primera hoja del libro. La lectura de XLSX es directa de OOXML (`ZipFile` + XML), no modifica el Excel, y no usa el libro como una base de datos SQL.

# 6. Motor MDD: indexación y validación de rutas

`metadata.MetadataMDD.__init__()` abre el MDD con `MDM.Document`, y `_indexar()` recorre, hasta profundidad configurada, `Variables`, `Fields`, `ClassFields` y elementos de iteradores. Guarda mapeos por `Name` y `FullName` para localizar variables incluso cuando sus subcampos no son entradas raíz.

Métodos esenciales:

- `get(nombre)` recupera el objeto metadata o `None`.
- `etiqueta(nombre, fallback)` obtiene `Label` cuando es legible.
- `categorias(nombre)` produce pares `(código, etiqueta)` desde categorías declaradas.
- `tiene_rp(nombre)` identifica si existe un subcampo de respuesta `Rp`.
- `campos_grid(nombre)` encuentra rutas indexadas del loop, preferentemente respuestas `*.Rp`.
- `categorias_campo(raiz, campo)` consigue el diccionario de respuestas del subcampo desde el MDD.
- `etiqueta_item(raiz, campo)` recupera la etiqueta del atributo iterado.
- `cerrar()` libera el documento MDD.

**Ejemplo de ruta nativa admisible:**

```text
B1[{_1}].Rp
B8a[{_3}].Rp
```

La expresión debe coincidir con la gramática `metadata.RUTA` antes de enviarse a ADO. No se ejecutan cláusulas SQL arbitrarias introducidas en el Excel. La generación de rutas de respaldo usa **códigos obtenidos del propio MDD**, no un rango inventado de índices.

# 7. Motor TOM: construcción, población y exportación

`extraccion.extraer()` crea una única conexión TOM con `motor.crear_documento_tom(mdd, ddf, libreria, idioma)`. Internamente la carga usa:

```python
# Idea central del motor: el segundo argumento se omite de verdad.
documento.DataSet.Load(mdd, pythoncom.Missing, ddf, 'mrDataFileDsc')
```

`pythoncom.Missing` es importante: pasar `None` no equivale al argumento omitido en determinadas interfaces COM de Dimensions. `idioma='AUTO'` conserva el idioma del documento; forzar `ESM` puede lanzar COM `-2147024809` si no es compatible.

Para cada variable, `construir_solicitud()` configura el TOP (`TOTAL_BASE` o variable de cruce), el SIDE (variable o `.Rp` del grid), un `Banner` y un nombre técnico de tabla `PPT_####`.

En el motor reutilizado se ejecutan, a grandes rasgos:

```text
Axes.AddNew(BAN_####)
Axes[BAN_####].Specification = ...
Tables.AddNew(PPT_####, SIDE * axis(BAN_####), título)
Tables[PPT_####].CellItems ...
documento.Populate('PPT_####')
```

El proceso exporta las tablas con el exportador `mrHtmlExport` a un `TemporaryDirectory`, analiza el HTML mediante marcadores semánticos y lo asocia a cada tabla preferentemente por ancla exacta. Si TOM omite anclas, una asociación por orden se permite solo cuando existe correspondencia uno a uno; de otro modo registra `ERROR_HTML`.

**La finalización de `Populate` no implica validación de la base N**. `adaptar()` interpreta porcentajes y medias, marca `OK_SIN_BASE` si no pudo comprobar N, y evita publicar conteos que no concuerdan con el DDF.

# 8. Motor ADO: verificación directa de respuestas

`ConexionADO` abre `ADODB.Connection` con proveedor `mrOleDB.Provider.2` y consulta la tabla virtual `VDATA`. Por cada campo indexado emite una consulta controlada:

```sql
SELECT B1[{_1}].Rp FROM VDATA
```

`respuestas()` admite códigos individuales o múltiples, pero **no convierte texto libre arbitrario en categorías**. Los códigos deben tener respaldo en el diccionario MDD. Se descartan valores vacíos o no reconocidos.

`verificar_conteos()` calcula conteos directos del DDF y compara contra los porcentajes TOM. Para considerarlo un respaldo aceptable exige que al menos el **60 %** de las categorías mostradas emparejen por etiqueta y que la diferencia máxima de porcentajes no supere la tolerancia configurada (`0.6` puntos porcentuales por defecto). Si falla, deja registro `[ADO RECHAZADO]` y **no sustituye** resultados TOM con cifras potencialmente incompatibles.

**Importante con ponderación:** `extraccion.py` no inicia ADODB cuando `--ponderador` tiene valor. Eso evita comparar porcentajes ponderados con conteos no ponderados; como contrapartida, podría dejar N sin respaldo y restringir recuperación de loops. La decisión está en el condicional `if usar_ado and not ponderador:`.

# 9. Loops y tratamiento estadístico

`loops.es_loop()` detecta la estructura candidata por los términos `grid`, `iterator` o `loop` en la columna F. `extraccion.extraer()` exige, además, confirmación del MDD mediante `tiene_rp()` o `campos_grid()`. Esta comprobación protege variables simples recodificadas que fueron etiquetadas como grids en el Excel.

**Algoritmo resumido de `extraer_loop()`**:

1. Obtener rutas con `metadata.campos_grid()`, por ejemplo `B1[{_1}].Rp`.
2. Obtener opciones de respuesta declaradas en MDD para cada ruta.
3. Ejecutar una consulta ADO por atributo; convertir respuestas válidas a códigos.
4. Calcular base válida de cada atributo: `N_i = casos con al menos una respuesta reconocida`.
5. Contar cada código una vez por caso y calcular `porcentaje_ij = 100 * n_ij / N_i`.
6. Conservar `bases_por_item` y `detalles_loop` con ruta, etiqueta, respuesta, n y porcentaje.
7. Distinguir RU/RM por atributo a partir de respuestas observadas con varios códigos.
8. Si es RM o las proporciones no constituyen una distribución completa, pasar de `apiladas`/`escala` a `barras` para evitar un 100 % engañoso.
9. Guardar los atributos completos en JSON y CSV; dividirlos visualmente después en `ppt.paginar_loops()`.

Ejemplo matemático: si el atributo A tiene N=100 y 45 selecciones de `_1`, su porcentaje es **45 %**. Si B tiene N=80 y 20 selecciones, su porcentaje es **25 %**. No se debe aplicar la base 100 a ambos atributos. En RM la suma de categorías puede exceder 100 % sin error estadístico.

`--max-atributos 0` consulta todos los atributos; `--atributos-por-grafico 8` pagina de ocho en ocho. Si hay 13 atributos, se generan dos paneles (8+5), que podrían ocupar una misma diapositiva. **La paginación no elimina los detalles exportados**. Si un subcampo falla se deja constancia, sin rellenarlo con N=0.

Cuando se detectan más de 12 opciones para una visualización no apilada de loop, el código puede mostrar las 12 series principales, **manteniendo las opciones completas en el CSV**.

# 10. Contrato de datos y serialización

`modelos.Variable` transporta `fila`, `nombre`, `titulo`, `tipo`, `estructura`, `escala`, `instruccion`, `grafico`, `cruce`, `inferido`, `advertencias`.

`modelos.Resultado` transporta `variable`, `estado`, `tipo`, `titulo_mdd`, `respuesta`, `base`, `base_fuente`, `registros`, `categorias`, `series`, `media`, `bases_por_item`, `detalles_loop`, `nota`, `metodo`.

Campos importantes:

| Campo | Qué representa |
|:--|:--|
| `registros[]` | Resultado simple por categoría: etiqueta, código, porcentaje y conteo n disponible. |
| `series[]` | Una o varias series sobre `categorias[]`; utilizable por el motor gráfico. |
| `base` | N global de la variable cuando es defendible; `null` si no se verificó o difiere por atributo. |
| `base_fuente` | Origen explícito (`TOM: real`, `ADODB ...`, etc.). |
| `bases_por_item[]` | N de cada atributo de loop. |
| `detalles_loop[]` | Registro completo de atributo × respuesta. |
| `estado` | Resultado de la extracción y su aptitud para visualizar. |
| `nota`/`metodo` | Explicación metodológica y decisiones de respaldo. |

`Resultado.renderizable()` solo acepta estados `OK` y `OK_SIN_BASE`, con valores numéricos suficientes. El serializador `guardar_json()` escribe `version: 6`; `leer_json()` acepta versiones **5 y 6**. El CSV usa separador punto y coma y codificación `utf-8-sig`.

**Nombres de salida**: desde `--salida X.pptx`, la CLI genera `X.resultados.json`, `X.reporte.csv`, `X.detalle.csv`, `X.diagnostico.log` y, opcionalmente, `X.dimensions.htm`. En `presentar`, escribe PPTX y log, sin regenerar CSV/JSON. En `validar`, solo reporte CSV/log; no consulta MDD/DDF.

# 11. Motor PowerPoint y sus personalizaciones

`ppt.crear()` genera presentación 13.333 × 7.5 pulgadas; usa layout en blanco, agrega cabecera, divisor y paneles a izquierda/derecha. `panel()` incluye nombre, etiqueta, tipo, estructura, gráfico, frecuencias cuando estén disponibles, N y notas. `paginar_loops()` divide un resultado grande en paneles independientes y mantiene correspondencia entre categorías/series/bases.

`grafico()` utiliza `pptx.chart.data.CategoryChartData` y tipos como `BAR_CLUSTERED`, `COLUMN_CLUSTERED`, `DOUGHNUT`, `BAR_STACKED_100`, `LINE_MARKERS`, `RADAR`; `tabla()` utiliza una tabla nativa. Todos permanecen editables en PowerPoint. Se rellena cada serie con colores de `oneshot/paleta.py`, y en la dona se colorean los **puntos** de la serie, no solo la serie completa.

**Mapa de tipos y reglas especiales**:

| Tipo solicitado | Tipo PowerPoint habitual | Situaciones que lo modifican |
|:--|:--|:--|
| `barras` | `BAR_CLUSTERED` | Puede usarse como respaldo en varias condiciones. |
| `columnas` | `COLUMN_CLUSTERED` | Conserva series. |
| `dona` | `DOUGHNUT` | Si hay múltiples series/cruce, más de nueve categorías o suma no cercana a 100 %, puede convertirse en barras. |
| `apiladas`/`escala` | `BAR_STACKED_100` | Solo cuando los porcentajes son compatibles con una distribución completa; RM puede pasar a barras. |
| `promedio` | `COLUMN_CLUSTERED` | Una media numérica, formato no porcentual. |
| `linea` | `LINE_MARKERS` | Series y colores editables. |
| `radar` | `RADAR` | Series y colores editables. |
| `tabla` | Tabla nativa (máx. 15 registros en esta implementación) | No crea un objeto chart. |
| `funnel` | Barras como respaldo | No constituye un embudo PowerPoint especializado en esta versión. |

**Verificación de salida:** cuando `--incluir-errores` está desactivado, `crear()` cuenta los objetos chart/tabla por lámina; si no coincide con los paneles previstos, lanza una excepción y no guarda un PPTX incompleto. La última lámina puede tener un panel si el total es impar.

## 11.1. Paleta S26002825

Primeros colores de `oneshot/paleta.py`, en orden por serie/punto:

| Orden | Hex | Uso |
|:--:|:--|:--|
| 1 | `#1DAFAD` | Turquesa. |
| 2 | `#121B5A` | Azul marino. |
| 3 | `#E50158` | Magenta. |
| 4 | `#103C50` | Azul petróleo. |
| 5 | `#E2DA51` | Amarillo. |
| 6 | `#451B67` | Violeta. |
| 7 | `#2DB574` | Verde. |
| 8 | `#FF740F` | Naranja. |

El archivo contiene seis tonos adicionales para completar la rotación. No es una paleta inferida por el sistema operativo; proviene de los objetos de gráfico del PPTX de referencia.

## 11.2. Puntos de intervención visual

Todas estas personalizaciones pertenecen a `oneshot/ppt.py`, función `grafico()` (o a `paleta.py` para colores). No requieren cambiar `extraccion.py` ni el DDF:

```python
# Quitar cuadrícula de gráficos con ejes:
chart.value_axis.has_major_gridlines = False
chart.value_axis.has_minor_gridlines = False

# Ocultar escala vertical/horizontal de porcentaje (según tipo):
chart.value_axis.visible = False

# Etiquetas de porcentaje sin decimales:
d.number_format = '0"%"' if r.tipo != 'promedio' else '0.00'

# Etiquetas blancas únicamente si corresponde:
if tipo in ('apiladas', 'escala', 'dona'):
    d.font.color.rgb = rgb((255, 255, 255))
```

**Advertencias de ingeniería:** aplicar configuraciones `value_axis` solamente a gráficos que tienen ese eje; no forzar una `XL_LABEL_POSITION` no admitida por un tipo de gráfico; los segmentos claros (amarillo) pueden perder contraste con texto blanco. En la **V6 empaquetada**, `d.number_format` todavía es `0.0"%"`, `has_major_gridlines` está activado, y el orden automático de barras/etiquetas blancas **no está incorporado**. Son cambios locales posteriores que requieren pruebas de regresión.

**Ordenar barras horizontales:** antes de construir `CategoryChartData`, y solo si el tipo final sigue siendo `barras`, obtener una permutación de categorías y aplicarla a **todas** las series. El orden requerido en el arreglo puede ser inverso al visual por la orientación del eje de PowerPoint; verificar en un caso conocido (ejemplo B16/B17). No ordenar el CSV ni las respuestas del DDF.

**Mejora recomendada para soporte:** en `grafico()` hay bloques `try: ... except Exception: pass` que silencian errores de estilo. Sustituirlos por logs WARNING/ERROR durante el desarrollo permite diferenciar un problema de formato de uno de extracción. No se debe suprimir el error de creación del gráfico completo.

# 12. Diccionario de mensajes de terminal

El formato configurado en `cli.py` es `fecha-hora NIVEL [ETIQUETA] detalle`, enviado tanto a **stdout** como a `X.diagnostico.log`. En modo normal es nivel `INFO`; las excepciones individuales pueden registrar `ERROR` y continuar con otras variables.

| Mensaje | Emisor habitual | Significado / acción |
|:--|:--|:--|
| `[EXCEL]` | `excel_config` / `cli` | Cantidad de variables cargadas, inferidas, identificación heurística del estudio o ubicación del reporte de validar. |
| `[CONFIGURACIÓN]` | `cli` | Cantidad de variables del rango seleccionado y gráficos inválidos. |
| `[VISUAL]` | `cli` | Paleta, paneles y límite real de extracción de atributos. |
| `[COM]` | `motor_dimensions_tom` | Operación COM (Dispatch, carga de DDF, propiedades). Si se detiene aquí, revisar registro del componente y HRESULT. |
| `[IDIOMA]` | `motor_dimensions_tom` | Idioma activo `AUTO` u otra configuración. `AUTO` evita forzar un idioma no permitido. |
| `[MDD LOOPS]` | `extraccion` | Número de estructuras *confirmadas* como loops por el MDD. |
| `[VARIABLES]` / `[EXPRESION]` | motor TOM | SIDE y especificación que TOM utilizará para tabular. |
| `[PONDERACION]` | motor TOM | Si se aplicó una variable de peso o no. |
| `[VALORES]` / `[PROMEDIO]` | motor TOM | Modo de medida y medios de precisión. |
| `[TABLA]` | motor TOM | Declaración SIDE × TOP. |
| `[DIFSIG]` | motor reutilizado | Parámetros heredados del motor; en el flujo V6 principal no se activan diferencias estadísticas por defecto. |
| `[TOM OK]` | `extraccion` | `Populate` de esa variable terminó; faltan lectura semántica, conteos o reconciliación. |
| `[TOM ERROR]` | `extraccion` | Error durante construcción/población de la tabla; estado `ERROR_TOM`. |
| `[EXPORTADOR]` | motor TOM | Exportador `mrHtmlExport` localizado. |
| `[HTML]` | `extraccion` / motor TOM | Asociación de HTML a tablas; si se usó orden como respaldo, exige verificar correspondencia. |
| `[HTML ERROR]` | `extraccion` | No pudo interpretarse una tabla; estado `ERROR_HTML`. |
| `[DATOS]` | `extraccion` | Estado ya interpretado; N y cantidad de opciones. `base=None` no significa N=0. |
| `[ADO]` | `extraccion` | Conexión DDF activada o no disponible. |
| `[ADO RECHAZADO]` | `ado` | Conteos DDF no emparejan suficientemente con porcentajes TOM; cifras no se publican. |
| `[LOOP ITEM ERROR]` | `loops` | Falló un subcampo específico; la extracción puede continuar con otros atributos. |
| `[LOOP OK]` | `loops` | Atributos recuperados, opciones, rango N y número de fallos. |
| `[LOOP REEMPLAZADO]` | `extraccion` | El loop extraído por ADO sustituyó una matriz TOM previamente calculada. |
| `[LOOP]` | `extraccion` | ADODB no disponible: conserva TOM si representa datos válidos. |
| `[LOOP ERROR]` | `extraccion` | Excepción general en el loop; normalmente estado `ERROR_LOOP`. |
| `[ESTADOS]` | `cli` | Conteo de resultados por estado antes de construir PPTX. |
| `[RESULTADOS]` | `cli` | Rutas al JSON y reporte CSV generados. |
| `[PPTX]` | `cli` | Slides y paneles generados; elementos omitidos. |
| `[AVISO] --exigir-catalogo` | `cli` | Parámetro obsoleto que se ignora, no es un fallo. |

**Ejemplo orientativo, no una ejecución de los archivos privados reales:**

```text
2026-... INFO [EXCEL] 137 variables | 0 gráficos inferidos | estudio PE26=True
2026-... INFO [CONFIGURACIÓN] 137 variables, 0 inferidas | errores catálogo 0
2026-... INFO [VISUAL] paleta=s26002825, atributos por gráfico=8, límite extracción=TODOS
2026-... INFO [MDD LOOPS] 12 estructuras indexadas confirmadas
2026-... INFO [VARIABLES] PPT_0027 | medida=porcentajes | 1 rutas expandidas
2026-... INFO [EXPRESION] PPT_0027 = B1.Rp*axis(BAN_0027)
2026-... INFO [TOM OK] fila=27 variable=B1
2026-... INFO [ADO] Respaldo DDF habilitado para verificar conteos
2026-... INFO [LOOP OK] B1 | atributos=... / ... | opciones=... | bases=... | fallos=0
2026-... INFO [ESTADOS] {'OK': ...}
2026-... INFO [PPTX] ... diapositivas; ... paneles; ... errores omitidos
```

**Niveles de log:** `INFO` operación realizada, `WARNING` problema recuperable o revisión necesaria, `ERROR` fallo de componente/variable. Los `Traceback` muestran el archivo, función y línea. El primer `[... ERROR]` antes del traceback suele indicar dónde empezar a investigar.

# 13. Estados de procesamiento y comportamiento posterior

| Estado | Origen típico | ¿Se presenta en PPTX por defecto? |
|:--|:--|:--:|
| `OK` | TOM o loop ADO representable. | Sí. |
| `OK_SIN_BASE` | Valores válidos sin N verificada. | Sí, con aviso metodológico. |
| `PENDIENTE_DDF` | Modo `validar`; aún no se tabuló. | No. |
| `SIN_MDD` | Metadata no contiene la variable. | No. |
| `SIN_ESTRUCTURA` | Loop sin rutas reconocidas. | No. |
| `SIN_DATOS` | No hay respuesta válida o magnitudes representables. | No. |
| `ERROR_TOM` | Fallo de ejes, tabla o `Populate`. | No. |
| `ERROR_HTML` | Exportación/asociación/interpretación de tablas. | No. |
| `ERROR_LOOP` | Excepción de loop. | No. |
| `GRAFICO_INVALIDO` | Instrucción S desconocida. | No. |
| `NO_COMPATIBLE` | Caso analítico no soportado. | No. |

Con `--incluir-errores` se permiten paneles diagnósticos sin gráficos para estados no representables. El objetivo no es sustituir datos perdidos por ceros. Los estados del CSV permiten diferenciar “no había respuesta” de “la conexión no funcionó”.

# 14. Incidencias conocidas y diagnóstico dirigido

## 14.1. COM: argumento inválido al establecer idioma

**Síntoma:** `pywintypes.com_error ... -2147024809` en `documento.Language = 'ESM'`.

**Interpretación:** la carga MDD/DDF puede haber funcionado; falla la asignación de idioma. Ejecutar con `--idioma AUTO`, comprobar idioma del estudio y no forzar `ESM` salvo que esté disponible.

## 14.2. ADO no disponible pero TOM funciona

**Síntoma:** `[ADO] No disponible`, `OK_SIN_BASE` o loops sin detalles.

**Acciones:** verificar registro de `mrOleDB.Provider.2`, compatibilidad de plataforma, permisos sobre archivos, y no usar `--sin-ado`. No reconstruir `n` desde un porcentaje redondeado. Con ponderación activa, la omisión ADO es deliberada.

## 14.3. MDD indexa el loop, pero falla una ruta

**Síntoma:** `[LOOP ITEM ERROR] B1 B1[{_1}].Rp ...`.

**Acciones:** comprobar `metadata.campos_grid('B1')`, tipo del campo `Rp`, sintaxis real de VDATA, diccionario de opciones y versión del proveedor. Usar `--desde/--hasta` para aislar filas. La consola indica la ruta exacta.

## 14.4. HTML no coincide con tabla TOM

**Síntoma:** `[HTML] Asociación por orden ...`, `ERROR_HTML`, tabla sin base o sin categorías.

**Acciones:** ejecutar `--guardar-html`, localizar `PPT_####` y confirmar que base, columnas y categorías son las esperadas. Una asociación por orden es un respaldo, no garantía semántica independiente.

## 14.5. Donas o apilados reemplazados por barras

El código evita representaciones engañosas por suma no cercana a 100 %, exceso de categorías, RM o cruces. Revisar `Resultado.tipo`, `Resultado.nota` y `grafico()` antes de interpretarlo como error. Si la dona desapareció tras personalizar texto blanco, comprobar las posiciones de etiqueta permitidas y restaurar temporalmente la configuración predeterminada.

## 14.6. Datos correctos, gráficos con etiquetas equivocadas

Separar diagnóstico: si JSON y CSV son correctos, el problema es del módulo `ppt.py`. Reproducir con `presentar --datos ...` sin tocar TOM/ADO. Contrastar el objeto chart y sus valores con `python-pptx` para detectar un problema de render/orden.

# 15. Procedimientos de mantenimiento

## 15.1. Validación de código (sin Dimensions)

```powershell
python -m unittest discover -s pruebas -v
```

En la V6 empaquetada se verificaron **11 pruebas** (6 heredadas + 5 nuevas). Estas comprueban lectura/configuración, coherencia de modelos, bases nulas, loops sintéticos, rutas indexadas, prevención de apilado falso, paginación y paleta OOXML. No prueban la instalación de `MDM.Document`, `TOM.Document` ni el DDF real.

## 15.2. Auditoría de un PPTX existente

```powershell
python .\auditar_ppt.py `
  --pptx "Presentacion_Automatica_V6.pptx" `
  --salida "Auditoria_PPT.json"
```

Muestra en terminal `Slides`, `Sin gráficos`, `N=0`, `Conteos None` y ruta del JSON, con detalle por diapositiva. El análisis se diseñó para detectar problemas históricos y **no recalcula cifras del DDF**. Una presentación con último panel único puede aparecer como `con_un_grafico` sin implicar error.

## 15.3. Ciclo seguro de cambios

1. Copiar el ZIP/código a una carpeta con versión controlada (por ejemplo `V6.1`).
2. Cambiar solamente el módulo correspondiente: Excel=`excel_config.py`; loops=`metadata.py`/`loops.py`/`ado.py`; visual=`ppt.py`/`paleta.py`.
3. Ejecutar 11 pruebas y añadir casos específicos si se cambia la lógica.
4. Para cambios visuales, generar desde JSON; para cambios estadísticos, volver a extraer desde el MDD/DDF.
5. Conservar CSV, JSON, log y HTML de una muestra representativa antes y después para comparar.
6. Revisar variables RU, RM, escala, recodificaciones, grids con N variables y donuts/colores de segmentos muy pequeños.
7. Guardar una copia del `ppt.py` usado realmente: puede contener ajustes posteriores ausentes del paquete original.

## 15.4. Puntos de extensión recomendados

- **Catálogo universal de tipos:** extraer los conjuntos `PE26_*` de `excel_config.py` a reglas configurables por estudio.
- **Ponderación por pregunta:** mover de `--ponderador` global a columnas configurables del Excel y documentar bases ponderadas/no ponderadas.
- **Cruces sobre loops:** diseñar una matriz `atributo × opción × segmento`, con denominadores explícitos.
- **Diferencias significativas:** reutilizar capacidades del motor vendor solo tras definir pruebas, banners y dirección de comparaciones, no activarlas implícitamente.
- **Lectura de etiquetas de dona:** separar validación de `data_labels` por tipo de gráfico; registrar excepciones en vez de ocultarlas.
- **Comparadores automáticos:** validar resultados JSON/CSV contra un informe patrón con tolerancias y criterios metodológicos.

# 16. Seguridad, límites operativos y trazabilidad

- Las rutas SQL se validan por expresión regular; el sistema no está diseñado para ejecutar consultas SQL libres de usuarios.
- JSON y CSV contienen datos agregados de la investigación; incluso cuando no contienen respuestas individuales, deben manejarse conforme a las reglas de confidencialidad del estudio.
- El HTML de TOM puede incluir etiquetas originales y resultados tabulados; conservarlo solo cuando sea necesario para diagnóstico.
- Una base vacía o `None` debe representarse como “no verificada”, no como cero; una `n` real exige comprobación.
- Para RM, la suma de porcentajes puede superar 100 %. No cambiar denominadores para que un gráfico se vea más ordenado.
- En gráficos de barras con varias series, ordenar solamente un vector de valores corrompe la correspondencia categoría–serie; aplicar una única permutación a todos.
- El código V6 no garantiza que siempre haya exactamente dos paneles por diapositiva: errores omitidos, loops paginados y totales impares afectan el recuento.
- No asumir que la presentación conserva todos los estilos de `S26002825`: la V6 conserva la **paleta**, no su plantilla completa.

# 17. Mapa rápido: “¿dónde modifico esto?”

| Necesidad | Archivo y función |
|:--|:--|
| Cambiar nombres y número de columnas Excel | `oneshot/excel_config.py → cargar()` |
| Cambiar inferencia de gráfico | `oneshot/excel_config.py → inferir()` |
| Añadir una opción de línea de comandos | `oneshot/cli.py → argumentos()` y `main()` |
| Reconocer nuevos niveles de loop | `oneshot/metadata.py → _indexar()`, `campos_grid()` |
| Ajustar rutas SQL validadas | `oneshot/metadata.py → RUTA` y `oneshot/ado.py → consultar()` |
| Modificar cálculo de N por atributo | `oneshot/loops.py → extraer_loop()` |
| Ajustar tolerancia TOM frente a DDF | `oneshot/ado.py → verificar_conteos()` |
| Cambiar banner/estadística TOM | `oneshot/extraccion.py → construir_solicitud()` / motor vendor |
| Ajustar asociación de tablas HTML | `oneshot/extraccion.py → extraer()` / motor vendor |
| Cambiar formato y orden de barras | `oneshot/ppt.py → grafico()` |
| Cambiar etiquetas de dona/apiladas | `oneshot/ppt.py → grafico()` |
| Cambiar colores | `oneshot/paleta.py → S26002825`, `obtener()` |
| Cambiar cuántos atributos por panel | `oneshot/ppt.py → paginar_loops()` / CLI |
| Cambiar estructura de resultados JSON | `oneshot/modelos.py`, `oneshot/salidas.py` |
| Cambiar columnas de CSV | `oneshot/salidas.py → guardar_csv()`, `guardar_detalle()` |
| Añadir pruebas de regresión | `pruebas/test_offline.py` o `test_loops_paleta.py` |


# 18. Etiquetas adicionales del motor TOM reutilizado

El archivo `oneshot/vendor/motor_dimensions_tom.py` también puede emitir las siguientes etiquetas, **especialmente si se reutiliza el motor para otras tareas de banners o tablas**. No todas aparecen en el recorrido normal de OneShot V6, porque el generador principal utiliza solo una parte de sus funciones.

| Etiqueta adicional | Lectura operativa |
|:--|:--|
| `[BANNER]` / `[BANNER DIFSIG]` | Definición de eje TOP y, en el motor heredado, comparaciones de significancia. |
| `[GRID]` | Reconocimiento de una tabla de iteradores, con SIDE y TOP propios de grid. |
| `[BASES]` | Bases reales y ponderadas recuperadas durante el procesamiento semántico. |
| `[FORMATO]` / `[ETIQUETA MDD]` | Ajuste o recuperación de etiquetas, o aviso por texto metadata faltante. |
| `[INDICADORES]` / `[GRID DIFSIG]` | Postproceso de marcas de diferencias significativas; no equivale por sí solo a un contraste configurado en OneShot. |
| `[AVISO ESTADÍSTICO]` | Alerta de consistencia estadística, por ejemplo respecto de las bases. |
| `[AVISO PONDERACION]` | Advertencia particular de pesos o bases; revisar `--ponderador`. |
| `[PRECISION]` | Precisión de resultados y símbolos que el exportador no pudo representar fielmente. |
| `[POPULATE]` | Tabulación de tabla solicitada por el motor heredado. |
| `[ORDEN]` | Ordenación de filas de una tabla externa; **no es** la ordenación de barras del PPTX. |
| `[COLORES]` | Conservación de rellenos en un Excel de tablas, función legada no utilizada para la paleta de PPTX. |
| `[DASHBOARD]`, `[PLAN]`, `[REFORMATEAR]` | Procesos auxiliares del motor original que no intervienen normalmente en el generador de PowerPoint. |
| `[COM ERROR]`, `[FATAL]` | Fallo COM o error no recuperado; conservar la primera excepción completa y su traceback. |
| `[FINALIZADO]` | Finalización de una ejecución del motor heredado; no reemplaza la confirmación `[PPTX]` de OneShot. |

Cuando exista una discrepancia entre mensajes del motor vendor y el estado final del generador, dar prioridad a **`*.reporte.csv`**, **`*.detalle.csv`** y **`[ESTADOS]`**: reflejan la adaptación semántica con la que se decidió qué se graficó.

---

**Fuente revisada:** archivos Python y README del paquete `Generador_PPT_V6_Loops_Paleta_S26002825.zip` adjunto en la conversación. Se distinguieron comportamiento implementado, pruebas offline y recomendaciones de mejora. Esta documentación no sustituye la validación estadística del proyecto con el MDD/DDF nativos y su proveedor de Dimensions.
