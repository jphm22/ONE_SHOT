---
title: "OneShot AutoCharting V6"
subtitle: "Manual de usuario · Generación automática de PowerPoint desde Excel + MDD/DDF"
author: "Documentación del proyecto"
date: "10 de noviembre de 2026"
lang: es-PE
---

> **Alcance y versión.** Este manual describe el proyecto incluido en `Generador_PPT_V6_Loops_Paleta_S26002825.zip`. Los cambios que hayas realizado posteriormente en tu copia de `oneshot/ppt.py` (ocultar ejes, quitar cuadrículas, redondear porcentajes, ordenar barras o cambiar el color de etiquetas) se explican como personalizaciones, pero **no se asume que ya estén incorporados al ZIP de referencia**.

# 1. ¿Qué hace OneShot AutoCharting?

OneShot AutoCharting V6 convierte una lista de variables del cuestionario en una presentación `.pptx` con **gráficos editables**. Toma la configuración de **un único Excel**, consulta la metadata de un `.mdd` y las respuestas del `.ddf` de Dimensions y construye, normalmente, **dos paneles por diapositiva**. Para preguntas largas de tipo *loop/grid*, divide sus atributos entre varios paneles sin desechar el detalle estadístico.

El proceso también entrega resultados estructurados en JSON, una auditoría CSV por variable y un detalle CSV por categoría o atributo. Se puede **regenerar la presentación desde el JSON** sin volver a consultar Dimensions cuando solo se cambia el diseño.

**Qué no hace automáticamente:** no verifica criterios metodológicos no definidos (por ejemplo, exclusiones de códigos especiales, ponderación específica por pregunta, análisis de significancia o recodificaciones de negocio); estos deben ser revisados. Tampoco convierte el MDD/DDF en TXT: trabaja con los archivos nativos. El HTML opcional es una **exportación de tablas calculadas por TOM**, no una conversión del MDD.

# 2. Requisitos previos

| Requisito | ¿Para qué se necesita? |
|:--|:--|
| Windows y Python 3.11 recomendado | Ejecutar la aplicación y la conexión COM de Dimensions. |
| Dimensions / UNICOM / IBM Data Collection | Tener registrados `MDM.Document`, `TOM.Document` y `mrOleDB.Provider.2`. |
| MDD y DDF del mismo estudio | Estructura, códigos, etiquetas y respuestas. |
| `Variables_Graficos.xlsx` o versión clasificada | Variables que se deben graficar desde la fila 5. |
| Bibliotecas de Python | `pywin32`, `python-pptx` y `openpyxl`. Esta última se importa a través del motor TOM reutilizado. |

Instala los paquetes desde PowerShell:

```powershell
python -m pip install --upgrade pywin32 python-pptx openpyxl
```

No necesitas instalar Excel ni automatizar la aplicación PowerPoint para generar el `.pptx`: se construye con `python-pptx`. **La extracción de MDD/DDF sí requiere Dimensions instalado y accesible por COM/ADO**.

# 3. Preparación de los archivos

1. Descomprime **todo** el proyecto, sin mover ni borrar la carpeta `oneshot/`.
2. Copia el Excel, el MDD y el DDF a la carpeta del proyecto, o conserva sus rutas absolutas.
3. Comprueba que el Excel corresponde al estudio y que el MDD y el DDF sean compatibles entre sí.
4. Revisa la columna **S: “Observación adicional por la BU / SL”**. Puede contener un tipo de gráfico; si está vacía, el programa infiere uno y deja una advertencia.
5. Mantén una copia sin modificar de tu Excel y de las fuentes nativas.

Estructura esperada:

```text
Generador_PPT_V6/
  generar_ppt.py
  generar_ppt_dimensions_v3.py       # compatibilidad con comandos anteriores
  auditar_ppt.py
  oneshot/
    cli.py  excel_config.py  metadata.py  ado.py  loops.py
    extraccion.py  modelos.py  ppt.py  paleta.py  salidas.py
    vendor/motor_dimensions_tom.py
  pruebas/
  Variables_Graficos.xlsx
  PE26-003581-01.mdd
  PE26-003581-01.ddf
```

# 4. Cómo configurar el Excel

El lector toma las variables **desde la fila 5**. La V6 consulta las siguientes **posiciones de columna**, no busca automáticamente los encabezados por nombre:

| Columna | Contenido | Uso |
|:--:|:--|:--|
| E | Tipo de pregunta | Distingue, entre otros, variables numéricas para promedio. |
| F | Estructura (p. ej. `Simple Qst`, `Block/Grid with iterators`) | Ayuda a detectar candidatos a loop. **No equivale necesariamente a RU/RM.** |
| G | `VARIABLE` | Nombre técnico exacto del MDD. Es obligatorio. |
| H | Etiqueta de la variable (BDD) | Título de respaldo cuando el MDD no ofrece etiqueta legible. |
| J | `¿Escala?` | Ayuda a inferir gráficos de escala cuando S está vacía. |
| S | `Observación adicional por la BU / SL` | Tipo de gráfico, opcionalmente con cruce. |

Los tipos admitidos en S son `barras`, `columnas`, `dona`/`anillo`, `apiladas`, `escala`, `promedio`, `linea`, `radar`, `tabla` y `funnel`/`embudo`. Hay sinónimos reconocidos. `funnel` utiliza actualmente un gráfico de barras como representación de respaldo.

**Ejemplos de la columna S**:

| VARIABLE | Valor en S | Efecto |
|:--|:--|:--|
| `F1` | `dona` | Distribución de género en anillo, si su suma es compatible con 100 %. |
| `A2` | `barras` | Barras horizontales con porcentajes por categoría. |
| `A3` | `barras \| F1` | Cruce por género (barras con series por segmento). |
| `B8a` | `escala` | Escala, si existe distribución apta para apilar. |
| `F2` | `promedio` | Media aritmética; **no** se interpreta como porcentaje. |

La sintaxis es **`tipo | VARIABLE_DE_CRUCE`**: una sola barra vertical. El cruce debe referirse a un nombre simple existente en el MDD. El programa no cruza automáticamente un loop completo con otra variable: para ese escenario se necesita una ruta indexada o una extensión adicional.

**Nota sobre tipos RU/RM:** un total de porcentajes superior a 100 % puede ser correcto en respuestas múltiples. No interpretes un RM como error ni lo conviertas forzosamente en un apilado al 100 %.

# 5. Primer uso: generar la presentación

Abre PowerShell en la carpeta principal y ejecuta:

```powershell
python .\generar_ppt.py `
  --excel "Variables_Graficos.xlsx" `
  --mdd "PE26-003581-01.mdd" `
  --ddf "PE26-003581-01.ddf" `
  --salida "Presentacion_Automatica_V6.pptx" `
  --guardar-html
```

Si utilizas el Excel ya clasificado, sustituye únicamente el valor de `--excel`. **No se solicita un catálogo independiente**.

Durante la ejecución verás mensajes `[EXCEL]`, `[MDD LOOPS]`, `[TOM OK]`, `[DATOS]`, `[ADO]`, `[LOOP OK]`, `[ESTADOS]` y finalmente `[PPTX]`, según las variables y el proveedor. Una línea `[TOM OK]` confirma la tabulación, pero no necesariamente que la base N esté verificada; revisa el reporte CSV.

# 6. Otros modos de trabajo

El comando principal admite cuatro acciones:

| Acción | Necesita | Qué hace | Qué entrega |
|:--|:--|:--|:--|
| `todo` (predeterminado) | Excel + MDD + DDF | Extrae, registra resultados y crea PPTX. | `.pptx`, `.resultados.json`, `.reporte.csv`, `.detalle.csv`, `.diagnostico.log`; HTML si se solicita. |
| `validar` | Solo Excel | Comprueba variables y tipos de gráfico. **No valida existencia en MDD.** | `.reporte.csv` con `PENDIENTE_DDF`, `.diagnostico.log`. |
| `extraer` | Excel + MDD + DDF | Tabula, valida datos y genera archivos intermedios, **sin PPTX**. | JSON, CSV, log y HTML opcional. |
| `presentar` | JSON previamente generado | Produce otra presentación sin abrir Dimensions. | `.pptx` y `.diagnostico.log`. |

**Validar solo Excel**:

```powershell
python .\generar_ppt.py validar `
  --excel "Variables_Graficos.xlsx" `
  --salida "Validacion.pptx"
```

**Extraer únicamente resultados**:

```powershell
python .\generar_ppt.py extraer `
  --excel "Variables_Graficos.xlsx" `
  --mdd "PE26-003581-01.mdd" `
  --ddf "PE26-003581-01.ddf" `
  --salida "Datos_Estudio.pptx" `
  --guardar-html
```

El valor de `--salida` determina el nombre base de los reportes; en modo `extraer`, el archivo `.pptx` **no se crea**.

**Regenerar solo el PowerPoint**:

```powershell
python .\generar_ppt.py presentar `
  --datos "Datos_Estudio.resultados.json" `
  --salida "Presentacion_Rediseno.pptx" `
  --paleta s26002825
```

# 7. Parámetros más útiles

| Parámetro | Valor predeterminado | Cuándo usarlo |
|:--|:--|:--|
| `--desde`, `--hasta` | Desde fila 5; sin límite final | Probar un tramo de filas del Excel. |
| `--hoja` | Primera hoja | Elegir otra hoja de configuración. |
| `--idioma` | `AUTO` | Mantener idioma activo del MDD; evitar forzar `ESM` si falla. |
| `--paleta` | `s26002825` | Alternar a `v5` si deseas recuperar la paleta anterior. |
| `--atributos-por-grafico` | `8` | Repartir loops extensos entre paneles, sin perder atributos. |
| `--max-atributos` | `0` (todos) | Limitar la **extracción** de atributos de un loop. Distinto de paginación visual. |
| `--max-categorias` | `12` | Mostrar las principales categorías en ciertos gráficos simples. |
| `--decimales` | `3` | Precisión de tabulación TOM; no altera por sí solo la etiqueta del PPTX. |
| `--ponderador` | Vacío | Activar ponderación TOM. **Deshabilita el respaldo ADODB en esta versión.** |
| `--sin-ado` | Desactivado | Omitir la verificación directa del DDF; puede impedir recuperar bases de loops. |
| `--incluir-errores` | Desactivado | Añadir paneles de error al PPTX para depuración. |
| `--guardar-html` | Desactivado | Guardar el HTML de resultados exportado por TOM. |
| `--variable-total`, `--codigo-total` | `TOTAL_BASE`, `_1` | Cambiar la definición del total según el estudio. |

**Probar B1/B3/B8a y variables próximas**:

```powershell
python .\generar_ppt.py `
  --excel "Variables_Graficos.xlsx" `
  --mdd "PE26-003581-01.mdd" `
  --ddf "PE26-003581-01.ddf" `
  --desde 27 --hasta 44 `
  --salida "Prueba_Loops.pptx" --guardar-html
```

# 8. ¿Qué significan los archivos de salida?

Para `--salida "Informe.pptx"`:

| Archivo | Contenido | Para qué revisarlo |
|:--|:--|:--|
| `Informe.pptx` | Diapositivas con gráficos nativos editables. | Revisión visual y entrega. |
| `Informe.resultados.json` | Resultados íntegros, atributos, bases verificadas y series. | Reprocesar el diseño con `presentar`. |
| `Informe.reporte.csv` | **Una fila por variable** con estado, base, gráfico y advertencias. | QA y localizar variables omitidas. |
| `Informe.detalle.csv` | Categoría/segmento, `n`, porcentaje, `base_N`, fuente. | Validación numérica y loops. |
| `Informe.diagnostico.log` | Registro cronológico de operaciones y excepciones. | Resolver errores con el desarrollador. |
| `Informe.dimensions.htm` | Tablas exportadas por TOM. **Solo si usas `--guardar-html`.** | Revisar la salida del proveedor y discrepancias. |

El CSV utiliza `;` como separador y `UTF-8 con BOM`, para facilitar su apertura en Excel de Windows. El JSON permite una revisión técnica y contiene las respuestas recuperadas; maneja estos archivos como datos del estudio.

# 9. Estados del reporte y decisiones

| Estado | Qué significa | Qué hacer |
|:--|:--|:--|
| `OK` | Resultado utilizable; la base puede venir de TOM o DDF. | Confirmar sentido metodológico. |
| `OK_SIN_BASE` | Porcentajes presentes, pero la base N no se verificó. | Revisar log, HTML, proveedor y `--sin-ado`. |
| `PENDIENTE_DDF` | Solo se hizo `validar` con Excel. | Ejecutar `extraer` o `todo`. |
| `SIN_MDD` | Variable no encontrada en metadatos nativos. | Verificar nombre G, archivo MDD y rutas indexadas. |
| `SIN_DATOS` | No hubo valores representables suficientes. | Revisar filtros, subcampos, TOM y DDF. |
| `SIN_ESTRUCTURA` | No se encontraron rutas de loop comprobables en MDD. | Comprobar iteradores y subcampos `Rp`. |
| `ERROR_TOM` | Falló la creación o tabulación en Dimensions TOM. | Consultar `[TOM ERROR]` y traceback. |
| `ERROR_HTML` | Falló la asociación o interpretación del HTML tabulado. | Ejecutar con `--guardar-html`. |
| `ERROR_LOOP` | Falló la recuperación del loop. | Consultar `[LOOP ERROR]`/`[LOOP ITEM ERROR]`. |
| `GRAFICO_INVALIDO` | Tipo no reconocido en S. | Corregir la instrucción en Excel. |
| `NO_COMPATIBLE` | Ejemplo: cruce externo sobre loop completo. | Configurar análisis por ruta o ampliar código. |

Por defecto, el PPTX **omite variables no representables**. Por eso el número de diapositivas puede ser inferior al número de variables dividido por dos. Con un número impar de paneles, la última lámina tendrá uno solo. `--incluir-errores` permite mostrar problemas en lugar de omitirlos.

# 10. Cómo leer la terminal

Ejemplo **ilustrativo** (los números cambian según el estudio):

```text
2026-... INFO [EXCEL] 137 variables | 0 gráficos inferidos | estudio PE26=True
2026-... INFO [CONFIGURACIÓN] 137 variables, 0 inferidas | errores catálogo 0
2026-... INFO [VISUAL] paleta=s26002825, atributos por gráfico=8, límite extracción=TODOS
2026-... INFO [MDD LOOPS] 12 estructuras indexadas confirmadas
2026-... INFO [TOM OK] fila=27 variable=B1
2026-... INFO [ADO] Respaldo DDF habilitado para verificar conteos
2026-... INFO [LOOP OK] B1 | atributos=15 / 15 | opciones=10 | bases=... | fallos=0
2026-... INFO [ESTADOS] {'OK': ..., 'OK_SIN_BASE': ...}
2026-... INFO [PPTX] ... diapositivas; ... paneles; ... errores omitidos
```

- `INFO`: avance normal. `WARNING`: advertencia que permite continuar. `ERROR`: operación fallida; consulta el traceback.
- `[EXCEL]`: cuántas variables se leyeron y cuántos gráficos se infirieron.
- `[CONFIGURACIÓN]`: conjunto seleccionado con `--desde/--hasta`.
- `[VISUAL]`: paleta y paginación de loops.
- `[MDD LOOPS]`: loops verdaderamente confirmados por metadata, no solo por la etiqueta del Excel.
- `[TOM OK]`: tabulación de una variable terminada; **no garantiza N validado**.
- `[DATOS]`: resultado interpretado de TOM, estado, base y número de opciones.
- `[ADO]`: respaldo de conteos DDF disponible o rechazado.
- `[LOOP OK]`: cantidad de atributos recuperados, opciones, rango de bases y fallos.
- `[LOOP REEMPLAZADO]`: una matriz obtenida por ADO reemplazó una salida TOM, si pudo validarse.
- `[ESTADOS]`: resumen de estados por variable.
- `[PPTX]`: número final de láminas y paneles, o elementos omitidos.

Para interpretar todas las etiquetas de terminal y los errores COM, consulta el **Manual técnico**, sección “Diccionario de mensajes”.

# 11. Personalizaciones visuales frecuentes

Se modifican en `oneshot/ppt.py`, función `grafico()`; la configuración de colores se define en `oneshot/paleta.py`. Estas son las ubicaciones, **no una garantía de que tu copia tenga el mismo número de línea**.

| Cambio deseado | Ubicación / propiedad |
|:--|:--|
| Quitar líneas de cuadrícula | `chart.value_axis.has_major_gridlines = False` y `has_minor_gridlines = False`. |
| Ocultar el eje de porcentajes | `chart.value_axis.visible = False` para gráficos con eje. |
| Etiquetas a cero decimales | `s.data_labels.number_format = '0"%"'` para porcentajes; conservar medias con formato propio. |
| Barras de mayor a menor, de arriba abajo | Reordenar categorías y **todas** las series juntas, solamente si `tipo == 'barras'`; considerar el orden de ejes de PowerPoint. |
| Texto blanco en apiladas y escalas | `s.data_labels.font.color.rgb = rgb((255,255,255))`; probar posición centrada. |
| Texto blanco en dona | Aplicar el color dentro de la rama `tipo == 'dona'`, sin forzar una posición incompatible; verificar contraste en segmentos claros. |
| Cambiar la paleta | Modificar `S26002825` en `oneshot/paleta.py`, o usar `--paleta v5`. |

**Precaución:** el archivo `ppt.py` incluido en el ZIP V6 todavía establece `d.number_format = '0.0"%"'` y `has_major_gridlines = True`; el resto de cambios conversados fueron ediciones realizadas después sobre tu instalación. Si quieres documentar exactamente esa instalación, conserva una copia actual de tu `ppt.py` modificado.

Para aplicar un cambio **solo visual**, lo más eficiente es volver a generar desde `*.resultados.json` utilizando el modo `presentar`.

# 12. Solución de problemas frecuentes

| Síntoma | Posible causa y medida recomendada |
|:--|:--|
| `documento.Language = ESM` lanza `-2147024809` | Utiliza `--idioma AUTO`; no fuerces un idioma que el estudio no admite. |
| `ModuleNotFoundError: win32com` | Instala `pywin32` con el mismo intérprete que ejecuta el script. |
| `ModuleNotFoundError: openpyxl` | El motor TOM reutilizado lo importa; instala `openpyxl`. |
| Variable simple aparece como `SIN_MDD` | Nombre G incorrecto, MDD distinto o metadatos sin ese campo. |
| Loop sin respuestas | Confirma que existan rutas MDD `B1[{_1}].Rp`, que ADO esté disponible y que no uses `--sin-ado`. |
| Porcentajes pero base no verificada | Revisar `*.detalle.csv`, `*.diagnostico.log` y la consistencia TOM/ADO. Nunca deduzcas N multiplicando porcentajes redondeados. |
| Un apilado se transforma en barras | Las proporciones no forman una distribución completa, son RM o hay cruce; el sistema evita un falso 100 %. |
| Se esperaba dona y sale barra | Dona con varias series, más de 9 categorías o suma incompatible con 100 %: respaldo previsto en el motor. |
| La dona deja de mostrarse tras cambiar etiquetas | Revertir únicamente el bloque de formato de etiquetas en `grafico()`; evitar valores de posición no compatibles y revisar el error/registro. |
| Menos de dos gráficos en la última diapositiva | Cantidad impar de paneles válidos; comportamiento esperado. |
| El programa no genera ningún PPTX | No hay resultados representables; revisar `.reporte.csv`. V6 evita guardar presentaciones vacías. |
| `--exigir-catalogo` | Es una opción heredada ignorada en V6; ya no se requiere otro catálogo. |

# 13. Lista de control antes de entregar un informe

- [ ] Se usaron el MDD y DDF correctos y del mismo corte/ola.
- [ ] La columna G contiene nombres técnicos existentes; S no tiene tipos desconocidos.
- [ ] Se revisaron todas las filas del `.reporte.csv` con estado distinto de `OK`.
- [ ] Se verificó `base_N` en `.detalle.csv`, especialmente cuando hay loops o cruces.
- [ ] Las preguntas RM no se interpretaron como distribuciones de 100 %.
- [ ] Se revisaron categorías especiales (ninguno, no precisa, etc.) y filtros metodológicos.
- [ ] Las etiquetas largas y el orden de barras se leen correctamente en PowerPoint.
- [ ] Se comprobó que el PPTX muestra gráficos editables y no contiene páginas inesperadamente vacías.
- [ ] Se conservaron JSON, CSV y log junto con la versión del código que generó el reporte.

---

**Referencia técnica:** código fuente del paquete `Generador_PPT_V6_Loops_Paleta_S26002825.zip`, en particular `oneshot/cli.py`, `excel_config.py`, `extraccion.py`, `loops.py` y `ppt.py`. Las conexiones COM y resultados del estudio deben validarse en la instalación Windows/Dimensions correspondiente.
