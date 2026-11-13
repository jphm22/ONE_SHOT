# Generador PPT V6 — Loops + paleta S26002825

Proyecto Python modular que lee `Variables_Graficos.xlsx` a partir de la fila 5 y produce gráficos nativos (editables) en PowerPoint a partir de los archivos **MDD/DDF nativos**, sin utilizar TXT.

## Novedades frente a V5

1. **Loops/Grid:** lee `Block/Grid with iterators` incluyendo los subcampos de preguntas como `B1`, `B3`, `B8a`, `B11`, `B12`, `B19`, `T2` y `T13`. `metadata.py` explora recursivamente `Variables`, `Fields`, `ClassFields` y elementos de iterador del MDD. Cuando TOM no puede interpretar un loop, el respaldo ADODB consulta campos del tipo `B1[{_1}].Rp`, utilizando rutas y códigos del MDD (misma estrategia que `generar_ppt_dimensions_corregido.py`).
2. **N por atributo:** DDF proporciona las bases y frecuencias por subcampo. Para loops RM no se fuerza la suma a 100 %, ni se inventa un N global cuando cada atributo tiene su propia base.
3. **Paleta predeterminada:** usa los colores de los gráficos de `S26002825 2026-11-02T230611.142.pptx`, no colores aproximados de la plantilla. Hex: `#1DAFAD`, `#121B5A`, `#E50158`, `#103C50`, `#E2DA51`, `#451B67`, `#2DB574`, `#FF740F`, con extensiones de la misma presentación. Se aplica a todas las series, puntos de anillos y líneas. `--paleta v5` recupera la paleta anterior.
4. **Todos los atributos:** `--max-atributos 0` (predeterminado) consulta todos los atributos encontrados; `--atributos-por-grafico 8` los distribuye en paneles de 8 sin eliminar los demás, manteniendo dos gráficos por lámina siempre que existan al menos dos paneles.
5. **Trazabilidad:** `*.detalle.csv` ahora registra todas las combinaciones atributo × respuesta de los loops, incluidos `n`, porcentaje, base y ruta DDF. `*.resultados.json` conserva las distribuciones completas para volver a generar el PPT sin consultar Dimensions.

## Requisitos en Windows

- Python 3.11 recomendado, PowerPoint no es imprescindible para la construcción OOXML.
- IBM/SPSS/UNICOM Dimensions con `TOM.Document`, `MDM.Document` y `mrOleDB.Provider.2` instalados y registrados.
- `pip install pywin32 python-pptx`

Extraiga el ZIP completo y copie Excel, `.mdd` y `.ddf` en la carpeta donde ejecutará el generador, o especifique rutas absolutas. **No se necesita otro catálogo.**

## Comando principal

```powershell
python .\generar_ppt.py `
  --excel "Variables_Graficos_Clasificados.xlsx" `
  --mdd "PE26-003581-01.mdd" `
  --ddf "PE26-003581-01.ddf" `
  --salida "Presentacion_Automatica_V6.pptx" `
  --guardar-html
```

También puede usar el Excel original `Variables_Graficos.xlsx`: los gráficos sin tipo explícito se infieren igual que en V5.

### Ajustes opcionales

```powershell
# Gráfico con más atributos (puede aumentar densidad visual)
python .\generar_ppt.py --excel "Variables_Graficos.xlsx" --mdd "PE26-003581-01.mdd" --ddf "PE26-003581-01.ddf" --salida "Presentacion_V6.pptx" --atributos-por-grafico 10

# Renderizar de nuevo usando resultados existentes, sin abrir MDD/DDF
python .\generar_ppt.py presentar --datos "Presentacion_Automatica_V6.resultados.json" --salida "Presentacion_Rediseno.pptx" --paleta s26002825

# Solo un rango de filas para comprobar variables loop
python .\generar_ppt.py --excel "Variables_Graficos.xlsx" --mdd "PE26-003581-01.mdd" --ddf "PE26-003581-01.ddf" --desde 27 --hasta 44 --salida "Prueba_Loops.pptx" --guardar-html
```

`--idioma` sigue siendo `AUTO` por defecto para no provocar el error `documento.Language = 'ESM'`. Si su MDD realmente admite otro idioma, puede seleccionarlo explícitamente.

## Archivos del proyecto

| Archivo | Responsabilidad |
| --- | --- |
| `generar_ppt.py` | Entrada CLI única |
| `oneshot/excel_config.py` | Lectura del Excel y reglas de selección de gráfico |
| `oneshot/metadata.py` | Indexación recursiva del MDD; descubrimiento de iteradores |
| `oneshot/ado.py` | Consulta nativa DDF, incluyendo rutas indexadas de loop |
| `oneshot/loops.py` | Frecuencias, porcentajes y N por atributo; manejo RM |
| `oneshot/extraccion.py` | TOM, exportación semántica, orquestación y respaldo ADODB |
| `oneshot/vendor/motor_dimensions_tom.py` | Motor TOM preexistente, preservado |
| `oneshot/modelos.py` | Objetos de resultados y metadatos estadísticos |
| `oneshot/paleta.py` | Paleta S26002825 y alternativa V5 |
| `oneshot/ppt.py` | Gráficos editables, paleta y paginación de loops |
| `oneshot/salidas.py` | JSON y reportes CSV |
| `pruebas/` | Pruebas sin Windows/Dimensions |

## Reportes y advertencias

- `*.diagnostico.log`: errores `[TOM ERROR]`, `[LOOP ITEM ERROR]` y `[LOOP OK]`.
- `*.reporte.csv`: estado `OK`/`SIN_DATOS`/`ERROR_LOOP` y número de atributos recuperados.
- `*.detalle.csv`: cada categoría por atributo, con N y porcentajes.
- `*.resultados.json`: persistencia para `presentar` sin Dimensions.
- `*.dimensions.htm`: exportación temporal de TOM solamente con `--guardar-html`.

Si una variable de loop carece de diccionario de respuestas en el MDD, la consulta del ítem se omite y se registra el problema. No se crean datos ni etiquetas ficticias. El comportamiento completo sigue sujeto a la versión de proveedor COM/ADODB instalada.

## Validaciones realizadas

Las pruebas locales verifican parsing de campos indexados, lectura de un loop sintético de 13 atributos con distintas bases, prevención de 100 % engañoso para RM, paginación en dos gráficos y colores OOXML de la referencia, además de las seis pruebas heredadas de V5. **No son una prueba contra el DDF real**: para validar la conexión de su sistema, ejecute el comando anterior en Windows y examine los reportes.

Para ejecutar pruebas: `python -m unittest discover -s pruebas -v`.
