# Cambios V5 → V6

- `oneshot/metadata.py`: búsqueda recursiva de campos indexados, `campos_grid()`, `etiqueta_item()` y `categorias_campo()`.
- `oneshot/ado.py`: consulta SQL de rutas válidas como `B1[{_1}].Rp` sin permitir SQL libre.
- `oneshot/loops.py`: **nuevo** procesamiento por subcampo de loop, N por atributo, porcentajes, tratamiento de RM y reporte detallado.
- `oneshot/extraccion.py`: prioriza matrices ADODB verificadas para loops, mantiene TOM para variables simples y como fallback; evita tratar recodificaciones simples como loops.
- `oneshot/modelos.py`: añade `bases_por_item` y `detalles_loop` al JSON (compatible con lectura de JSON V5).
- `oneshot/paleta.py`: **nuevo** conjunto de colores exactos de las series del PPTX S26002825.
- `oneshot/ppt.py`: colores personalizados a todas las series y segmentos de dona, apilados válidos, y paneles sucesivos para loops largos.
- `oneshot/cli.py`: agrega `--paleta`, `--atributos-por-grafico` y consulta todos los atributos por defecto (`--max-atributos 0`).
- `oneshot/salidas.py`: reporte CSV incluye cantidad de atributos y detalle expandido de loops; JSON v6 con lectura retrocompatible de v5.
- `pruebas/test_loops_paleta.py`: nuevas pruebas independientes de COM.

**Advertencia de validación:** Se utilizaron pruebas offline y metadatos de referencia; la ejecución real con `MDM.Document`, `TOM.Document` y `mrOleDB.Provider.2` debe verificarse en Windows con los archivos nativos de su estudio.
