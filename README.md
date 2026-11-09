# Reorganizador de presupuestos

Aplicación web que recibe un presupuesto de obra en Excel, encuentra la hoja del **presupuesto detallado** y la reorganiza:

```
ANTES                                   DESPUÉS
Unidad estructural                      Capítulo
 └─ Capítulo                             └─ Unidad estructural
     └─ Actividad                            └─ Actividad
```

Todo se procesa **en el navegador**: el archivo no se sube a ningún servidor.

## Cómo se usa

1. Abre la página (GitHub Pages) y arrastra el `.xlsx` o `.xlsm`.
2. Revisa la hoja detectada y las columnas (Ítem, APU, Descripción, Und, Cant, Vr. unitario, Vr. total). Si algo quedó mal, se corrige con los selectores.
3. Opcional: marca columnas adicionales (otros escenarios o fases de VR. TOTAL) y si deben sumarse.
4. Revisa el rótulo: debe decir **Cierra exacto**. Lee las alertas de estructura.
5. Elige la salida y descarga:
   - **Libro nuevo**: solo las hojas reorganizadas. Las actividades van como valores y los subtotales como fórmulas.
   - **Mismo libro, hojas nuevas vinculadas**: tu archivo completo con tres hojas al final. Cada actividad es una fórmula que apunta a su celda original (`='Ppto Detallado'!I33`) y conserva su formato. El resto del libro no se toca.

### Hojas de salida

| Hoja | Contenido |
|---|---|
| `PPTO x CAPITULO` | Capítulo › UE › actividad, con agrupación por niveles (+/−), autofiltro, columnas UE / CAPÍTULO / SUBGRUPO y FILA ORIGEN (en el modo vinculado es un hipervínculo a la fila original). |
| `VERIFICACION` | Cierre contra el original: costos directos, indirectos, total, cada UE (contra su SUB-TOTAL) y cada capítulo. Columna ESTADO = OK / REVISAR. |
| `MATRIZ CAP x UE` | Matriz capítulo × unidad estructural con `SUMIFS`, y un control contra costos directos que debe dar 0. |

## Reglas de lectura

No depende del nombre de la hoja, de la fila del encabezado ni de la posición de las columnas:

- **Hoja**: la que tenga una fila de encabezado con DESCRIPCIÓN / UND / CANT / VR UNIT / VR TOTAL y más estructura de presupuesto (UE y actividades). Se puede cambiar.
- **Unidad estructural**: la primera fila de grupo después de un `SUB-TOTAL` (o del encabezado). Códigos tipo `01`, `S-5`, etc. Se identifica por código + nombre, y si se repite se numera `(2)`.
- **Capítulo**: fila sin unidad, con código entero o con valor total.
- **Actividad**: fila con unidad; o con cantidad y valor unitario; o con código decimal (`12.056`).
- **Subgrupo**: fila sin código ni valor dentro de un capítulo (p. ej. "EXCAVACIÓN SEMISÓTANO").
- **Costos indirectos**: lo que va después de `SUB-TOTAL COSTOS DIRECTOS` (o `SUB-TOTAL CD`) se conserva al final en su orden original.
- **Fin**: `TOTAL PRESUPUESTO …`. Lo que haya debajo se ignora.

Decisiones de reorganización:

- La actividad va al capítulo **donde está ubicada** en el presupuesto, aunque su código diga otro.
- Los capítulos se unifican **por código**, con el nombre más frecuente. Los conflictos de nombre quedan como alerta.
- Un capítulo sin código numérico se asigna por nombre a otro capítulo existente (con alerta). Si no hay coincidencia, queda como capítulo propio.

## Desarrollo

Sin framework ni paso de compilación: HTML + módulos ES. Las librerías van incluidas en `lib/` (JSZip 3.10.1 y ExcelJS 4.4.0), así que la app no depende de CDNs.

```
src/xlsx-reader.js   lector .xlsx propio sobre JSZip (valores cacheados, estilos, combinaciones)
src/classify.js      detección de hoja y encabezado, clasificación de filas y pivote
src/layout.js        estructura de las hojas de salida y fórmulas
src/writer-new.js    salida "libro nuevo" (ExcelJS)
src/writer-inject.js salida "mismo libro" (agrega hojas al .xlsx sin reescribir lo demás)
src/app.js           interfaz
```

Pruebas locales con Node 18+ y Python 3 (openpyxl):

```bash
npm install
node test/run.mjs ruta/al/presupuesto.xlsx test/out      # genera _NUEVO y _VINCULADO e imprime el cierre
python3 test/verify_formulas.py ruta/al/presupuesto.xlsx test/out/<archivo>_VINCULADO.xlsx   # evalúa todas las fórmulas nuevas
python3 test/xmlcheck.py test/out/<archivo>_VINCULADO.xlsx                                   # valida el XML inyectado
python3 -m http.server 8000                             # app en http://localhost:8000
```

`.gitignore` excluye `*.xlsx` / `*.xlsm` para que ningún presupuesto termine en el repositorio.
