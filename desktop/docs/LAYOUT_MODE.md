# Modo "Igual al PDF" (layout) — diseño

## 1. Por qué existe

El pipeline original (DESIGN.md) piensa todo PDF como "tablas + texto suelto": cada tabla va a su hoja, el
texto que sobra a una hoja "Datos del documento". Eso es correcto para listados, extractos y reportes, pero
**destruye los formularios**: una *Hoja de Ruta*, un remito con cajas de encabezado, una factura con bloques.

Caso real que lo motivó (hoja de ruta de reparto, 2 páginas):

* las 6 cajas "ETIQUETA / Valor" del encabezado (TRANSPORTISTA, AUXILIAR, VEHÍCULO, PATENTE, REPARTO…)
  quedaron fusionadas *dentro* de la tabla de paradas, con etiquetas y valores de cajas distintas mezclados en
  una sola celda ("TRANSPORTISTA AUXILIAR <nombre 1> <nombre 2>");
* la segunda tabla (ENVASES) fue a otra hoja, con el nombre cortado a 31 caracteres;
* el resto (título, N° de viaje, KM, firmas, observaciones, mapa) fue a una hoja de "texto suelto", sin
  posición, y se perdió el mapa y el logo;
* aparecieron dos **"N/A" que en el PDF no se ven**: el PDF los dibuja y luego tapa cada uno con un
  rectángulo blanco (así lo hace el navegador al imprimir a PDF). El extractor leía el texto de la página sin
  considerar qué quedaba tapado.

El modo "Igual al PDF" reproduce **cada página como una hoja**: celdas en la posición que tienen en el PDF,
cajas con sus bordes y rellenos, etiquetas y valores juntos en la misma celda, tipografías y colores, logos y
mapas como imágenes, anchos/altos proporcionales, impresión en A4 ajustada al ancho. El modo "Solo tablas"
(el anterior) sigue existiendo para listados.

## 2. Contrato (ya implementado en `src/pdf2excel/models.py`)

`Options.layout: "auto"|"layout"|"tables"` (default `"auto"`), `Options.layout_images: bool`.

`ConversionReport` gana: `layout_pages: list[LayoutPage]`, `suggested_mode`, `mode` (el que usa `export()`),
`mode_reason` (frase en castellano para la UI). `analyze()` **siempre** calcula `tables`/`text_items` como
hasta ahora y **además** `layout_pages` (ver §6 para el límite por tamaño). `mode` se inicializa con
`Options.layout` (o con `suggested_mode` si es `"auto"`) y la UI puede cambiarlo antes de exportar.

Modelos: `BorderSide`, `TextRun`, `LayoutCell`, `LayoutImage`, `LayoutPage` (leer sus docstrings).
Coordenadas siempre en **puntos PDF, origen arriba a la izquierda** (como PyMuPDF). Filas/columnas 0-based.

Firmas públicas que cada módulo debe respetar (las usan los demás):

```python
# core/visibility.py  (sin Qt)
def hidden_span_rects(page) -> list[pymupdf.Rect]:
    """bbox de los spans de texto que NO se ven: tapados por un relleno opaco dibujado después, fuera de la
    página o del clip, tamaño ~0, opacidad 0, color igual al fondo opaco que tienen debajo, o texto
    invisible (render mode 3) cuando la página NO es un escaneo con capa OCR."""
def visible_words(page, clip=None) -> list[tuple]:
    """Igual que ``page.get_text("words", clip=clip)`` pero sin las palabras ocultas."""

# core/layout.py  (sin Qt)
def extract_layout_page(page, *, embed_images: bool = True, scanned: bool = False) -> LayoutPage:
    """Una página → LayoutPage (grilla + celdas + imágenes). ``scanned=True``: la página es un escaneo con
    capa OCR (texto invisible = texto real; la imagen de fondo NO se incluye)."""

# excel/layout_writer.py  (sin Qt)
def write_layout_sheets(wb: Workbook, report: ConversionReport, opts: Options, used_names: set[str],
                        name_prefix: str = "") -> list[str]:
    """Agrega al libro una hoja por LayoutPage del reporte (nombre "Página N", con prefijo en lotes) y
    devuelve los nombres creados. Configura impresión, ancho de columnas, alto de filas, bordes, rellenos,
    rich text, imágenes, sin líneas de cuadrícula."""

# core/mode.py  (sin Qt)
def suggest_mode(report_like) -> tuple[EffectiveMode, str]: ...   # ver §5
```

## 3. Algoritmo de extracción (`core/layout.py`)

Entrada: una `pymupdf.Page`. Salida: `LayoutPage`. Pasos (el implementador puede mejorarlos si lo mide):

1. **Texto visible.** Spans con `page.get_texttrace()` (trae `seqno`, color, opacidad, tipo de render) o
   `get_text("dict")`; descartar todo lo de `hidden_span_rects`. Guardar en `hidden_texts` lo descartado.
   Orden de dibujo: un span queda tapado si existe un relleno **opaco** (alfa ≥ 0.99, cualquier color) con
   `seqno` mayor cuyo rectángulo lo contiene (≥ 90 % del bbox). Ojo: los rellenos blancos *anteriores* al
   texto son fondo, no lo tapan.
2. **Primitivas gráficas** (`get_drawings`): (a) segmentos de borde horizontales/verticales — líneas, lados de
   rectángulos con trazo, y rectángulos de relleno **finos** (alto o ancho ≤ ~1.5 pt, que es como muchos
   generadores dibujan los bordes; en la hoja de ruta de ejemplo hay 303); fusionar tramos colineales
   contiguos; (b) rellenos de caja (color ≠ blanco, área > fino); (c) cajas cerradas.
3. **Imágenes** (`get_image_info` + `Pixmap`): recomprimir a JPEG q85 (fotos/mapas) o PNG (logos con
   transparencia), lado mayor ≤ 1600 px; ignorar < 4 pt; con `scanned=True` ignorar la imagen de fondo.
4. **Grilla.** Bordes x candidatos = x de segmentos verticales, lados izq/der de cajas y rellenos; bordes y
   candidatos análogos. Para texto que **no** está dentro de ninguna caja, sumar su borde izquierdo (y derecho
   si está alineado a la derecha/centrado) como borde candidato. Agrupar bordes a ≤ 1.5 pt. Las tracks entre
   bordes consecutivos son columnas/filas. **Límite**: ≤ ~80 columnas y ≤ ~250 filas por página; si se pasa,
   subir la tolerancia de agrupamiento gradualmente (no recortar contenido). Tracks de menos de ~2 pt que
   no separan nada se fusionan con la vecina.
5. **Celdas.**
   * Cada **caja cerrada** (rectángulo delimitado por bordes y/o relleno) → una `LayoutCell` con el rango de
     tracks que cubre (`row_span`/`col_span`), aunque no tenga texto (campo sin completar: caja vacía con sus
     bordes — así se ve un formulario en blanco).
   * Las **celdas de una tabla dibujada** (grilla regular de cajas) son simplemente cajas contiguas.
   * Todo el texto contenido en una caja va **a esa celda**: líneas por orden de lectura; dentro de una línea,
     runs por estilo. "ETIQUETA (6 pt gris) / Valor (9 pt negrita)" queda en **una** celda con 2 líneas y
     estilos distintos (rich text), no en dos.
   * Texto fuera de cajas: agrupar palabras en líneas (misma línea de base ± 40 % del alto; hueco horizontal
     < ~1.5×tamaño) y líneas en párrafos (mismo borde izquierdo y separación vertical < ~1.6×tamaño). Cada
     párrafo → una celda (combinada de borde izq. a borde der. del texto, ajustada a la grilla).
   * Alineación horizontal por la posición del texto dentro de su caja (izq./centro/der., tolerancia 15 %);
     vertical `center` por defecto (`top` si el texto de varias líneas pega arriba).
   * Bordes de cada celda desde los segmentos que recorren sus lados (ancho y color); relleno si una caja
     rellena la cubre ≥ 80 %.
   * Estilo de run: negrita/cursiva por flags/nombre de fuente, tamaño y color reales, familia aproximada
     (Arial/Helvetica → "Arial"; Times → "Times New Roman"; Courier → "Courier New"; resto → None).
6. **Valores tipados** (sin inventar): una celda de **un solo run** se convierte a `value` solo si es
   inequívoca — fecha `dd/mm/aaaa` válida; número con formato argentino (`1.234,56`, `$ 1.234,56`, `12,5 %`)
   o entero sin ceros a la izquierda de ≤ 15 dígitos **dentro de una columna de datos** (corrida vertical de
   ≥ 2 celdas contiguas con el mismo `col`/`col_span`; usar `core.typing_ar.classify_column/infer_column_locale`
   para decidir por columna como hace el modo tablas). Horas `hh:mm`, IDs (N°, CUIT, teléfonos, patentes,
   códigos con letras), encabezados (celdas en negrita arriba de una corrida) y cualquier ambigüedad quedan
   como texto. `Options.typed_values=False` → todo texto (el pipeline lo respeta pasando `typed=False`).
7. **Dimensiones**: `col_widths_pt`/`row_heights_pt` = tamaños reales de las tracks; `n_words` = palabras
   visibles.

Casos que **debe** resolver (los cubre el corpus de formularios, §7): cajas "ETIQUETA/valor", tabla con
encabezado de color y filas con líneas finas, filas de sección que combinan todas las columnas, texto libre
a la derecha/izquierda de cajas, bloques en dos columnas, celdas de totales alineadas a la derecha, páginas
con mapa/imagen grande, texto tapado por rectángulos, páginas rotadas (`/Rotate`), sin ningún borde dibujado
(solo texto alineado), y páginas vacías (devolver `LayoutPage` sin celdas).

## 4. Writer (`excel/layout_writer.py`)

* Una hoja por `LayoutPage`, nombre `Página N` (`<prefijo> · Página N` en lotes, ≤ 31 caracteres, únicos).
* Columnas: `ws.column_dimensions[...].width` = puntos / 5.25 (en el archivo el ancho ya incluye el relleno de
  5 px: 1 unidad = 7 px = 5.25 pt; restar el relleno achicaba cada columna 3.75 pt), mínimo 0.2, máximo 255. Filas: `row_dimensions[...].height`
  en puntos (máx. 409). **No se agranda la grilla** (como la hoja se imprime ajustada al ancho, agrandarla achicaba todo el
  texto): columnas y filas miden lo mismo que en el PDF. El texto que no entra en su celda (medido con las métricas
  de Helvetica/Times/Courier; Calibri para el resto) se achica **solo en esa celda** (`cell_font_scale`, hasta 0.7).
* Combinar rangos (`merge_cells`) solo cuando `row_span`/`col_span` > 1; escribir el valor en la esquina y el
  estilo (borde/relleno) en **todas** las celdas del rango (si no, Excel pierde los bordes del rango).
* Texto con varios estilos → `openpyxl.cell.rich_text.CellRichText` + `TextBlock(InlineFont(...))`; un solo
  estilo → `Font`. Tamaños: `max(6, round(size_pt * font_scale))` con la misma escala de la hoja. Color,
  negrita, cursiva, familia. `Alignment(wrap_text, horizontal, vertical)`.
* Bordes `thin` (< 1.2 pt), `medium` (< 2.5 pt), `thick` (≥ 2.5 pt) con su color; relleno sólido.
* `value` tipado → escribir el número/fecha y su `number_format`; si no, texto (con `data_type='s'` cuando
  parece número pero se decidió dejar texto, para que Excel no lo "corrija").
* Imágenes: `openpyxl.drawing.image.Image` desde `BytesIO` (requiere Pillow) con `TwoCellAnchor` (esquina superior
  izquierda en `(row, col, dx, dy)` y esquina opuesta medida sobre la grilla), así la imagen sigue a las columnas
  y filas en cualquier programa y no tapa el texto de abajo. No agrandar el libro: ya vienen recomprimidas.
* Hoja: `showGridLines = False`, `page_setup.paperSize = 9` (A4), orientación según `landscape`,
  `fitToWidth = 1`, `fitToHeight = 0` (`sheet_properties.pageSetUpPr.fitToPage = True`), márgenes ~1 cm,
  `print_area` = toda la grilla. Sin autofiltro, sin panes inmovilizados.
* El libro conserva (si corresponde) la hoja `Origen` oculta con diagnóstico y la hoja `Revisión` NO se
  genera en este modo (no hay celdas "para revisar": lo dudoso se deja como texto).

## 5. Elección del modo (`core/mode.py`)

`suggest_mode` recibe lo que `analyze()` ya sabe (páginas, tablas crudas, `LayoutPage`s, `ocr_applied`) y
devuelve `("layout"|"tables", frase_en_castellano)`. Regla inicial (ajustable con el corpus):

1. Sin ninguna tabla y sin cajas dibujadas (solo texto corrido) → `tables` (hoja "Texto", como antes).
2. Hubo **merge multipágina** (una tabla continúa en otra página) o hay una tabla dominante que cubre ≥ 85 % de
   las palabras visibles de la mayoría de las páginas (listados, extractos, reportes) → `tables`.
3. Documento escaneado (`ocr_applied`) con tablas detectadas → `tables` (sin bordes confiables, la grilla
   sería ruido).
4. En cualquier otro caso (≥ 2 tablas/cajas por página, mucho contenido fuera de tablas, formularios) →
   `layout`.

La frase explica en castellano llano: "Parece un formulario (varias cajas y datos sueltos): conviene
respetar el diseño del PDF." / "Parece un listado: conviene una tabla limpia." La UI la muestra bajo el selector.

## 6. Integración

* `analyze()` calcula `layout_pages` en una etapa nueva `"layout"` (después de `"tables"`, antes de
  `"typing"`), página por página con chequeo de cancelación y progreso. Para documentos de más de **40
  páginas** solo se calcula si `Options.layout == "layout"` (no se gasta tiempo en listados grandes) y
  `suggested_mode` queda en `tables`.
* `export()`: `report.mode == "layout"` → `build_workbook` agrega las hojas de `write_layout_sheets` **en vez
  de** las hojas de tablas/texto/revisión; `"tables"` → comportamiento actual sin cambios. `export_batch` y
  `single_workbook` soportan reportes de modos mezclados.
* Los textos tapados se filtran **también en el modo tablas** (`visible_words` en `cells.words_in_rect`,
  `text_blocks`, candidatos): el bug de los "N/A" fantasma existía ahí.
* CLI: `--layout {auto,layout,tables}` (alias `--igual-al-pdf` / `--solo-tablas`), `--no-images`.
* UI: ver §8.

## 7. Corpus y medición

`tests/corpus/generate_forms.py` genera PDFs sintéticos de formularios (nombres **ficticios**, nunca datos
reales de personas ni empresas): hoja de ruta (réplica estructural del caso real, con texto tapado por
rectángulos blancos, logo y "mapa" raster), remito con cajas de encabezado y pie de firmas, factura con bloques
y totales, formulario de dos columnas de campos, tabla con sección combinada, página rotada, texto sin bordes.

`tests/layout_score.py` mide fidelidad **renderizando el .xlsx con LibreOffice** (`soffice --headless
--convert-to pdf`; si no está instalado, los tests de fidelidad se saltan) y comparando contra el PDF
original: recall de palabras visibles, palabras sobrantes (por ejemplo, texto que estaba tapado), error de
posición normalizado por el tamaño de la página, cantidad de cajas con bordes, y celdas que mezclan etiquetas
de cajas distintas. Umbrales iniciales: recall ≥ 0.98, sobrantes = 0, error de posición medio ≤ 2.5 % del
ancho de página. Con `--png` guarda lado a lado "PDF original | Excel renderizado" para inspección visual.

**Verificación con documentos reales.** Los PDF reales de los usuarios son privados: se usan solo en la
máquina local para comparar visualmente y **nunca** se copian al repo ni a los tests versionados (el repo es
público). Los tests usan solo el corpus sintético con datos ficticios.

## 8. UI

* Pantalla "Esto encontré": selector compacto **Formato del Excel: ◉ Igual al PDF · ○ Solo tablas de datos**,
  preseleccionado con `report.suggested_mode` (o con la preferencia fija del usuario), con la frase
  `mode_reason` debajo en gris. Cambiarlo actualiza `report.mode` y la vista previa.
* Vista previa en modo layout: widget `LayoutPreview` (QTableWidget con `setSpan`, fuentes, colores, rellenos y
  bordes aproximados, imágenes como miniatura) con un selector de página; en modo tablas, la tabla actual.
* Ajustes avanzados: "Formato del Excel: Automático (recomendado) / Siempre igual al PDF / Siempre solo
  tablas" → `Options.layout`.
* La etapa `"layout"` se muestra como "Armando el diseño de la hoja…".
* Textos en `ui/strings_es.py`, rioplatense.

## 9. No-objetivos

Reproducir fuentes exactas, degradados, curvas o dibujos vectoriales arbitrarios (un logo vectorial se
rasteriza como imagen si PyMuPDF lo expone como tal; si no, se omite), campos de formulario AcroForm
interactivos (se tratan como texto), edición dentro de la app.

## 10. Cambios de la ronda de fidelidad visual (R1)

Cambio en integración: no se tocó `models.py`.

* **Writer**: grilla exacta (sin `scale`), conversión de ancho de columna lineal (5.25 pt por unidad), fuente
  reducida por celda con métricas reales, imágenes con anclaje de dos esquinas. Medido sobre el corpus: error de
  posición mediano de 2.6-2.9 % a 0.5-1.5 %; el aviso amarillo de la página 2 de la hoja de ruta ya no queda
  tapado por el mapa.
* **Extractor**: reglas de 2-3.5 pt dibujadas con trazos "l" (línea azul bajo un título, línea de firma) se
  conservan; los tramos de borde colineales de distinto grosor ya no se pisan (se perdían los lados de un
  recuadro); los bordes de toda barra rellena son líneas de la grilla y el color cubre también los huecos sin
  texto (franja gris continua); tildes y cruces vectoriales de las casillas pasan a ✔ / ✗ en la celda; un valor
  escrito sobre una línea de subrayado es otra celda que su etiqueta; dos columnas numéricas alineadas a la
  derecha con poco hueco ya no se juntan; un texto suelto justo al ancho de su celda se amplía sobre pistas
  libres (evita "###" y habilita tipar números).
* **Tipado**: el formato de miles (`#,##0.00`) solo se usa si algún valor de la columna lo trae; si no, `0.00`.
  La columna de un número se define por su borde de alineación, no por el ancho de la celda.
* **Escaneos**: el texto del OCR toma un tamaño por renglón (el que mide lo mismo con Helvetica) y se descartan
  los restos de 1 pt; con "Igual al PDF" forzado sobre un escaneo, `mode_reason` avisa que no se reproducen cajas
  ni colores. El modo automático sigue mandándolos a "Solo tablas".
* **Medidor** (`tests/layout_score.py`): los números/fechas tipados no cuentan como celdas mezcladas, la cadena de
  palabras prefiere el mismo renglón y una celda solo se marca si NINGUNA aparición del texto es limpia; un PDF sin
  capa de texto no reporta "sobrantes".

Pendiente (no resuelto en R1): negrita/colores de un escaneo, marca de baja confianza del OCR por celda,
alineación vertical de cajas hermanas (vacías vs. con valor), bordes dobles en cajas cruzadas por una línea.

## 11. Cambios de la ronda de fidelidad visual (R2)

Cambio en integración: no se tocó `models.py`.

* **Writer, filas justas de alto de letra** (`balanced_row_heights`): el extractor mide una fila de texto con
  el alto de la letra (8 pt para 8 pt) pero Excel necesita ~1.2 ×; antes la letra caía al 70 % (totales del
  remito, formulario de dos columnas, títulos). Ahora la fila baja le pide alto prestado a las pistas de
  relleno vecinas (hasta 3 filas hacia cada lado, a la par para no mover el centro), sin cambiar el total ni
  correr nunca una línea dibujada; un subrayado crece solo hacia arriba y queda alineado abajo. Solo se corrigen
  filas "justas": un texto que se envuelve en más renglones que el PDF sigue achicando la letra de su celda.
* **Extractor**: una palabra precedida por un espacio real del PDF y a menos de 1 × tamaño no se corta como
  "columna" (título con letras espaciadas, "HOJA DE RUTA" salía en dos celdas y a 70 %).
* **Escaneos** (`core/layout_raster.py`, nuevo): sin vectores, las cajas, líneas y rellenos salen de los
  píxeles (OpenCV: líneas = "más oscuro que el entorno" con apertura morfológica larga, así una línea sobre un
  relleno gris no se pierde; rellenos = regiones lisas de color). Entran al extractor como los `Segment`/`Box` de
  un PDF digital, de modo que las celdas, los bordes, las columnas y el tipado por columna funcionan igual
  (antes: texto flotante, valores en columnas distintas por renglón, frases partidas). También se estima por
  palabra la negrita (grosor del trazo relativo al esperado para ese tamaño) y el color de tinta (etiquetas
  grises), el tamaño corrige que la caja del OCR sea ~10 % más angosta que la tinta, es uno por tramo de renglón,
  se une al más frecuente de la página (±12 %) y la negrita se mide con métricas de negrita.
* **Medidor**: un PDF sin vectores ni texto se compara contra el detector de píxeles (ya no "0 contra 0") y
  `problems()` exige que el Excel conserve al menos la mitad de los tramos de borde del PDF.

Pendiente tras R2: errores de lectura del OCR en letra chica y gris ("CUIT" → "curT"; no hay confianza por
celda en la capa de texto de tesseract, así que no se puede marcar), sangría interna (`indent`) de las celdas,
alineación vertical de cajas hermanas, bordes de lados compartidos sin deduplicar, imagen del escaneo como
referencia en otra hoja.

## 12. Cambios de la ronda de fidelidad visual (R3)

Cambio en integración: no se tocó `models.py`.

* **Alto de fila en píxeles** (`quantize_row_heights`, usado por `balanced_row_heights(..., pixel_snap=True)`):
  Excel y LibreOffice redondean el alto de cada fila a píxeles enteros a 96 ppp (0.75 pt; LibreOffice hacia abajo).
  Con 45 filas de 20 pt cada una quedaba en 19.5 y el contenido de abajo (firmas, pie) subía ~3 %. Ahora cada fila se
  escribe en un múltiplo exacto de 0.75 pt y el error se arrastra a la fila siguiente (la posición acumulada nunca se
  aparta más de un píxel); una fila cuyo texto entra no queda más baja de lo que necesita.
* **Letra al tamaño del PDF también con líneas arriba y abajo**: la casilla con borde en un renglón de 8.8 pt o un
  título de 15 pt sin aire caían al 70 %. Si las pistas vecinas no alcanzan (hay una línea en el medio), una segunda
  pasada pide el alto que falta a la pista **vacía** más cercana (hasta 60 % de su alto, nunca una pista con texto);
  lo que queda en el medio se corre unos puntos en bloque. El total de la hoja no cambia.
* **Texto de una línea que se derrama** (`overflow_layout`): un texto sin envolver puede ocupar las pistas libres de al
  lado (como hace Excel con celdas vecinas vacías) mientras ninguna celda del extractor las use y la celda no tenga
  borde de ese lado; solo así se decide si "entra". El número de página pegado al borde derecho ya no baja al 70 %:
  se alinea hacia donde hay lugar.
* **Escaneos**: los números cortos (`120`, `300`, hasta 4 dígitos) no definen el tamaño de su tramo (la caja del OCR
  de dígitos sueltos engaña) y toman el de la página, así la columna CANT. sale pareja. Las lecturas que huelen mal
  por sus mayúsculas (`curT` por `CUIT`: palabra de letras que arranca en minúscula con una mayúscula adentro) dejan el
  aviso `OCR_DOUBT_WARNING` en `LayoutPage.warnings` y el writer agrega una **nota** (comentario de Excel, triángulo
  rojo) en esa celda. No hay confianza por palabra del OCR; es una heurística conservadora.
* **Medido**: la compresión vertical que se ve al renderizar con LibreOffice (contenido de abajo ~7 % más arriba) tiene
  dos causas. La del redondeo de filas (arriba) se corrigió. La otra es del renderizador: LibreOffice interpreta una
  unidad de ancho de columna como ~7.4 px (dígito de Carlito sin redondear) y Excel como 7 px, así que en LibreOffice las
  columnas salen ~5.5 % más anchas, el ajuste a una página escala todo hacia abajo y las filas parecen más bajas. En
  Excel las columnas miden lo que dice el PDF y la proporción se conserva; no se compensa para no empeorar Excel.

Pendiente tras R3: sangría interna (`indent`) de las celdas y alineación vertical de cajas hermanas (el texto queda más
pegado al borde que en el PDF), lados compartidos entre cajas sin deduplicar (cada celda lleva su borde; no se nota),
esquinas redondeadas, importes de la caja de totales de un escaneo que quedan como texto cuando la celda es justa.
