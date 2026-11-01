# Corpus sintético de PDFs

Corpus determinista para medir la calidad de extracción (DESIGN.md §9). Los PDFs **no se commitean**:
se regeneran con una semilla fija (dos corridas dan los mismos bytes, salvo `con_password.pdf`, cuyo
cifrado AES-256 usa sales aleatorias). Lo único versionado acá es `generate.py`, `ground_truth.json`
y este README.

```bash
cd desktop
python -m tests.corpus.generate tests/corpus/_generated            # ≈ 5 s, 24 PDFs
python -m tests.corpus.generate /tmp/out --ground-truth tests/corpus/ground_truth.json   # también reescribe el GT
pytest tests/test_corpus_accuracy.py tests/test_quality_selection.py -q            # digitales (≈ 45 s)
pytest ... --runslow --corpus-report corpus_report.md                              # + escaneos con OCR y tabla markdown
```

`conftest.py` usa `tests/corpus/_generated/` si está completo (así lo deja CI); si no, genera el corpus
una vez por sesión en un directorio temporal. `PDF2EXCEL_CORPUS_DIR=/ruta` fuerza otra carpeta ya generada.

## Casos

| PDF | Qué prueba | Páginas |
|---|---|---|
| `remito_grid`, `remito_nogrid`, `prueba` | 7×5 con/sin bordes, 3×3 mínima; importes `350.00` (US) | 1 |
| `hard_factura_ar` | repeatRows, wrap, `$ 1.234,56`, Subtotal/IVA/TOTAL con SPAN sin grilla, encabezado y pie | 2 |
| `hard_extracto_banco` | sin bordes, 70 movimientos, negativos, Débito/Crédito mutuamente vacías, `dd/mm/aa`, "Saldo final:" suelto | 2 |
| `hard_lista_mixta` | 2 tablas en una página (grilla gris + línea bajo el encabezado), `s/d`, fechas | 1 |
| `cont_sin_header` | 110 filas que siguen en 3 páginas **sin** repetir encabezado (merge Regla 2) | 3 |
| `header_dos_filas` | encabezado de 2 filas con celda combinada arriba | 1 |
| `apaisado` | landscape, 12 columnas, enteros `1.234` | 1 |
| `porcentajes` | `21%`, `10,5 %` → fracción | 1 |
| `numeros_us` | todas las columnas numéricas `1,234.56` / `350.00` | 1 |
| `negativos_banco` | `1.234,56-` y `(1.234,56)` | 1 |
| `con_password` | AES-256, contraseña `1234` (sin ella → `UserError("encrypted")`) | 1 |
| `sin_tablas` | 3 párrafos: el Excel lleva solo la hoja "Texto" | 1 |
| `vacio` | página en blanco → `kind == "empty"` | 1 |
| `mixto` | pág. 1 digital con tabla + pág. 2 solo imagen (carta sin tablas) | 2 |
| `texto_basura` | capa de texto con > 30 % de caracteres raros (fuente sin ToUnicode) → `garbage_text` | 1 |
| `*_scan` | `make_scan`: render 200 dpi gris, rotación ±1,2°, blur 0,6, ruido, JPEG q70, sin capa de texto | = origen |
| `hard_factura_ar_scan90` | ídem, girado 90° (OSD) | 2 |
| `remito_grid_scan150` | ídem a 150 dpi | 1 |

Los `_scan*` heredan el ground truth del original (`scorer.gt_for` quita los sufijos `_scan`, `_scan90`,
`_scan150` y `_ocr`).

## Ground truth

`ground_truth.json`: una entrada por PDF digital.

```json
"remito_grid.pdf": {
  "tables": [{"header": ["Código", "Producto", "Cantidad", "Precio unitario", "Total"],
              "rows": [[{"v": "A-100", "t": "str"}, {"v": "Mezcal joven", "t": "str"},
                        {"v": 10, "t": "int"}, {"v": 350.0, "t": "float"}, {"v": 3500.0, "t": "float"}], ...]}],
  "expected_tables": 1,
  "kinds": ["digital"],
  "notes": "..."
}
```

- Celdas: `{"v": valor, "t": "float" | "int" | "str" | "date"}` o `null` (vacía). Las fechas van en ISO.
- El GT describe lo que **debería** salir en el Excel, no lo que sale hoy: p. ej. en `header_dos_filas` el
  encabezado es el combinado ("Precio ($) Lista"), y en `hard_lista_mixta` la celda `s/d` es texto (así
  está impresa).
- Las filas de total (Subtotal/IVA/TOTAL, "Saldo final:") forman parte de la tabla, con `null` en las
  celdas que el PDF deja vacías.
- `tests/test_corpus_accuracy.py::test_ground_truth_json_in_sync` falla si el JSON no coincide con lo que
  produce el generador: después de tocar `generate.py`, regenerar el JSON con `--ground-truth`.

## Métricas y umbrales (`tests/scorer.py`, `tests/test_corpus_accuracy.py`)

- `cells_pct`: por fila GT (encabezado incluido), la mejor fila extraída; la celda cuenta si coincide tras
  normalizar (números a 2 decimales, fechas ISO, texto en minúsculas con espacios colapsados; `None == ""`).
  Un texto que parsea al número/fecha del GT cuenta como contenido correcto; el tipo lo mide `type_pct`.
- `rows_pct`: filas GT con match exacto completo. `spurious_rows_pct`: filas extraídas que no son la mejor
  de ninguna fila GT (ruido de encabezados de página, párrafos absorbidos). `spurious_cols`: columnas de más.
- `type_pct`: celdas GT float/int/date cuyo valor extraído es número o fecha real (no `str`).
- Umbrales que rompen el build: digitales → celdas ≥ 99, tablas detectadas = esperadas, espurias ≤ 1,
  tipos ≥ 99; casos nuevos (`header_dos_filas`, `apaisado`, `porcentajes`, `numeros_us`, `negativos_banco`)
  → celdas ≥ 95 (subir a 99 cuando pasen); escaneos → OCR aplicado y celdas ≥ 92 (factura, remitos,
  continuación), ≥ 88 (lista mixta), ≥ 90 (`scan90`), ≥ 85 (`scan150`).
- Excepciones documentadas en `KNOWN_SPURIOUS_MAX` (limitaciones del core, no del corpus): hoy solo
  `hard_lista_mixta` (extend_tail absorbe el párrafo que sigue a la tabla 1). El reporte las muestra igual.

`test_quality_selection.py` verifica, página por página, que el candidato que elige `select_best` (sin GT)
tiene `cells_pct` ≥ máximo − 0,5 entre `hybrid_union` y `hybrid`.

## Cómo agregar un caso

1. En `generate.py`, escribí `build_<nombre>(path: Path, rng: random.Random) -> dict` que dibuje el PDF
   con reportlab (`invariant=1`, ya lo hacen `_doc`/`_canvas`) y devuelva `gt_entry(tables, kinds, notes)`
   usando `gt_table`, `gt_float`, `gt_int`, `gt_str`, `gt_date`. Todo valor aleatorio sale de `rng`.
2. Agregalo **al final** de `CASES` (el RNG es compartido: insertarlo en el medio cambia los datos de los
   casos siguientes y, por lo tanto, su GT). Si querés su versión escaneada, agregá un `ScanSpec` a `SCANS`.
3. Regenerá el GT: `python -m tests.corpus.generate /tmp/out --ground-truth tests/corpus/ground_truth.json`
   y revisá el diff del JSON (solo debería aparecer tu entrada).
4. Sumalo a la lista que corresponda en `tests/test_corpus_accuracy.py` (`STRICT_DIGITAL`,
   `RELAXED_DIGITAL`, `SCAN_MIN_CELLS`) o escribile un test propio si prueba otra cosa (como `vacio`).
5. `pytest tests/test_corpus_accuracy.py tests/test_quality_selection.py --runslow`.

## PDFs reales del equipo: `tests/corpus_real/`

Carpeta ignorada por git (`.gitignore`) para PDFs reales **anonimizados** del equipo, con su propio
`ground_truth.json` en el mismo formato. Sirve para medir en local lo que el corpus sintético no cubre
(fotos de celular, plantillas propias) sin publicar datos:

```bash
mkdir -p tests/corpus_real && cp ~/Descargas/remito_real.pdf tests/corpus_real/
# escribí tests/corpus_real/ground_truth.json con la entrada "remito_real.pdf" (ver formato arriba)
python - <<'EOF'
import json, sys
from pathlib import Path
sys.path.insert(0, ".")
from tests.scorer import score, mismatches, format_table
from pdf2excel.core.pipeline import convert
from pdf2excel.models import Options
root = Path("tests/corpus_real"); gt = json.loads((root / "ground_truth.json").read_text(encoding="utf-8"))
results = {}
for name, entry in gt.items():
    report = convert(root / name, Options(output_dir=root / "out", open_when_done=False), password=entry.get("password"))
    results[name] = score(report.tables, entry)
    print(name, results[name], *mismatches(report.tables, entry, 5), sep="\n  ")
print(format_table(results))
EOF
```

Cuando un PDF real falla, lo ideal es reproducir el problema en un caso sintético nuevo (paso anterior):
queda en CI sin exponer el documento original.
