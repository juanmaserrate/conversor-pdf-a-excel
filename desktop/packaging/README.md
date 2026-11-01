# Empaquetado, instalador y releases

Todo lo que hay que saber para convertir `desktop/` en un `ConversorPDFaExcel-X.Y.Z-setup-win64.exe`
(y en un `.zip` portable y un `.dmg` para macOS). El camino normal es **no hacer nada a mano**: un
`git tag vX.Y.Z` dispara el workflow `.github/workflows/build-desktop.yml`, que testea, construye, firma
(si hay certificado) y publica el release. Esta guía explica qué hace cada pieza y cómo reproducirlo en
una PC si hace falta.

| Archivo | Para qué |
|---|---|
| `conversor.spec` | Receta de PyInstaller (onedir). Entry `src/pdf2excel/__main__.py`. |
| `prune_qt.py` | Post-build: borra módulos Qt, plugins y traducciones que no se usan (−31 MB medidos). |
| `fetch_tesseract.py` | Baja `tesseract.exe` (Windows) y `tessdata_fast` spa/eng/osd con sha256 verificado. |
| `installer.iss` | Instalador Inno Setup 6 por usuario (sin UAC), con "Enviar a". |
| `file_version_info.txt` | Propiedades del `.exe` (versión, descripción en castellano). |
| `stamp_version.py` | Escribe la versión del tag en `__init__.py`, `installer.iss` y `file_version_info.txt`. |
| `../src/pdf2excel/update_check.py` | La app lee `latest.json` del último release y avisa si hay versión nueva. |

---

## 1. Construir localmente en Windows (pasos exactos)

Requisitos: Windows 10/11 x64, [Python 3.13 x64](https://www.python.org/downloads/windows/) (tildar
"Add python.exe to PATH"), [Git](https://git-scm.com/download/win), [7-Zip](https://www.7-zip.org/)
(para extraer el instalador de Tesseract y armar el zip portable) e
[Inno Setup 6](https://jrsoftware.org/isdl.php) (para el instalador). Abrí **PowerShell**:

```powershell
git clone https://github.com/juanmaserrate/conversor-pdf-a-excel.git
cd conversor-pdf-a-excel\desktop

py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e ".[dev]"                      # app + pytest + ruff + mypy + pyinstaller

# 1) Binarios de OCR (una sola vez; quedan en resources\ y están en .gitignore)
python packaging\fetch_tesseract.py resources

# 2) (opcional) versión: solo si vas a publicar; en desarrollo queda la de src\pdf2excel\__init__.py
python packaging\stamp_version.py v0.1.0

# 3) Bundle onedir -> dist\ConversorPDFaExcel\
pyinstaller --noconfirm --clean packaging\conversor.spec

# 4) Podar Qt (imprime los MB liberados)
python packaging\prune_qt.py dist\ConversorPDFaExcel

# 5) Smoke: tiene que imprimir "selftest ok" (usa el OCR y el modelo de layout DENTRO del bundle)
python -m tests.corpus.generate tests\corpus\_generated
.\dist\ConversorPDFaExcel\ConversorPDFaExcel.exe --selftest tests\corpus\_generated\remito_grid_scan.pdf > selftest.log
type selftest.log

# 6) Instalador -> dist\installer\ConversorPDFaExcel-0.1.0-setup-win64.exe
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" packaging\installer.iss

# 7) Zip portable (alternativa sin instalar nada)
& "C:\Program Files\7-Zip\7z.exe" a -mx=7 dist\installer\ConversorPDFaExcel-0.1.0-portable-win64.zip .\dist\ConversorPDFaExcel
```

Notas:

- El `.exe` tiene `console=False` (no abre ventana negra). Por eso `--selftest` se redirige a un archivo:
  en una consola de PowerShell el texto puede no verse, pero con `> selftest.log` siempre queda escrito.
- Si PyInstaller falla con `FileNotFoundError ... pymupdf/layout/resources/onnx/*.yaml` al correr el
  bundle, es que falta `collect_data_files("pymupdf")` en el `.spec`: es **obligatorio**, no lo saques.
- Si en el venv quedó instalado `opencv-python` (el completo) además de `opencv-python-headless`, el
  `.spec` avisa: el bundle crece ~116 MB. Hacé `pip uninstall opencv-python`.
- `--selftest` necesita `tesseract`: en Windows el vendoreado en `resources\tesseract\`; en macOS/Linux
  el del sistema (`brew install tesseract tesseract-lang` / `apt install tesseract-ocr tesseract-ocr-spa`).

### macOS / Linux

Mismos pasos sin Inno Setup: `pyinstaller --noconfirm --clean packaging/conversor.spec`, luego
`python packaging/prune_qt.py dist/ConversorPDFaExcel.app` (macOS) o `dist/ConversorPDFaExcel` (Linux) y
`hdiutil create -volname "Conversor PDF a Excel" -srcfolder dist/ConversorPDFaExcel.app -ov -format UDZO dist/ConversorPDFaExcel-mac.dmg`.
En Linux el smoke se corre con `QT_QPA_PLATFORM=offscreen dist/ConversorPDFaExcel/ConversorPDFaExcel --selftest archivo.pdf`.

---

## 2. Tamaño (medido)

Linux x64, Python 3.13, PySide6-Essentials 6.12, pymupdf 2.0 + layout, onnxruntime 1.31.1, numpy 2.5,
tessdata_fast incluida (17 MB), `optimize=1`, sin UPX:

| Etapa | Tamaño onedir |
|---|---|
| `pyinstaller` recién salido (venv de desarrollo con `opencv-python` completo) | 557 MB |
| después de `prune_qt.py` (−31,4 MB) | 526 MB |
| ídem descontando `opencv_python.libs` (Qt5+GTK+ffmpeg que trae el opencv NO headless; CI no lo instala) | **≈ 410 MB** |

Desglose principal tras la poda: pymupdf 105 MB (51 MB son los modelos ONNX del layout), PySide6 89 MB,
cv2 71 MB, onnxruntime 30 MB, numpy 41 MB, tessdata 17 MB, Python/stdlib ≈ 20 MB. En Windows se suman
tesseract.exe + DLL (≈ 45 MB) y no están las libs de GTK/X11 (≈ 20 MB): la estimación de DESIGN §8
(330–360 MB onedir, instalador LZMA2 120–140 MB) sigue siendo razonable; el tamaño real lo imprime el
job `build-windows` en el *summary* de cada corrida. Para verificarlo a mano: `du -sh dist/ConversorPDFaExcel`
(Linux/macOS) o `(Get-ChildItem dist\ConversorPDFaExcel -Recurse | Measure-Object Length -Sum).Sum / 1MB`
(PowerShell).

Qué más se podría recortar si molesta: `pymupdf/layout/resources/onnx/` trae varias versiones del modelo
(solo se usa `layout_rf2.4.1+imf1` + `table_grid_model_v*`; hay que medir cuáles carga antes de borrar);
opencv-headless → apertura morfológica en numpy (−70 MB, plan B de DESIGN §10).

---

## 3. Firmar (Windows y macOS)

Sin firma la app funciona igual; la firma solo evita los avisos de SmartScreen/Gatekeeper.

**Windows (Authenticode)**: hace falta un certificado de firma de código OV/EV (≈ USD 200/año) o
Azure Trusted Signing. En CI se usa si existen los secrets `WIN_CERT_PFX_B64` (el `.pfx` en base64:
`[Convert]::ToBase64String([IO.File]::ReadAllBytes("cert.pfx")) | Set-Clipboard`) y `WIN_CERT_PASSWORD`.
Se firma el `.exe` del bundle **antes** de ISCC y el setup **después**, con timestamp de DigiCert. A mano:

```powershell
$st = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe" | Select-Object -Last 1
& $st sign /f cert.pfx /p "clave" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 dist\ConversorPDFaExcel\ConversorPDFaExcel.exe
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" packaging\installer.iss
& $st sign /f cert.pfx /p "clave" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 dist\installer\*.exe
```

**macOS**: secrets `APPLE_CERT_P12_B64` + `APPLE_CERT_PASSWORD` (Developer ID Application) para
`codesign --deep --options runtime --timestamp`, y `APPLE_ID` + `APPLE_TEAM_ID` + `APPLE_APP_PASSWORD`
(contraseña de app) para `xcrun notarytool submit --wait` + `stapler`. Sin ellos el `.dmg` se publica sin
firmar y el usuario tiene que abrir la app con clic derecho → Abrir la primera vez.

### Si Windows muestra "Windows protegió tu PC" (SmartScreen)

Pasa con cualquier programa nuevo sin firma (o con firma nueva sin reputación). No es un virus:

1. Tocá **Más información** (link chico debajo del texto).
2. Tocá **Ejecutar de todas formas**.

Aparece solo la primera vez por versión. Para evitarlo del todo: certificado OV/EV o distribuir el
instalador por la red interna / un recurso compartido (SmartScreen marca lo que viene de Internet).

---

## 4. Publicar un release

```bash
git checkout main && git pull
# (opcional) actualizar CHANGELOG / README
git tag v0.1.0
git push --tags
```

Eso dispara `build-desktop.yml`:

1. `test` (ubuntu): ruff + mypy + corpus sintético + pytest completo (`--runslow`). Si falla, no se construye nada.
2. `build-windows`: `fetch_tesseract.py` (con caché), `stamp_version.py v0.1.0`, PyInstaller, `prune_qt.py`,
   smoke `--selftest` sobre un PDF escaneado, firma (si hay secrets), ISCC, firma del setup, zip portable.
3. `build-macos`: `.app` + `.dmg` (+ codesign/notarización si hay secrets).
4. `release`: genera `latest.json` (`{"version","url","notes_url"}`) y crea el GitHub Release con el
   setup, el zip, el dmg y `latest.json`, con notas automáticas.

La app consulta `https://github.com/<owner>/conversor-pdf-a-excel/releases/latest/download/latest.json`
(`update_check.py`, timeout 3 s, falla en silencio) y muestra "Hay una versión nueva" con el link al setup.
Actualizar = correr el setup nuevo encima (misma `AppId`; `[InstallDelete]` limpia `_internal` viejo).

Si el tag estaba mal: `git tag -d v0.1.0 && git push --delete origin v0.1.0`, borrar el release en GitHub
y volver a taggear. Los pushes a `main` también construyen (sin release) y dejan los artefactos 14 días
en la pestaña *Actions*, útil para probar antes de publicar.

---

## 5. Actualizar Tesseract / tessdata (constantes de `fetch_tesseract.py`)

Las URLs y los sha256 están **fijos** en el script para que un build sea reproducible y nadie pueda
colar un binario distinto. Para cambiar de versión:

1. Editá `TESSERACT_WIN_VERSION` / `TESSERACT_WIN_URL` (instalador NSIS de 64 bits del build
   UB-Mannheim; se publica en <https://github.com/UB-Mannheim/tesseract/wiki>) y/o
   `TESSDATA_FAST_TAG` (`tessdata_fast`) y `TESSERACT_SRC_TAG` (`pdf.ttf` y `configs/pdf`).
   Ojo: solo algunas versiones están como *asset* en GitHub Releases (5.4.0.20240606 sí; 5.3.3 y 5.3.4
   dan 404 y viven únicamente en <https://digi.bib.uni-mannheim.de/tesseract/>). Probá la URL con
   `curl -IL` antes de fijarla. Sin tocar código, el workflow también acepta las variables del repo
   `TESSERACT_WIN_URL` / `TESSERACT_WIN_SHA256` (Settings → Secrets and variables → Variables).
2. Obtené los hashes reales: `python packaging/fetch_tesseract.py resources --print-hashes --platform win32`
   (baja todo a un temporal y los imprime). Compará el del instalador con el que publica UB-Mannheim.
3. Pegá los hashes en `TESSDATA_DOWNLOADS` / `TESSERACT_WIN_SHA256` y corré `python packaging/fetch_tesseract.py resources`
   (tiene que terminar en `ok` para cada archivo). `--dry-run` muestra la tabla URL/sha256 sin descargar.
4. Commit. El job de Windows invalida solo su caché (la clave es el hash del script).

Estado actual: los cinco archivos de `tessdata/` (spa, eng, osd de `tessdata_fast` 4.1.0; `pdf.ttf` y
`configs/pdf` de tesseract 5.3.4, necesarios para `-c textonly_pdf=1 pdf`) tienen sha256 verificados
(coinciden con los paquetes `tesseract-ocr-*` de Debian/Ubuntu). El sha256 del instalador de Windows está
como `PENDIENTE`: la máquina donde se escribió esto no llegaba a github.com; completarlo con el paso 2
antes del primer build de Windows (el script falla con un mensaje claro hasta entonces). Como atajo sin
tocar código, el job acepta `PDF2EXCEL_TESSERACT_WIN_URL` y `PDF2EXCEL_TESSERACT_WIN_SHA256` como
variables de entorno.

---

## 6. Qué hace `prune_qt.py` (y qué nunca toca)

Borra: bibliotecas y módulos PySide6 de Qml/Quick/VirtualKeyboard/Pdf/WebEngine/Multimedia/3D/Charts/Sql/
Xml/Test/… (regex `LIB_DROP`), directorios de plugins no usados (`tls`, `sqldrivers`, `qmltooling`,
`wayland-*`, `platforminputcontexts`…), traducciones que no sean `qt*_es` / `qt*_en`, el directorio `qml/`
y las herramientas de desarrollo (`designer`, `lupdate`…), y en Windows `opengl32sw.dll`,
`d3dcompiler_47.dll` y `opencv_videoio_ffmpeg*.dll`. **Nunca** toca `Qt6Core/Gui/Widgets/Network/Svg/
OpenGL/DBus/XcbQpa`, ni `platforms/`, `imageformats/`, `iconengines/`, `styles/`, `platformthemes/`, ni
nada fuera del árbol de Qt. Es idempotente (la segunda corrida libera 0 MB). En macOS acepta el `.app`
y limpia los symlinks rotos de `Contents/Resources`.

---

## 7. Licencia (AGPL-3.0)

`desktop/` se distribuye bajo **GNU AGPL-3.0 o posterior** porque PyMuPDF y pymupdf-layout (Artifex)
son AGPL. Consecuencias prácticas para quien distribuye el instalador:

- El código fuente completo de la versión distribuida tiene que estar disponible bajo la misma licencia:
  lo está, en este repo público; el instalador copia `LICENSE.txt` junto a la app y la app muestra el
  link al repo en "Acerca de".
- Si se modifica la app y se distribuye (incluso solo dentro de una organización, o se ofrece como
  servicio por red), hay que publicar esas modificaciones bajo AGPL.
- Si se quisiera cerrar el código, hace falta una licencia comercial de Artifex para PyMuPDF.
- Tesseract (Apache-2.0), Qt/PySide6 (LGPL-3.0, enlazado dinámicamente: los `.dll`/`.so` de Qt van
  sueltos en `_internal/`, que es lo que la LGPL exige), onnxruntime (MIT), OpenCV (Apache-2.0), numpy
  (BSD), openpyxl (MIT) son compatibles con esta distribución. No se incluye Ghostscript (evitamos una
  segunda AGPL).
