# 📄➡️📊 Conversor de PDF a Excel (Google Colab)

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/juanmaserrate/conversor-pdf-a-excel/blob/main/Conversor_PDF_a_Excel.ipynb)

Convierte PDFs a **Excel**, **CSV** o **JSON** desde el navegador, sin instalar nada.

## Cómo usarlo

1. Tocá el botón **Abrir en Colab** de arriba.
2. En Colab: **Entorno de ejecución → Ejecutar todas** (`Ctrl+F9`).
   - Si aparece el aviso *"Este notebook no fue creado por Google"*, tocá **Ejecutar de todos modos**.
3. Elegí el modo y el formato.
4. Tocá **Elegir archivos**, subí uno o varios PDFs y el resultado se descarga solo
   (varios PDFs → un `.zip`).

### Modos

| Modo | Qué hace |
|---|---|
| **Todo el texto + tablas** | Pone todo el contenido del PDF en una sola hoja. |
| **Solo tablas** | Cada tabla detectada va en una hoja separada. |

## Privacidad

Cada persona ejecuta el notebook en su propia máquina temporal de Google.
Los PDFs no se comparten con nadie más y se borran al cerrar la sesión.

## Créditos

Usa el código de [TsvetanG2/PDF-To-Excel-Converter](https://github.com/TsvetanG2/PDF-To-Excel-Converter)
(licencia MIT), que se descarga automáticamente al ejecutar el notebook.
