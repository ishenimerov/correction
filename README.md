# Translation Corrections

Applies reviewed translations from an Excel file to a JSON translations file.
Reviewers highlight corrected cells (yellow by default); only those cells are written back.

JSON format:

```json
{ "messages_title": { "en": "List of created Content", "lo": "…" } }
```

Excel layout (as exported from the JSON): row 1 has `Key` plus one column per language
(`English`, `Lao`, …), row 2 has the language codes in brackets (`(en)`, `(lo)`, …).
Row 2 is optional.

## Web app (`index.html`)

A static page, no build step. Files are processed in the browser and never uploaded anywhere.

1. Upload the translations JSON and the reviewed Excel file.
2. Pick the Excel column and the JSON language code (filled in from the column).
3. Pick which highlight colours count as corrections. The page lists every fill colour it found in that column, with counts.
4. Review the result. Rows are grouped as **Will change**, **Will be added**, **Key not in JSON** and **Already the same**. Untick any row you don't want.
5. Download the JSON. It keeps the original key order, indentation, line endings and BOM, so the diff only shows the changed values.

Run locally: open `index.html` in a browser. It loads [ExcelJS](https://github.com/exceljs/exceljs) from jsDelivr, so it needs an internet connection.

### Deploy to Vercel

Import the GitHub repo in Vercel, keep **Framework Preset: Other**, leave the build command empty, and set the output directory to the repo root. Vercel serves `index.html` as the site.

## Command line (`cli/corrections.py`)

```bash
python -m pip install openpyxl
python cli/corrections.py translations.json reviewed.xlsx --lang lo --output translations_updated.json
```

| Option | Default | Meaning |
|---|---|---|
| `--lang CODE` | `lo` | Language to update in the JSON |
| `--column NAME` | from `--lang` | Excel column header to read |
| `--color RGB` | `FFFF00` | Fill colour that marks a correction |
| `--output FILE` | input JSON | Where to write; without it the input JSON is overwritten |
| `--dry-run` | off | Show the counts without writing |

The CLI and the web app produce byte-identical output for the same inputs and settings.
The CLI only matches plain RGB fills; for theme or indexed colours use the web app.
