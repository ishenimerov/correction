#!/usr/bin/env python3
"""
corrections.py

Reads highlighted (default: yellow) cells of one language column from a reviewed
Excel file and writes those values into a JSON translations file:
  { "key": { "en": "...", "lo": "...", ... } }

The Excel layout is the one json_to_excel.py writes: row 1 has "Key" plus one
column per language, row 2 has the language codes in brackets, e.g. "(lo)".

Usage:
    python corrections.py <translations.json> <reviewed.xlsx> [options]

Options:
    --lang CODE      Language to update in the JSON (default: lo)
    --column NAME    Excel column header to read (default: found from --lang)
    --color RGB      Highlight colour that marks a correction (default: FFFF00)
    --output FILE    Output JSON path (default: overwrites the input JSON)
    --dry-run        Show what would change without writing anything

For a browser version with upload, preview and download, open index.html (or the deployed site).

Requirements:
    python -m pip install openpyxl
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("Error: 'openpyxl' not found.\nInstall it with: python -m pip install openpyxl")
    sys.exit(1)

# Header names json_to_excel.py writes, used when row 2 has no "(code)" row.
LANG_NAMES = {
    "en": "English", "tr": "Turkish", "fr": "French", "ar": "Arabic",
    "el": "Greek", "ro": "Romanian", "ru": "Russian", "uz": "Uzbek",
    "sr_cyrl": "Serbian (Cyrillic)", "sr_Latn": "Serbian (Latin)",
    "sq": "Albanian", "es": "Spanish", "lo": "Lao", "pl": "Polish",
}


def find_columns(ws, lang: str, column: str | None) -> tuple[int, int, int]:
    """Return (key_col, value_col, first_data_row) as 0-based column indexes."""
    headers = [cell.value for cell in ws[1]]
    codes = [cell.value for cell in ws[2]]
    if "Key" not in headers:
        print("❌ Error: no 'Key' column in row 1 of the Excel file.")
        sys.exit(1)
    key_col = headers.index("Key")
    has_code_row = any(isinstance(c, str) and c.startswith("(") for c in codes)

    if column:
        candidates = [i for i, h in enumerate(headers) if h == column]
    else:
        # Prefer the exact code in row 2, then the display name / code in row 1.
        candidates = [i for i, c in enumerate(codes) if c == f"({lang})"]
        if not candidates:
            names = {LANG_NAMES.get(lang), lang}
            candidates = [i for i, h in enumerate(headers) if h in names]

    if not candidates:
        print(f"❌ Error: no Excel column found for '{column or lang}'. Columns: {headers}")
        print("   Pass the header name with --column.")
        sys.exit(1)

    return key_col, candidates[0], 3 if has_code_row else 2


def cell_rgb(cell) -> str | None:
    """Fill colour as 'RRGGBB', or None for no fill / theme or indexed colours."""
    if not cell.fill or not cell.fill.fill_type:
        return None
    rgb = cell.fill.fgColor.rgb
    return rgb[-6:].upper() if isinstance(rgb, str) else None


def extract_corrections(xlsx_path: Path, lang: str, column: str | None, color: str) -> dict[str, str]:
    """Extract all highlighted cells of the chosen language column."""
    wb = load_workbook(xlsx_path)
    ws = wb.active
    key_col, value_col, first_row = find_columns(ws, lang, column)
    print(f"🔎 Reading column '{ws[1][value_col].value}' — cells filled with #{color}")

    corrections = {}
    for row in ws.iter_rows(min_row=first_row):
        key_cell = row[key_col]
        value_cell = row[value_col]
        if key_cell.value and value_cell.value is not None and cell_rgb(value_cell) == color:
            corrections[str(key_cell.value)] = str(value_cell.value)

    return corrections


def apply_corrections(json_path: Path, corrections: dict[str, str], lang: str,
                      output_path: Path, dry_run: bool) -> None:
    # Read bytes so CRLF line endings survive (read_text would turn them into LF).
    raw = json_path.read_bytes().decode("utf-8")
    bom = raw.startswith("﻿")
    if bom:
        raw = raw[1:]
    data: dict = json.loads(raw)

    updated, added, unchanged, not_found = 0, 0, 0, []
    for key, new_value in corrections.items():
        if key not in data:
            not_found.append(key)
        elif data[key].get(lang) == new_value:
            unchanged += 1
        else:
            if lang in data[key]:
                updated += 1
            else:
                added += 1
            data[key][lang] = new_value

    print(f"✏️  Updated: {updated}   ➕ Added: {added}   = Already the same: {unchanged}")
    if not_found:
        print(f"⚠️  {len(not_found)} keys from Excel not found in JSON:")
        for k in not_found:
            print(f"   • {k}")

    if dry_run:
        print("\n[Dry run] Nothing written.")
        return

    # Keep the input's line endings and trailing newline so the diff stays small.
    newline = "\r\n" if "\r\n" in raw else "\n"
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if newline != "\n":
        text = text.replace("\n", newline)
    if raw.endswith(("\n", "\r\n")):
        text += newline
    if bom:
        text = "﻿" + text
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        f.write(text)

    print(f"✅ Saved → {output_path}")


def main():
    # Windows consoles / pipes may default to cp1252, which can't print the emoji.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Apply highlighted Excel corrections to a JSON translations file."
    )
    parser.add_argument("json", help="Path to the JSON translations file")
    parser.add_argument("xlsx", help="Path to the reviewed Excel file")
    parser.add_argument("--lang", default="lo",
                        help="Language code to update in the JSON (default: lo)")
    parser.add_argument("--column", default=None,
                        help="Excel column header to read (default: found from --lang)")
    parser.add_argument("--color", default="FFFF00",
                        help="Highlight colour marking a correction, RRGGBB (default: FFFF00)")
    parser.add_argument("--output", default=None,
                        help="Output JSON path (default: overwrites input)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would change without writing")
    args = parser.parse_args()

    json_path = Path(args.json)
    xlsx_path = Path(args.xlsx)
    output_path = Path(args.output) if args.output else json_path
    color = args.color.lstrip("#").upper()[-6:]

    for p in [json_path, xlsx_path]:
        if not p.exists():
            print(f"❌ Error: File not found: {p}")
            sys.exit(1)

    print(f"📂 Reading Excel: {xlsx_path}")
    corrections = extract_corrections(xlsx_path, args.lang, args.column, color)
    print(f"🟡 Found {len(corrections)} highlighted corrections")

    print(f"📂 Reading JSON:  {json_path}")
    apply_corrections(json_path, corrections, args.lang, output_path, args.dry_run)


if __name__ == "__main__":
    main()
