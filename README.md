# kkr-csv-format-discover

Small Windows-friendly helper that **detects a CSV “format profile”** from a file and helps you reuse it for consistent exports (especially when you later generate CSVs with `kkr-query2xlsx` or similar tooling).

It focuses on **format, not data** - delimiter, encoding, decimal separator, line endings, quoting, date format, and basic column structure.

---

## Why it exists

CSV files coming from different sources often look “similar” but break pipelines because of tiny differences:
- `;` vs `,`
- `0,00` vs `0.00`
- `\r\n` vs `\n`
- UTF-8 with/without BOM
- quoting sometimes present, sometimes not
- date strings in slightly different patterns

This tool reads a CSV and gives you a **ready-to-paste “format description”** + quick diagnostics and visual checks.

---

## Features

### Format detection (single file)
- Encoding detection (with BOM diagnostic)
- Field delimiter detection (typical `; , \t |`)
- Decimal separator heuristic (`,` or `.`)
- Line terminator detection (`\n`, `\r\n`, `\r`)
- Quoting mode detection + “quote char observed” diagnostic
- Basic date format inference (best-effort)

### Compare mode (two files, format-only)
- Compare detected format fields (A vs B)
- Show aligned sample rows **by column name**:
  - same columns are under each other
  - missing column in one file -> empty cell (visual alignment)
  - one horizontal scroll bar for header + A row + B row
- Select which row to preview for file A and file B (1-based)

### Copy helpers
- Copy full “format description” to clipboard (includes format fields and column samples)
- Copy date format (A or B) directly from compare view

---

## Privacy / data safety

This app **does not upload anything** and does **not write extracted data to disk**.

However - during runtime it can display and/or copy *fragments of CSV content*:
- **Preview** shows sample rows
- **Columns** may show **sample values**
- **Copy format description** includes `sample:` values for columns
- Compare mode shows cell values from selected rows

So: publishing the **code** is safe, but be careful with:
- screenshots (they can show real values)
- pasting “format description” into tickets/Slack if it contains sensitive `sample:` values
- committing real CSV files into the repo

---

## Requirements

- Python 3.x (stdlib-only)
- Windows recommended (it’s a `.pyw` GUI), but it can run anywhere Tkinter is available

No third-party dependencies.

---

## Run

### Option A: run directly
- Double-click `kkr-csv-format-discover.pyw` on Windows  
  or
- From terminal: `python kkr-csv-format-discover.pyw`

---

## How to use

### 1) Analyze one file
1. Click **Open CSV...**
2. Review **CSV profile fields** + **Diagnostics**
3. Use **Columns** tab for structure + sample values
4. Use **Preview** tab for a quick visual sanity check
5. Click **Copy format description** and paste it where you keep config notes

### 2) Compare two files (format)
1. Open a first CSV (File A)
2. Click **Compare...** and select File B
3. Review:
   - differences in delimiter/decimal/line terminator/quoting/date format
   - aligned sample rows (A vs B) with one horizontal scroll
4. (Optional) change **Row A** / **Row B** to compare later records
5. Use **Copy A date format** / **Copy B date format** if you only need date pattern

---

## Notes and limitations (important)

- Some checks are **heuristics**. Example: “quote char not observed” means “not seen in the sample” - **not a proof it never occurs**.
- Sampling is capped (e.g. first N rows) to keep the app responsive.
- CSV “sniffing” can be wrong on messy files. Always confirm with Preview if it matters.
- Decimal/date inference is best-effort; weird mixed columns may confuse it.

---

## Suggested workflow with kkr-query2xlsx

If you need to generate stable CSV exports from SQL:
1. Analyze an incoming CSV with this tool and capture the format fields
2. Configure your exporter (example: `kkr-query2xlsx`) with the same delimiter/decimal/line endings/date format
3. Use Compare mode when a provider “changes something” and you need to confirm what exactly drifted

Link:
- https://github.com/kkrysztofczyk/kkr-query2xlsx

---

## License

MIT - see `LICENSE`.
