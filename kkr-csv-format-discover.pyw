"""
kkr-csv_format_discover — CSV format analysis tool.

Part of the KKR helper tools family.
Pure stdlib (tkinter + csv), no pip installs required.

Purpose
-------
Detects CSV structural details and presents them in a layout compatible with
the CSV profile form in kkr-query2xlsx, so you can transfer the
format settings between the two programs.

Notes
-----
- Some fields shown here are *diagnostics* (e.g., BOM / Non-ASCII). They are
  useful when troubleshooting file interoperability and may later become
  first-class fields in kkr-query2xlsx, but today they are not stored in the
  kkr-query2xlsx CSV profile.
"""

from __future__ import annotations

import csv
import io
import os
import re
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_VERSION = "0.1"
APP_TITLE = "kkr — CSV Format Discover"
KKR_QUERY2XLSX_URL = "https://github.com/kkrysztofczyk/kkr-query2xlsx"

# ── Analysis ─────────────────────────────────────────────────────────────────

DATE_PATTERNS = [
    re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}"),
    re.compile(r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}"),
]


def _guess_col_type(values: list[str]) -> str:
    if not values:
        return "string"
    if all(v.lower() in ("true", "false", "0", "1", "yes", "no") for v in values):
        return "boolean"
    if all(any(p.match(v) for p in DATE_PATTERNS) for v in values):
        return "date"
    try:
        for v in values:
            float(v.replace(",", "."))
        return "number"
    except ValueError:
        pass
    return "string"


def _guess_decimal_separator(rows: list[dict], fields: list[str]) -> str:
    dot_count = 0
    comma_count = 0
    pat_dot = re.compile(r"^-?\d+\.\d+$")
    pat_comma = re.compile(r"^-?\d+,\d+$")
    for row in rows:
        for f in fields:
            v = (row.get(f) or "").strip()
            if pat_dot.match(v):
                dot_count += 1
            if pat_comma.match(v):
                comma_count += 1

    # Practical heuristics:
    # - if we see only comma-decimals, choose comma immediately (even if just 1 hit)
    # - otherwise choose the dominant one
    if comma_count > 0 and dot_count == 0:
        return ","
    if comma_count > dot_count and comma_count >= 1:
        return ","
    return "."


def _guess_quoting_strategy(raw: str, delimiter: str, quotechar: str) -> str:
    if not quotechar or quotechar not in raw:
        return "none"

    lines = [l for l in raw.splitlines() if l.strip()]
    if len(lines) >= 2:
        parts = lines[1].split(delimiter)
        if parts and all(
            p.strip().startswith(quotechar) and p.strip().endswith(quotechar)
            for p in parts
            if p.strip()
        ):
            return "all"
    return "minimal"


def _detect_doublequote(raw: str, quotechar: str) -> bool:
    return bool(quotechar) and (quotechar + quotechar) in raw


def _detect_escapechar(raw: str, quotechar: str) -> str:
    # CSV escapechar is only relevant when QUOTE_NONE or when the input actually uses it.
    # We do a conservative detection: \" or \' patterns.
    if not quotechar:
        return ""
    if quotechar == '"' and '\\"' in raw:
        return "\\"
    if quotechar == "'" and "\\'" in raw:
        return "\\"
    return ""


def _guess_date_format(rows: list[dict], fields: list[str]) -> str:
    """Try to guess a strftime-compatible date format from the data."""
    candidates = [
        (re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$"), "%Y-%m-%d %H:%M:%S"),
        (re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$"), "%Y-%m-%d %H:%M"),
        (re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}"), "%Y-%m-%dT%H:%M:%S"),
        (re.compile(r"^\d{4}-\d{2}-\d{2}$"), "%Y-%m-%d"),
        (re.compile(r"^\d{4}/\d{2}/\d{2}$"), "%Y/%m/%d"),
        (re.compile(r"^\d{2}\.\d{2}\.\d{4}$"), "%d.%m.%Y"),
        (re.compile(r"^\d{2}/\d{2}/\d{4}$"), "%d/%m/%Y"),
        (re.compile(r"^\d{2}-\d{2}-\d{4}$"), "%d-%m-%Y"),
    ]
    for row in rows[:50]:
        for f in fields:
            v = (row.get(f) or "").strip()
            if not v:
                continue
            for pat, fmt in candidates:
                if pat.match(v):
                    return fmt
    return ""


def _escape_visible(s: str) -> str:
    return (
        s.replace("\r\n", "\\r\\n")
        .replace("\n", "\\n")
        .replace("\r", "\\r")
        .replace("\t", "\\t")
    )


def _format_bytes(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1_048_576:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1_048_576:.1f} MB"


def _safe_sniff_dialect(sample_text: str):
    """
    Try sniffing delimiter & basic dialect properties, but avoid the common
    trap of "space" being selected as a delimiter.
    """
    try:
        # Delimiters we actually use in CSV exports / common upstream feeds.
        return csv.Sniffer().sniff(sample_text, delimiters=",;\t|")
    except csv.Error:
        return None


def _quoting_to_csv_const(quoting_name: str) -> int:
    mapping = {
        "minimal": csv.QUOTE_MINIMAL,
        "all": csv.QUOTE_ALL,
        "nonnumeric": csv.QUOTE_NONNUMERIC,
        "none": csv.QUOTE_NONE,
    }
    return mapping.get((quoting_name or "minimal").lower(), csv.QUOTE_MINIMAL)


SAMPLE_LIMIT = 200  # max records parsed for inference / preview


def analyze(path: str) -> dict:
    """Analyse a CSV/TSV file and return structural metadata."""
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        raw_bytes = f.read()

    # ── Encoding & BOM ───────────────────────────────────────────────────
    has_bom = raw_bytes[:3] == b"\xef\xbb\xbf"
    try:
        raw = raw_bytes.decode("utf-8-sig")
        encoding = "utf-8-sig" if has_bom else "utf-8"
    except UnicodeDecodeError:
        raw = raw_bytes.decode("latin-1")
        encoding = "latin-1"

    has_non_ascii = any(ord(ch) > 127 for ch in raw)

    # ── Line ending ──────────────────────────────────────────────────────
    has_crlf = "\r\n" in raw
    has_cr = (not has_crlf) and ("\r" in raw)
    if has_crlf:
        lineterminator = "\\r\\n"
    elif has_cr:
        lineterminator = "\\r"
    else:
        lineterminator = "\\n"

    # Basic empty-file guard (works even with multiline fields)
    if not raw.strip():
        raise ValueError("File is empty")

    # ── Delimiter & dialect (sniff) ──────────────────────────────────────
    lines = [l for l in raw.splitlines() if l.strip()]
    sample_text = "\n".join(lines[:50]) if lines else raw[:10_000]
    sniffed = _safe_sniff_dialect(sample_text)

    delimiter = sniffed.delimiter if sniffed else ","
    sniff_quotechar = sniffed.quotechar if sniffed else '"'
    skipinitialspace = bool(getattr(sniffed, "skipinitialspace", False)) if sniffed else False

    # ── Quoting / escaping (detect + normalize for kkr-query2xlsx compatibility) ──
    quoting = _guess_quoting_strategy(raw, delimiter, sniff_quotechar)
    quoting_const = _quoting_to_csv_const(quoting)

    # "Detected" quotechar: present in the input text
    quotechar_detected = sniff_quotechar if (sniff_quotechar and sniff_quotechar in raw) else ""

    # "Profile" quotechar: must always be a 1-char value (csv module requires it)
    # even when quoting=none.
    quotechar_profile = sniff_quotechar or '"'
    if len(quotechar_profile) != 1:
        quotechar_profile = '"'

    doublequote = _detect_doublequote(raw, quotechar_profile)
    escapechar = _detect_escapechar(raw, quotechar_profile)

    # ── Parse rows (use dialect-like params, not only delimiter) ──────────
    reader = csv.DictReader(
        io.StringIO(raw),
        delimiter=delimiter,
        quotechar=quotechar_profile,
        quoting=quoting_const,
        escapechar=escapechar or None,
        doublequote=doublequote,
        skipinitialspace=skipinitialspace,
    )
    fields = reader.fieldnames or []

    rows: list[dict] = []
    for i, row in enumerate(reader):
        if i >= SAMPLE_LIMIT:  # sample window for inference (decimal/date/types)
            break
        rows.append(row)

    rows_sampled = len(rows)

    # Diagnostic: did we observe the delimiter inside any parsed field value?
    # If yes, exporting with quoting=none may require either quoting or delimiter replacement.
    delimiter_in_field_sample = any(
        (delimiter in str((row.get(f) or "")))
        for row in rows
        for f in fields
    )


    # Total row count (records), not "non-empty lines".
    # Use csv.reader to properly handle multiline fields.
    total_rows = 0
    count_reader = csv.reader(
        io.StringIO(raw),
        delimiter=delimiter,
        quotechar=quotechar_profile,
        quoting=quoting_const,
        escapechar=escapechar or None,
        doublequote=doublequote,
        skipinitialspace=skipinitialspace,
    )
    header = next(count_reader, None)
    if header is not None:
        for _ in count_reader:
            total_rows += 1

    # ── Decimal separator ────────────────────────────────────────────────
    decimal = _guess_decimal_separator(rows, fields)

    # ── Date format ──────────────────────────────────────────────────────
    date_format = _guess_date_format(rows, fields)

    # ── Column info ──────────────────────────────────────────────────────
    columns = []
    for col in fields:
        values = [str(row.get(col, "") or "").strip() for row in rows]
        non_null = [v for v in values if v != ""]
        null_count = len(values) - len(non_null)
        unique = len(set(non_null))
        col_type = _guess_col_type(non_null)
        sample = [v[:40] for v in non_null[:3]]
        columns.append(
            {
                "name": col,
                "type": col_type,
                "null_count": null_count,
                "unique": unique,
                "sample": sample,
            }
        )

    return {
        "file_name": os.path.basename(path),
        "file_path": path,
        "file_size": size,
        "total_rows": total_rows,
        "total_columns": len(fields),
        "rows_sampled": rows_sampled,
        "sample_limit": SAMPLE_LIMIT,
        "delimiter_in_field_sample": delimiter_in_field_sample,
        "encoding": encoding,
        "has_bom": has_bom,
        "has_non_ascii": has_non_ascii,
        "delimiter": delimiter,
        "delimiter_replacement": "",  # not detectable; keep empty by default
        "decimal": decimal,
        "lineterminator": lineterminator,
        "quotechar_profile": quotechar_profile,
        "quotechar_detected": quotechar_detected,
        "quoting": quoting,
        "escapechar": escapechar,
        "doublequote": doublequote,
        "date_format": date_format,
        "skipinitialspace": skipinitialspace,
        "columns": columns,
        "fields": fields,
        "preview": rows[:5],
    }


def build_description(info: dict) -> str:
    """Build a plain-text description suitable for pasting."""
    lines = [
        f'=== CSV Format: {info["file_name"]} ===',
        f'Rows: {info["total_rows"]} | Columns: {info["total_columns"]} | Sampled: {info.get("rows_sampled",0)}/{info.get("sample_limit","?")}',
        "",
        "--- CSV profile fields (kkr-query2xlsx) ---",
        f'Encoding:                    {info["encoding"]}',
        f'Field delimiter:             {info["delimiter"]}',
        f'Delimiter replacement:        {info.get("delimiter_replacement","")}',
        f'Decimal separator:           {info["decimal"]}',
        f'Line terminator:             {info["lineterminator"]}',
        f'Quote character:             {info["quotechar_profile"]}',
        f'Quoting:                     {info["quoting"]}',
        f'Escape character:            {info["escapechar"] or "(none)"}',
        f'Doublequote:                 {"yes" if info["doublequote"] else "no"}',
        f'Date format:                 {info["date_format"] or "(not detected)"}',
        "",
        "--- Extra diagnostics (not in CSV profile yet) ---",
        f'BOM present:                 {"yes" if info["has_bom"] else "no"}',
        f'Non-ASCII chars present:     {"yes" if info["has_non_ascii"] else "no"}',
        f'Delimiter inside fields (sample): {"yes" if info.get("delimiter_in_field_sample") else "no"}',
        f'Quote char observed in file:  {info["quotechar_detected"] if info["quotechar_detected"] else "(not observed)"}',
        "",
        "--- Columns ---",
    ]
    for i, c in enumerate(info["columns"], 1):
        s = ", ".join(c["sample"])
        lines.append(
            f'{i}. "{c["name"]}" — {c["type"]} | '
            f'{c["unique"]} unique | {c["null_count"]} nulls | sample: {s}'
        )
    return "\n".join(lines)



# ── Format comparison (two files) ────────────────────────────────────────────

def _fmt_bool(v: bool) -> str:
    return "yes" if v else "no"


def _fmt_str(v: str) -> str:
    return v if (v is not None and v != "") else "(empty)"


def _format_field_value(field_key: str, info: dict) -> str:
    """Return a user-facing value for comparison tables."""
    if field_key == "encoding":
        return info.get("encoding", "")
    if field_key == "has_bom":
        return _fmt_bool(bool(info.get("has_bom")))
    if field_key == "has_non_ascii":
        return _fmt_bool(bool(info.get("has_non_ascii")))
    if field_key == "delimiter_in_field_sample":
        return _fmt_bool(bool(info.get("delimiter_in_field_sample")))
    if field_key == "delimiter":
        return info.get("delimiter", "")
    if field_key == "delimiter_replacement":
        return _fmt_str(info.get("delimiter_replacement", ""))
    if field_key == "decimal":
        return info.get("decimal", "")
    if field_key == "lineterminator":
        return info.get("lineterminator", "")
    if field_key == "quotechar_profile":
        return info.get("quotechar_profile", "")
    if field_key == "quotechar_detected":
        return info.get("quotechar_detected") or "(not observed)"
    if field_key == "quoting":
        return info.get("quoting", "")
    if field_key == "escapechar":
        return _fmt_str(info.get("escapechar", ""))
    if field_key == "doublequote":
        return _fmt_bool(bool(info.get("doublequote")))
    if field_key == "date_format":
        return info.get("date_format") or "(not detected)"
    return _fmt_str(str(info.get(field_key, "")))


COMPARE_FIELDS: list[tuple[str, str, str]] = [
    # (group, label, key)
    ("CSV profile fields", "Encoding", "encoding"),
    ("CSV profile fields", "Field delimiter", "delimiter"),
    ("CSV profile fields", "Delimiter replacement", "delimiter_replacement"),
    ("CSV profile fields", "Decimal separator", "decimal"),
    ("CSV profile fields", "Line terminator", "lineterminator"),
    ("CSV profile fields", "Quote character", "quotechar_profile"),
    ("CSV profile fields", "Quoting", "quoting"),
    ("CSV profile fields", "Escape character", "escapechar"),
    ("CSV profile fields", "Doublequote", "doublequote"),
    ("CSV profile fields", "Date format", "date_format"),
    ("Diagnostics", "BOM present", "has_bom"),
    ("Diagnostics", "Non-ASCII chars present", "has_non_ascii"),
    ("Diagnostics", "Delimiter inside fields (sample)", "delimiter_in_field_sample"),
    ("Diagnostics", "Quote char observed in file", "quotechar_detected"),
]


def compare_formats(a: dict, b: dict) -> dict:
    """
    Compare two analyzed CSV files (format only, not data).

    Returns a dict with:
    - rows: list of comparison rows (group, label, a, b, same)
    - columns: summary and differences about column headers
    """
    rows = []
    diff_count = 0
    for group, label, key in COMPARE_FIELDS:
        va = _format_field_value(key, a)
        vb = _format_field_value(key, b)
        same = (va == vb)
        if not same:
            diff_count += 1
        rows.append(
            {
                "group": group,
                "label": label,
                "a": va,
                "b": vb,
                "same": same,
                "key": key,
            }
        )

    cols_a = list(a.get("fields") or [])
    cols_b = list(b.get("fields") or [])
    set_a = set(cols_a)
    set_b = set(cols_b)

    only_a = [c for c in cols_a if c not in set_b]
    only_b = [c for c in cols_b if c not in set_a]

    same_set = (set_a == set_b)
    same_order = (cols_a == cols_b) if same_set else False

    order_mismatches = []
    if same_set and not same_order:
        pos_b = {c: i for i, c in enumerate(cols_b)}
        for i, c in enumerate(cols_a):
            if pos_b.get(c) != i:
                order_mismatches.append((c, i, pos_b.get(c)))

    columns = {
        "a_count": len(cols_a),
        "b_count": len(cols_b),
        "same_set": same_set,
        "same_order": same_order,
        "only_a": only_a,
        "only_b": only_b,
        "order_mismatches": order_mismatches[:50],  # keep it readable
    }

    return {"rows": rows, "diff_count": diff_count, "columns": columns}


def build_compare_description(a: dict, b: dict, cmp: dict) -> str:
    """Plain-text comparison report (clipboard-friendly)."""
    lines = [
        "=== CSV Format Compare (format only) ===",
        f'File A: {a.get("file_name","")}  |  {a.get("file_path","")}',
        f'File B: {b.get("file_name","")}  |  {b.get("file_path","")}',
        f'A: rows={a.get("total_rows","?")} cols={a.get("total_columns","?")} sampled={a.get("rows_sampled",0)}/{a.get("sample_limit","?")}',
        f'B: rows={b.get("total_rows","?")} cols={b.get("total_columns","?")} sampled={b.get("rows_sampled",0)}/{b.get("sample_limit","?")}',
        "",
        f"Differences in fields: {cmp.get('diff_count', 0)}",
        "",
    ]

    current_group = None
    for r in cmp.get("rows", []):
        if r["group"] != current_group:
            current_group = r["group"]
            lines.append(f"--- {current_group} ---")
        if r["same"]:
            lines.append(f'{r["label"]}: SAME ({r["a"]})')
        else:
            lines.append(f'{r["label"]}: A={r["a"]}  |  B={r["b"]}')

    c = cmp.get("columns", {})
    lines += [
        "",
        "--- Columns (headers) ---",
        f'Count: A={c.get("a_count",0)}  |  B={c.get("b_count",0)}',
    ]
    if c.get("same_set") and c.get("same_order"):
        lines.append("Headers: SAME (names + order)")
    elif c.get("same_set") and not c.get("same_order"):
        lines.append("Headers: SAME set, DIFFERENT order")
        mism = c.get("order_mismatches") or []
        if mism:
            lines.append("Order mismatches (col, posA, posB) [first 50]:")
            for col, pa, pb in mism:
                lines.append(f" - {col}: {pa} -> {pb}")
    else:
        lines.append("Headers: DIFFERENT")
        if c.get("only_a"):
            lines.append("Only in A:")
            for name in c["only_a"][:200]:
                lines.append(f" - {name}")
        if c.get("only_b"):
            lines.append("Only in B:")
            for name in c["only_b"][:200]:
                lines.append(f" - {name}")

    return "\n".join(lines)




# ── Row extraction (for compare preview) ───────────────────────────────────

def _open_csv_text_stream(path: str, encoding: str):
    """Open a CSV file as text with a safe encoding choice."""
    enc = (encoding or "utf-8").lower().strip()
    if enc in ("utf8", "utf-8"):
        enc = "utf-8"
    elif enc in ("utf8-sig", "utf-8-sig"):
        enc = "utf-8-sig"
    # newline="" is recommended for csv module.
    return open(path, "r", encoding=enc, newline="", errors="replace")


def read_data_row_by_index(info: dict, row_index_1based: int) -> tuple[list[str], list[str] | None]:
    """
    Read a single data row by 1-based index (excluding header).

    Returns: (header, row_or_none)
    """
    path = info.get("file_path") or ""
    if not path:
        return ([], None)

    try:
        idx = int(row_index_1based)
    except Exception:
        idx = 1
    if idx < 1:
        idx = 1

    delimiter = info.get("delimiter", ",")
    quotechar = info.get("quotechar_profile", '"') or '"'
    quoting_const = _quoting_to_csv_const(info.get("quoting", "minimal"))
    escapechar = info.get("escapechar") or None
    doublequote = bool(info.get("doublequote"))
    skipinitialspace = bool(info.get("skipinitialspace", False))

    with _open_csv_text_stream(path, info.get("encoding", "utf-8")) as f:
        reader = csv.reader(
            f,
            delimiter=delimiter,
            quotechar=quotechar,
            quoting=quoting_const,
            escapechar=escapechar,
            doublequote=doublequote,
            skipinitialspace=skipinitialspace,
        )
        header = next(reader, None)
        if header is None:
            return ([], None)

        for i, row in enumerate(reader, start=1):
            if i == idx:
                return (list(header), list(row))

    return (list(header), None)


def format_row_block(header: list[str], row: list[str] | None, *, max_cell: int = 80) -> str:
    """Format header + one row as an aligned, horizontally-scrollable block."""
    if not header:
        return "(no header)"

    if row is None:
        # still show headers for context
        row = ["" for _ in header]
        note = "(row not found)"
    else:
        note = ""

    n = max(len(header), len(row))

    def _h(i: int) -> str:
        return header[i] if i < len(header) else f"COL_{i+1}"

    def _v(i: int) -> str:
        v = row[i] if i < len(row) else ""
        v = str(v)
        if len(v) > max_cell:
            v = v[: max_cell - 1] + "…"
        return v

    widths = []
    for i in range(n):
        widths.append(max(len(_h(i)), len(_v(i)), 3))

    hdr = "  ".join(_h(i).ljust(widths[i]) for i in range(n))
    sep = "  ".join("─" * widths[i] for i in range(n))
    val = "  ".join(_v(i).ljust(widths[i]) for i in range(n))

    if note:
        return note + "\n" + hdr + "\n" + sep + "\n" + val
    return hdr + "\n" + sep + "\n" + val


def _field_keys(fields: list[str]) -> list[tuple[str, int]]:
    """
    Convert a header list into stable column keys: (name, occurrence_index).

    This keeps duplicates distinguishable (e.g. "Amount" and another "Amount").
    """
    counts: dict[str, int] = {}
    out: list[tuple[str, int]] = []
    for name in fields:
        name = "" if name is None else str(name)
        counts[name] = counts.get(name, 0) + 1
        out.append((name, counts[name]))
    return out


def _merge_field_keys(
    keys_a: list[tuple[str, int]], keys_b: list[tuple[str, int]]
) -> list[tuple[str, int]]:
    """
    Merge column keys so that same-named columns align, and missing columns
    become empty cells on the other side.

    Rule: keep A order, then append keys that exist only in B.
    """
    merged = list(keys_a)
    seen = set(keys_a)
    for k in keys_b:
        if k not in seen:
            merged.append(k)
            seen.add(k)
    return merged


def _display_field_key(k: tuple[str, int]) -> str:
    name, occ = k
    base = name if name else "(empty)"
    return base if occ == 1 else f"{base} ({occ})"


def _row_map_from_header_and_row(
    header: list[str], row: list[str] | None
) -> dict[tuple[str, int], str]:
    keys = _field_keys(header)
    out: dict[tuple[str, int], str] = {}
    for i, k in enumerate(keys):
        out[k] = (row[i] if (row is not None and i < len(row)) else "") or ""
    return out


def _truncate_cell(v: str, *, max_len: int = 80) -> str:
    s = "" if v is None else str(v)
    if len(s) > max_len:
        return s[: max_len - 1] + "…"
    return s


# ── Native ttk theme (same logic as main KKR app) ───────────────────────────


def apply_native_ttk_theme(root: tk.Tk) -> None:
    """Apply a safe, native-looking ttk theme when available."""
    try:
        style = ttk.Style(root)
        available = set(style.theme_names())
        preferred: list[str] = []
        if sys.platform.startswith("win"):
            preferred = ["vista", "xpnative"]
        elif sys.platform == "darwin":
            preferred = ["aqua"]
        preferred.append("clam")
        for theme in preferred:
            if theme in available:
                try:
                    style.theme_use(theme)
                    break
                except tk.TclError:
                    continue
    except Exception:
        pass


# ── GUI ──────────────────────────────────────────────────────────────────────

MIN_WINDOW_WIDTH = 780
MIN_WINDOW_HEIGHT = 560


class CSVFormatDiscoverApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.root.minsize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        apply_native_ttk_theme(self.root)

        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 900, 760
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        self.info: dict | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = self.root

        # ── File selection ───────────────────────────────────────────────
        file_frame = ttk.LabelFrame(root, text="File", padding=(10, 10))
        file_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        file_frame.columnconfigure(1, weight=1)

        self.file_path_var = tk.StringVar(value="(no file selected)")
        ttk.Label(file_frame, text="Path:").grid(row=0, column=0, sticky="w")

        path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly")
        path_entry.grid(row=0, column=1, sticky="we", padx=(6, 0))

        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=0, column=2, sticky="e", padx=(10, 0))

        ttk.Button(btn_frame, text="Open CSV…", command=self._on_open).pack(side="left")

        ttk.Button(btn_frame, text="Compare…", command=self._on_compare).pack(
            side="left", padx=(8, 8)
        )

        self.btn_copy = ttk.Button(
            btn_frame,
            text="Copy format description",
            command=self._on_copy,
            state=tk.DISABLED,
        )
        self.btn_copy.pack(side="left")

        self.status_var = tk.StringVar(value="")
        ttk.Label(file_frame, textvariable=self.status_var, foreground="gray").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )

        # ── Format details ───────────────────────────────────────────────
        details_frame = ttk.LabelFrame(root, text="Format details", padding=(10, 10))
        details_frame.pack(fill=tk.X, padx=10, pady=5)
        details_frame.columnconfigure(0, weight=1)
        details_frame.columnconfigure(1, weight=1)

        profile_lf = ttk.LabelFrame(details_frame, text="CSV profile fields", padding=(10, 8))
        diag_lf = ttk.LabelFrame(details_frame, text="Diagnostics", padding=(10, 8))
        profile_lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        diag_lf.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # IMPORTANT: don't reserve a permanent "hint column".
        # Keeping a fixed 3rd column causes the value entries to get squeezed
        # (even for rows without hints). Hints are rendered *below* the field.
        for lf in (profile_lf, diag_lf):
            lf.columnconfigure(0, weight=0)
            lf.columnconfigure(1, weight=1, minsize=280)

        self.profile_vars: dict[str, tk.StringVar] = {}

        def _add_fields(parent, fields):  # noqa: ANN001
            row_idx = 0
            for (label, key, width, hint) in fields:
                var = tk.StringVar(value="")
                self.profile_vars[key] = var

                ttk.Label(parent, text=label).grid(
                    row=row_idx, column=0, sticky="w", pady=(2, 0)
                )
                entry = ttk.Entry(parent, textvariable=var, width=width, state="readonly")
                entry.grid(row=row_idx, column=1, sticky="we", padx=(6, 0), pady=(2, 0))
                row_idx += 1

                if hint:
                    ttk.Label(
                        parent,
                        text=hint,
                        foreground="gray",
                        wraplength=540,
                        justify="left",
                    ).grid(
                        row=row_idx, column=1, sticky="w", padx=(6, 0), pady=(0, 4)
                    )
                    row_idx += 1


        profile_fields = [
            ("Encoding:", "encoding", 18, None),
            ("Field delimiter:", "delimiter", 8, None),
            ("Delimiter replacement:", "delimiter_replacement", 10, "(manual; only if you want sanitizing)"),
            ("Decimal separator:", "decimal", 8, None),
            ("Line terminator:", "lineterminator", 12, r"(use \n / \r\n / \r / \t)"),
            ("Quote character:", "quotechar_profile", 8, None),
            ("Quoting:", "quoting", 14, None),
            ("Escape character:", "escapechar", 8, "(escape char; empty = quoting)"),
            ("Doublequote:", "doublequote", 8, None),
            ("Date format:", "date_format", 28, None),
        ]

        diag_fields = [
            ("BOM present:", "bom", 8, None),
            ("Non-ASCII chars present:", "non_ascii", 8, None),
            ("Delimiter inside fields (sample):", "delimiter_in_field_sample", 10, "(absence is NOT a proof)"),
            ("Quote char observed in file:", "quotechar_detected", 26, "(not observed ≠ never)"),
        ]

        _add_fields(profile_lf, profile_fields)
        _add_fields(diag_lf, diag_fields)

        self.stats_var = tk.StringVar(value="")
        ttk.Label(details_frame, textvariable=self.stats_var, foreground="gray").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        help_frame = ttk.Frame(details_frame)
        help_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=(8, 0))
        ttk.Label(help_frame, text="Need to generate CSV exports?").pack(side="left")
        link = tk.Label(help_frame, text="kkr-query2xlsx (GitHub)", fg="#0066cc", cursor="hand2")
        link.pack(side="left", padx=(6, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open_new_tab(KKR_QUERY2XLSX_URL))


        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        tab_cols = ttk.Frame(nb)
        tab_prev = ttk.Frame(nb)
        nb.add(tab_cols, text="Columns")
        nb.add(tab_prev, text="Preview")

        columns_frame = ttk.LabelFrame(tab_cols, text="Columns", padding=(10, 10))
        columns_frame.pack(fill=tk.BOTH, expand=True)
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.rowconfigure(0, weight=1)

        tree_cols = ("name", "type", "unique", "nulls", "sample")
        self.tree = ttk.Treeview(
            columns_frame, columns=tree_cols, show="headings", selectmode="browse"
        )
        self.tree.heading("name", text="Column name")
        self.tree.heading("type", text="Type")
        self.tree.heading("unique", text="Unique")
        self.tree.heading("nulls", text="Nulls")
        self.tree.heading("sample", text="Sample values")

        self.tree.column("name", width=220, minwidth=120)
        self.tree.column("type", width=80, minwidth=60, anchor="center")
        self.tree.column("unique", width=70, minwidth=50, anchor="center")
        self.tree.column("nulls", width=70, minwidth=50, anchor="center")
        self.tree.column("sample", width=520, minwidth=120)

        tree_yscroll = ttk.Scrollbar(columns_frame, orient="vertical", command=self.tree.yview)
        tree_xscroll = ttk.Scrollbar(columns_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_yscroll.set, xscrollcommand=tree_xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_yscroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        tree_xscroll.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        # ── Preview ──────────────────────────────────────────────────────
        preview_frame = ttk.LabelFrame(tab_prev, text="Preview (first 5 rows)", padding=(10, 10))
        preview_frame.pack(fill=tk.BOTH, expand=True)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.preview_text = tk.Text(
            preview_frame,
            height=7,
            wrap="none",
            state="disabled",
            font=("Consolas", 9) if sys.platform == "win32" else ("TkFixedFont",),
        )
        preview_xscroll = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_text.xview)
        preview_yscroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        self.preview_text.configure(
            xscrollcommand=preview_xscroll.set,
            yscrollcommand=preview_yscroll.set,
        )
        preview_yscroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        preview_xscroll.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.preview_text.grid(row=0, column=0, sticky="nsew")

    # ── Actions ──────────────────────────────────────────────────────────

    def _on_open(self) -> None:
        path = filedialog.askopenfilename(
            title="Select CSV / TSV file",
            filetypes=[
                ("CSV / TSV / TXT", "*.csv *.tsv *.txt"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.status_var.set("")
        try:
            self.info = analyze(path)
        except Exception as exc:
            self.status_var.set(f"Error: {exc}")
            return

        self.file_path_var.set(path)
        self._refresh_results()
        self.btn_copy.configure(state=tk.NORMAL)
        self.status_var.set(
            f"Analysed {self.info['total_rows']:,} rows (sampled {self.info.get('rows_sampled',0)}/{self.info.get('sample_limit','?')}), {self.info['total_columns']} columns."
        )

    def _on_copy(self) -> None:
        if not self.info:
            return
        desc = build_description(self.info)
        self.root.clipboard_clear()
        self.root.clipboard_append(desc)
        prev_text = self.btn_copy.cget("text")
        self.btn_copy.configure(text="✓  Copied!")
        self.root.after(2000, lambda: self.btn_copy.configure(text=prev_text))


    def _on_compare(self) -> None:
        """
        Compare formats of two files (format only, not data).

        If a file is already loaded, it becomes File A, and user selects File B.
        Otherwise user selects File A and File B.
        """
        # File A
        path_a = self.info.get("file_path") if self.info else None
        info_a = self.info

        if not path_a:
            path_a = filedialog.askopenfilename(
                title="Select first CSV / TSV file (File A)",
                filetypes=[
                    ("CSV / TSV / TXT", "*.csv *.tsv *.txt"),
                    ("All files", "*.*"),
                ],
            )
            if not path_a:
                return
            try:
                info_a = analyze(path_a)
            except Exception as exc:
                messagebox.showerror("Compare", f"Error analysing File A:\n{exc}")
                return

        # File B
        path_b = filedialog.askopenfilename(
            title="Select second CSV / TSV file (File B)",
            filetypes=[
                ("CSV / TSV / TXT", "*.csv *.tsv *.txt"),
                ("All files", "*.*"),
            ],
        )
        if not path_b:
            return

        try:
            info_b = analyze(path_b)
        except Exception as exc:
            messagebox.showerror("Compare", f"Error analysing File B:\n{exc}")
            return

        cmp = compare_formats(info_a, info_b)
        self._show_compare_window(info_a, info_b, cmp)

    def _show_compare_window(self, info_a: dict, info_b: dict, cmp: dict) -> None:
        win = tk.Toplevel(self.root)
        win.title("Compare CSV formats (format only)")
        win.minsize(820, 520)
        win.geometry("980x720")

        top = ttk.Frame(win, padding=(10, 10))
        top.pack(fill=tk.X)

        lbl_a = ttk.Label(top, text=f'A: {info_a.get("file_name","")}  |  {info_a.get("file_path","")}')
        lbl_b = ttk.Label(top, text=f'B: {info_b.get("file_name","")}  |  {info_b.get("file_path","")}')
        lbl_a.pack(anchor="w")
        lbl_b.pack(anchor="w", pady=(2, 0))

        summary = ttk.Label(
            top,
            text=f'Differences in fields: {cmp.get("diff_count",0)}  |  '
                 f'Columns: A={info_a.get("total_columns","?")} B={info_b.get("total_columns","?")}',
            foreground="gray",
        )
        summary.pack(anchor="w", pady=(6, 0))

        btns = ttk.Frame(top)
        btns.pack(anchor="e", fill=tk.X, pady=(8, 0))

        def _copy_compare() -> None:
            desc = build_compare_description(info_a, info_b, cmp)
            self.root.clipboard_clear()
            self.root.clipboard_append(desc)

        def _copy_text(s: str) -> None:
            self.root.clipboard_clear()
            self.root.clipboard_append(s)

        def _copy_date_a() -> None:
            _copy_text(info_a.get("date_format", "") or "")

        def _copy_date_b() -> None:
            _copy_text(info_b.get("date_format", "") or "")

        ttk.Button(btns, text="Copy A date format", command=_copy_date_a).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Copy B date format", command=_copy_date_b).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Copy comparison", command=_copy_compare).pack(side="right")

        # Field-level compare table
        fields_frame = ttk.LabelFrame(win, text="Fields", padding=(10, 10))
        fields_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))
        fields_frame.columnconfigure(0, weight=1)
        fields_frame.rowconfigure(0, weight=1)

        tree_cols = ("field", "a", "b")
        tree = ttk.Treeview(fields_frame, columns=tree_cols, show="headings", selectmode="browse")
        tree.heading("field", text="Field")
        tree.heading("a", text="File A")
        tree.heading("b", text="File B")

        tree.column("field", width=240, minwidth=160)
        tree.column("a", width=280, minwidth=180)
        tree.column("b", width=280, minwidth=180)

        yscroll = ttk.Scrollbar(fields_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")

        try:
            tree.tag_configure("group", foreground="gray")
            tree.tag_configure("diff", foreground="#b00020")  # safe-ish red
        except Exception:
            pass

        current_group = None
        for r in cmp.get("rows", []):
            if r["group"] != current_group:
                current_group = r["group"]
                tree.insert("", "end", values=(f"— {current_group} —", "", ""), tags=("group",))
            if r["same"]:
                tree.insert("", "end", values=(r["label"], r["a"], r["b"]))
            else:
                tree.insert("", "end", values=(f"≠ {r['label']}", r["a"], r["b"]), tags=("diff",))



        # Sample rows (aligned by column names; shared horizontal scrolling)
        sample_frame = ttk.LabelFrame(win, text="Sample rows (aligned columns)", padding=(10, 10))
        sample_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 5))
        sample_frame.columnconfigure(0, weight=1)
        sample_frame.rowconfigure(1, weight=1)

        ctrl = ttk.Frame(sample_frame)
        ctrl.grid(row=0, column=0, sticky="we")

        max_a = max(1, int(info_a.get("total_rows", 1) or 1))
        max_b = max(1, int(info_b.get("total_rows", 1) or 1))

        row_a_var = tk.IntVar(value=1)
        row_b_var = tk.IntVar(value=1)

        ttk.Label(ctrl, text="Row A:").pack(side="left")
        sp_a = tk.Spinbox(ctrl, from_=1, to=max_a, width=8, textvariable=row_a_var)
        sp_a.pack(side="left", padx=(5, 12))

        ttk.Label(ctrl, text="Row B:").pack(side="left")
        sp_b = tk.Spinbox(ctrl, from_=1, to=max_b, width=8, textvariable=row_b_var)
        sp_b.pack(side="left", padx=(5, 12))

        hint = ttk.Label(ctrl, text="(1-based data rows; header excluded)", foreground="gray")
        hint.pack(side="left", padx=(6, 0))

        # Build merged header order (name + occurrence index), so duplicates are handled.
        keys_a = _field_keys(list(info_a.get("fields") or []))
        keys_b = _field_keys(list(info_b.get("fields") or []))
        merged_keys = _merge_field_keys(keys_a, keys_b)

        # Treeview that scrolls headers + both rows together.
        col_ids = ["file"] + [f"c{i:03d}" for i in range(1, len(merged_keys) + 1)]
        tree = ttk.Treeview(sample_frame, columns=col_ids, show="headings", height=3)
        tree.heading("file", text="File / row")
        tree.column("file", width=120, minwidth=100, stretch=False)

        # Map col_id -> merged key
        colid_to_key: dict[str, tuple[str, int]] = {}
        for col_id, k in zip(col_ids[1:], merged_keys):
            colid_to_key[col_id] = k
            tree.heading(col_id, text=_display_field_key(k))
            tree.column(col_id, width=140, minwidth=90, stretch=False)

        yscroll = ttk.Scrollbar(sample_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(sample_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=1, column=0, sticky="nsew")
        yscroll.grid(row=1, column=1, sticky="ns", padx=(6, 0))
        xscroll.grid(row=2, column=0, sticky="we", pady=(6, 0))

        # Two fixed rows: A and B
        tree.insert("", "end", iid="ROW_A", values=["A"] + [""] * len(merged_keys))
        tree.insert("", "end", iid="ROW_B", values=["B"] + [""] * len(merged_keys))

        def _refresh_rows(event=None):  # noqa: ANN001
            try:
                ra = int(row_a_var.get() or 1)
            except Exception:
                ra = 1
            try:
                rb = int(row_b_var.get() or 1)
            except Exception:
                rb = 1

            ha, rva = read_data_row_by_index(info_a, ra)
            hb, rvb = read_data_row_by_index(info_b, rb)

            map_a = _row_map_from_header_and_row(ha, rva)
            map_b = _row_map_from_header_and_row(hb, rvb)

            # Fill values in the merged order; missing columns become empty cells.
            vals_a = [f"A (row {ra})"]
            vals_b = [f"B (row {rb})"]

            for col_id in col_ids[1:]:
                k = colid_to_key[col_id]
                vals_a.append(_truncate_cell(map_a.get(k, "")))
                vals_b.append(_truncate_cell(map_b.get(k, "")))

            tree.item("ROW_A", values=vals_a)
            tree.item("ROW_B", values=vals_b)

        # Update on arrows + Enter
        sp_a.configure(command=_refresh_rows)
        sp_b.configure(command=_refresh_rows)
        sp_a.bind("<Return>", _refresh_rows)
        sp_b.bind("<Return>", _refresh_rows)

        _refresh_rows()

        # Columns compare (headers)
        cols_frame = ttk.LabelFrame(win, text="Columns (headers)", padding=(10, 10))
        cols_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))
        cols_frame.columnconfigure(0, weight=1)
        cols_frame.rowconfigure(0, weight=1)

        cols_text = tk.Text(
            cols_frame,
            height=9,
            wrap="word",
            state="normal",
            font=("Consolas", 9) if sys.platform == "win32" else ("TkFixedFont",),
        )
        cols_text.pack(fill=tk.BOTH, expand=True)

        c = cmp.get("columns", {})
        lines = [
            f'Count: A={c.get("a_count",0)}  |  B={c.get("b_count",0)}',
        ]
        if c.get("same_set") and c.get("same_order"):
            lines.append("Headers: SAME (names + order)")
        elif c.get("same_set") and not c.get("same_order"):
            lines.append("Headers: SAME set, DIFFERENT order")
            mism = c.get("order_mismatches") or []
            if mism:
                lines.append("Order mismatches (col, posA -> posB) [first 50]:")
                for col, pa, pb in mism:
                    lines.append(f" - {col}: {pa} -> {pb}")
        else:
            lines.append("Headers: DIFFERENT")
            if c.get("only_a"):
                lines.append("Only in A:")
                for name in c["only_a"][:200]:
                    lines.append(f" - {name}")
            if c.get("only_b"):
                lines.append("Only in B:")
                for name in c["only_b"][:200]:
                    lines.append(f" - {name}")

        cols_text.insert("1.0", "\n".join(lines))
        cols_text.configure(state="disabled")

    # ── Refresh display ──────────────────────────────────────────────────

    def _refresh_results(self) -> None:
        info = self.info
        if not info:
            return

        self.profile_vars["encoding"].set(info["encoding"])
        self.profile_vars["bom"].set("yes" if info["has_bom"] else "no")
        self.profile_vars["non_ascii"].set("yes" if info["has_non_ascii"] else "no")
        self.profile_vars["delimiter"].set(info["delimiter"])
        self.profile_vars["delimiter_replacement"].set(info.get("delimiter_replacement", ""))
        self.profile_vars["delimiter_in_field_sample"].set("yes" if info.get("delimiter_in_field_sample") else "no")
        self.profile_vars["decimal"].set(info["decimal"])
        self.profile_vars["lineterminator"].set(info["lineterminator"])
        self.profile_vars["quotechar_profile"].set(info["quotechar_profile"])
        self.profile_vars["quotechar_detected"].set(info["quotechar_detected"] or "(not observed)")
        self.profile_vars["quoting"].set(info["quoting"])
        self.profile_vars["escapechar"].set(info["escapechar"] or "")
        self.profile_vars["doublequote"].set("yes" if info["doublequote"] else "no")
        self.profile_vars["date_format"].set(info["date_format"] or "(not detected)")

        self.stats_var.set(
            f'File: {info["file_name"]}  |  '
            f'Size: {_format_bytes(info["file_size"])}  |  '
            f'Rows: {info["total_rows"]:,} (sampled {info.get("rows_sampled",0)}/{info.get("sample_limit","?")})  |  '
            f'Columns: {info["total_columns"]}'
        )

        # Columns table
        self.tree.delete(*self.tree.get_children())
        for col in info["columns"]:
            sample_str = ", ".join(col["sample"])
            self.tree.insert(
                "",
                "end",
                values=(col["name"], col["type"], col["unique"], col["null_count"], sample_str),
            )

        # Preview
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")

        fields = info["fields"]
        col_widths = [len(f) for f in fields]
        for row in info["preview"]:
            for j, f in enumerate(fields):
                val = row.get(f, "")
                col_widths[j] = max(col_widths[j], len(str(val)[:40]))

        hdr_parts = [f.ljust(col_widths[j]) for j, f in enumerate(fields)]
        self.preview_text.insert("end", "  ".join(hdr_parts) + "\n")
        self.preview_text.insert("end", "  ".join("─" * w for w in col_widths) + "\n")

        for row in info["preview"]:
            parts = []
            for j, f in enumerate(fields):
                val = str(row.get(f, ""))[:40]
                parts.append(val.ljust(col_widths[j]))
            self.preview_text.insert("end", "  ".join(parts) + "\n")

        self.preview_text.configure(state="disabled")

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    CSVFormatDiscoverApp().run()


if __name__ == "__main__":
    main()
