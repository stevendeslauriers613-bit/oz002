"""The OZ Store - Cloud Accountant Package Manager (single file)."""
from __future__ import annotations

import io
import re
import urllib.parse
import zipfile
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Constants shared with the app
# --------------------------------------------------------------------------
ACCOUNTANT_ORDER = [
    "Beverages", "Dried Flower", "Edibles", "Extracts Ingested", "Extracts Inhaled",
    "Infused Pre-Rolls", "Milled Flower", "Pre-Rolls", "Seeds", "Topicals",
    "Vapes - Disposables", "Vapes - 510 Threads", "Vapes - Pods", "Accessories",
    "Apparel / Swag", "Gift Cards", "Other Income", "Raffles", "System Fees", "Uniforms",
]

CHECKLIST_ITEMS = [
    ("brinks", "Brinks Book"),
    ("green1", "Greenline: Inventory Summary"),
    ("green2", "Greenline: Financials"),
    ("green3", "Greenline: Payment Types"),
    ("mon1", "Moneris: Counter Sales"),
    ("mon2", "Moneris: Online Sales"),
    ("ocs1", "OCS: Account Statement"),
    ("ocs2", "OCS: Invoices"),
    ("gas1", "Gas Cards: Details"),
    ("gas2", "Gas Cards: Summary"),
]

EXPENSE_FIELDS = [
    ("rent", "Rent"), ("payroll", "Payroll"), ("utilities", "Utilities"), ("insurance", "Insurance"),
    ("systemfees", "System Fees"), ("repairs", "Repairs"), ("other", "Other"),
]

DOC_NAME = dict(CHECKLIST_ITEMS)

# Columns that must never be turned into numbers (IDs, SKUs, names...)
_TEXT_COLS = re.compile(
    r"\b(id|sku|upc|barcode|name|product|description|category|brand|supplier|vendor|store|date|code|plu)\b",
    re.I,
)
MAX_SHEET_ROWS = 50_000

# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def parse_money(value) -> float:
    """'$1,234.50', '(200)', '1 234,50' -> float. Blank/invalid -> 0.0."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 0.0 if value != value else float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    neg = (s.startswith("(") and s.endswith(")")) or s.startswith("-")
    s = s.replace(" ", "").replace("\u00a0", "")
    if "," in s and "." not in s and re.search(r",\d{1,2}$", s):
        s = s.replace(",", ".")  # European / Quebec decimal comma
    s = re.sub(r"[^\d.]", "", s)
    try:
        x = float(s)
    except ValueError:
        return 0.0
    return -x if neg else x


def total_expenses(payload: dict) -> float:
    exp = payload.get("expenses") or {}
    return sum(parse_money(exp.get(k)) for k, _ in EXPENSE_FIELDS)


def missing_docs(payload: dict) -> list[str]:
    chk = payload.get("checklist") or {}
    files = payload.get("files") or {}
    return [label for key, label in CHECKLIST_ITEMS if not (chk.get(key) or files.get(key))]


def checklist_done(payload: dict) -> int:
    return len(CHECKLIST_ITEMS) - len(missing_docs(payload))


# --------------------------------------------------------------------------
# Reading uploaded files
# --------------------------------------------------------------------------
def _read_ragged(text: str) -> pd.DataFrame:
    """Fallback for exports with title lines above the header (rows of different widths)."""
    import csv

    rows = list(csv.reader(io.StringIO(text)))
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        return pd.DataFrame()
    rows = [r + [""] * (width - len(r)) for r in rows]
    fill = [sum(1 for v in r if str(v).strip()) for r in rows]
    best = max(fill[:25])
    hdr = next(i for i, n in enumerate(fill[:25]) if n == best)
    df = pd.DataFrame(rows[hdr + 1:], columns=[str(v).strip() or f"Column {j + 1}" for j, v in enumerate(rows[hdr])])
    return df.replace(r"^\s*$", pd.NA, regex=True)


def _read_csv(data: bytes) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "cp1252"):
        try:
            text = data.decode(enc)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        try:
            df = pd.read_csv(io.StringIO(text))
            if df.shape[1] == 1:  # maybe ; or tab separated
                alt = pd.read_csv(io.StringIO(text), sep=None, engine="python")
                if alt.shape[1] > 1:
                    df = alt
            return df
        except (pd.errors.ParserError, pd.errors.EmptyDataError):
            return _read_ragged(text)
    raise ValueError(f"Could not decode CSV: {last_err}")


def read_table(name: str, data: bytes) -> pd.DataFrame | None:
    """CSV / Excel -> DataFrame. Anything else (PDF, images) -> None."""
    ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
    if ext in ("csv", "txt"):
        df = _read_csv(data)
    elif ext in ("xlsx", "xlsm", "xls"):
        df = pd.read_excel(io.BytesIO(data))
    else:
        return None
    return df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)


def _promote_header(df: pd.DataFrame, needle: str = "category", scan: int = 25) -> pd.DataFrame:
    """Some exports have title lines above the real header; find it."""
    if any(str(c).strip().lower() == needle for c in df.columns):
        return df
    for i in range(min(scan, len(df))):
        row = [str(v).strip() for v in df.iloc[i].tolist()]
        if needle in [r.lower() for r in row]:
            out = df.iloc[i + 1:].copy()
            out.columns = row
            return out.reset_index(drop=True)
    return df


def _coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]) or _TEXT_COLS.search(str(c)):
            continue
        s = df[c].astype(str).str.strip()
        filled = df[c].notna() & (s != "") & (s.str.lower() != "nan")
        if not filled.any():
            continue
        cleaned = (
            s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
            .str.replace(r"[$,%\s]", "", regex=True)
        )
        num = pd.to_numeric(cleaned, errors="coerce")
        if num[filled].notna().mean() >= 0.9:
            df[c] = num
    return df


def numeric_cols(df: pd.DataFrame) -> list[str]:
    return [
        c for c in df.columns
        if pd.api.types.is_numeric_dtype(df[c])
        and not pd.api.types.is_bool_dtype(df[c])
        and not _TEXT_COLS.search(str(c))
    ]


def clean_inventory(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Order categories like the accountant wants, drop subtotal rows, fix numbers."""
    notes: list[str] = []
    df = df.dropna(how="all").dropna(axis=1, how="all").copy()
    df = _promote_header(df)
    df.columns = [str(c).strip() for c in df.columns]

    cat_col = next((c for c in df.columns if c.lower() == "category"), None)
    if cat_col is None:
        notes.append("No 'Category' column found, so rows were kept in their original order.")
        return _coerce_numeric(df), notes
    if cat_col != "Category":
        df = df.rename(columns={cat_col: "Category"})

    cats = df["Category"].astype(str).str.strip()
    drop = cats.str.startswith(">") | cats.str.lower().isin(
        ["", "nan", "none", "total", "totals", "grand total"]
    )
    df, cats = df[~drop].copy(), cats[~drop]

    rank = {c.casefold(): i for i, c in enumerate(ACCOUNTANT_ORDER)}
    order = cats.str.casefold().map(rank)
    unknown = sorted(cats[order.isna()].unique())
    if unknown:
        notes.append("Categories not in the accountant order (placed at the end): " + ", ".join(unknown))
    df["Category"] = cats
    df["_o"] = order.fillna(len(ACCOUNTANT_ORDER))
    df = df.sort_values("_o", kind="stable").drop(columns="_o").reset_index(drop=True)
    return _coerce_numeric(df), notes


# --------------------------------------------------------------------------
# Totals read from uploaded documents (Greenline, Moneris, OCS, gas cards)
# --------------------------------------------------------------------------
FACT_COLS = {
    "green2": ["net sales", "total net sales", "net sale", "total sales", "gross sales", "sales", "total"],
    "green3": ["total", "total amount", "amount", "payment total"],
    "mon1": ["net sales", "total sales", "net amount", "sales", "amount", "total"],
    "mon2": ["net sales", "total sales", "net amount", "sales", "amount", "total"],
    "ocs1": ["balance", "amount due", "closing balance", "total"],
    "ocs2": ["invoice total", "invoice amount", "total", "net amount", "amount", "subtotal"],
    "gas1": ["total", "amount", "spend", "cost", "net amount"],
    "gas2": ["total", "amount", "spend", "cost", "net amount"],
}
_TOTAL_ROW = re.compile(r"^\s*(grand\s+)?(sub)?totals?\b", re.I)


def _first_text_col(df: pd.DataFrame):
    for c in df.columns:
        if not pd.api.types.is_numeric_dtype(df[c]):
            return c
    return None


def extract_total(cat: str, df: pd.DataFrame):
    """Find the one number that matters in a document: -> {col, total, rows} or None."""
    prefs = FACT_COLS.get(cat)
    if not prefs or df is None or df.empty:
        return None
    df = df.dropna(how="all").reset_index(drop=True)
    cols = {str(c).strip().casefold(): c for c in df.columns}
    label_col = _first_text_col(df)
    labels = df[label_col].astype(str) if label_col is not None else None
    total_rows = labels.str.match(_TOTAL_ROW) if labels is not None else None
    nums = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c])]

    def pick():
        for p in prefs:                       # exact header match first
            c = cols.get(p)
            if c is not None and c in nums:
                return c
        for p in prefs:                       # then "contains"
            for key, c in cols.items():
                if p in key and c in nums:
                    return c
        return None

    col = pick()
    if col is not None:
        s_ = df[col]
        if total_rows is not None and total_rows.any():
            v = s_[total_rows].dropna()
            total = float(v.iloc[-1]) if len(v) else float(s_[~total_rows].sum(skipna=True))
            rows = int((~total_rows).sum())
        else:
            total, rows = float(s_.sum(skipna=True)), int(s_.notna().sum())
        return {"col": str(col), "total": total, "rows": rows}

    # label / value layout, e.g. "Net Sales | 12,345.67"
    if label_col is not None:
        for p in prefs:
            hit = labels.str.strip().str.casefold() == p
            if hit.any():
                row = df[hit].iloc[0]
                for c in df.columns:
                    if c == label_col:
                        continue
                    v = parse_money(row[c])
                    if v:
                        return {"col": str(row[label_col]).strip(), "total": v, "rows": 1}
    # No recognised heading: use the Total row if there is one, or the only number column
    if col is None and total_rows is not None and total_rows.any() and nums:
        row = df[total_rows].iloc[-1]
        vals = [(c, row[c]) for c in nums if pd.notna(row[c]) and float(row[c]) != 0]
        if vals:
            c, v = vals[-1]
            return {"col": f"Total row, column {c}", "total": float(v), "rows": int((~total_rows).sum())}
    if col is None and cat in ('mon1', 'mon2', 'ocs2', 'gas1', 'gas2', 'green3') and len(nums) == 1 \
            and not re.search(r'qty|quantity|litre|liter|unit|count', str(nums[0]), re.I):
        c = nums[0]
        return {"col": str(c), "total": float(df[c].sum(skipna=True)), "rows": int(df[c].notna().sum())}

    return None


def summarize_facts(facts: list) -> dict:
    """facts: [{cat, file, col, total, rows}] -> the figures the financial summary uses."""
    def tot(*cats):
        return sum(f["total"] for f in facts if f["cat"] in cats)
    green2, mon, ocs = tot("green2"), tot("mon1", "mon2"), tot("ocs2")
    out = {"sales": 0.0, "sales_src": "", "cogs": 0.0, "cogs_src": ""}
    if green2 > 0:
        out["sales"], out["sales_src"] = green2, "Greenline: Financials"
    elif mon > 0:
        out["sales"], out["sales_src"] = mon, "Moneris sales reports"
    if ocs > 0:
        out["cogs"], out["cogs_src"] = ocs, "OCS invoices"
    return out


# --------------------------------------------------------------------------
# Excel styling helpers
# --------------------------------------------------------------------------
_FONT = "Arial"
_HDR_FILL = PatternFill("solid", fgColor="065F46")
_HDR_FONT = Font(name=_FONT, bold=True, color="FFFFFF", size=10)
_TOT_FILL = PatternFill("solid", fgColor="D1FAE5")
_BAD_FILL = PatternFill("solid", fgColor="FEE2E2")
_BODY = Font(name=_FONT, size=10)
_BOLD = Font(name=_FONT, size=10, bold=True)
_THIN = Border(bottom=Side(style="thin", color="E5E7EB"))
MONEY = '#,##0.00;[Red]-#,##0.00'
INT = '#,##0;[Red]-#,##0'


def _cell_value(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def _title(ws, row: int, text: str, size: int = 12):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=_FONT, bold=True, size=size, color="065F46")


def _header(ws, row: int, headers: list[str]):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font, c.fill = _HDR_FONT, _HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _put(ws, row: int, col: int, value, *, bold=False, fill=None):
    c = ws.cell(row=row, column=col, value=_cell_value(value))
    c.font = _BOLD if bold else _BODY
    c.border = _THIN
    if fill:
        c.fill = fill
    if isinstance(c.value, bool):
        pass
    elif isinstance(c.value, int):
        c.number_format = INT
    elif isinstance(c.value, float):
        c.number_format = MONEY
    return c


def _write_df(ws, df: pd.DataFrame, row: int = 1, totals: bool = True) -> int:
    """Write a styled table; returns the next free row."""
    df = df.head(MAX_SHEET_ROWS)
    _header(ws, row, [str(c) for c in df.columns])
    for i, rec in enumerate(df.itertuples(index=False), start=row + 1):
        for j, v in enumerate(rec, 1):
            _put(ws, i, j, v)
    r = row + len(df) + 1
    nums = numeric_cols(df) if totals else []
    if nums:
        for j, c in enumerate(df.columns, 1):
            if j == 1:
                _put(ws, r, j, "TOTAL", bold=True, fill=_TOT_FILL)
            elif c in nums:
                _put(ws, r, j, float(df[c].sum(skipna=True)), bold=True, fill=_TOT_FILL)
            else:
                _put(ws, r, j, None, fill=_TOT_FILL)
        r += 1
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
    return r + 1


def _block(ws, row: int, title: str, headers: list[str], rows: list[list], bad_when=None) -> int:
    _title(ws, row, title)
    _header(ws, row + 1, headers)
    for i, rec in enumerate(rows, start=row + 2):
        bad = bad_when(rec) if bad_when else False
        for j, v in enumerate(rec, 1):
            _put(ws, i, j, v, fill=_BAD_FILL if bad else None)
    return row + len(rows) + 3


def _autowidth(ws, cap: int = 48):
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            n = len(f"{c.value:,.2f}") if isinstance(c.value, float) else len(str(c.value))
            widths[c.column] = max(widths.get(c.column, 0), min(n, cap))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(10, w + 2)


def _sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]\*\?/\\:]", "", name).strip()[:31] or "Sheet"
    out, n = base, 2
    while out.lower() in used:
        suffix = f" ({n})"
        out, n = base[: 31 - len(suffix)] + suffix, n + 1
    used.add(out.lower())
    return out


def _to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _inventory_frame(payload: dict) -> pd.DataFrame | None:
    recs = payload.get("inventory_data")
    if not recs:
        return None
    df = pd.DataFrame(recs)
    df = df.replace({"": None})
    return _coerce_numeric(df)


def _num(v) -> float:
    try:
        return float(re.sub(r"[^0-9.\-]", "", str(v))) if v is not None else 0.0
    except ValueError:
        return 0.0


# --------------------------------------------------------------------------
# Report builders
# --------------------------------------------------------------------------
def inventory_workbook(df: pd.DataFrame, title: str = "Inventory") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_name(title, set())
    _write_df(ws, df)
    _autowidth(ws)
    return _to_bytes(wb)


def build_store_workbook(
    store: str, period: str, payload: dict, tables: dict[str, tuple[str, pd.DataFrame]],
    doc_log: list[list] | None = None,
) -> bytes:
    """One store, one month: Summary + Inventory + one sheet per tabular document + Documents log."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    used = {"summary"}

    ws["A1"] = f"The OZ Store - {store}"
    ws["A1"].font = Font(name=_FONT, bold=True, size=16, color="065F46")
    ws["A2"] = f"Period: {period}"
    ws["A3"] = f"Generated: {datetime.now():%Y-%m-%d %H:%M}"
    for a_ in ("A2", "A3"):
        ws[a_].font = _BODY
    missing = missing_docs(payload)
    ws["A4"] = ("Missing documents: " + ", ".join(missing)) if missing else "All documents received"
    ws["A4"].font = Font(name=_FONT, bold=True, size=10, color="B91C1C" if missing else "047857")

    files = payload.get("files") or {}
    notes = payload.get("notes") or {}
    chk = payload.get("checklist") or {}

    def names(key):
        v = files.get(key) or []
        if isinstance(v, dict):
            v = [v]
        return ", ".join(f.get("name", "") for f in v)

    rows = [
        [label, "Received" if (chk.get(key) or files.get(key)) else "MISSING", names(key), notes.get(key, "")]
        for key, label in CHECKLIST_ITEMS
    ]
    r = _block(ws, 6, "Document checklist", ["Document", "Status", "File(s)", "Note"], rows,
               bad_when=lambda rec: rec[1] == "MISSING")

    cooh_text = payload.get("cooh_text", "")
    cooh_amt = parse_money(cooh_text) if str(cooh_text).strip() else None
    r = _block(ws, r, "Cash & chargebacks", ["Item", "Value"], [
        ["Cash on hand", cooh_amt if cooh_amt is not None else "Not entered"],
        ["Cash on hand date", payload.get("cooh_date") or "Not entered"],
        ["Moneris chargebacks", payload.get("chargebacks_text") or "Nothing to report."],
    ])

    fin = financials(payload)
    inc = fin["inc"]
    exp = payload.get("expenses") or {}
    exp_rows = [[label, parse_money(exp.get(key)), "Entered by you"] for key, label in EXPENSE_FIELDS]
    exp_rows += [[x["label"], x["amount"], x["src"]] for x in inc["extra_expenses"]]
    exp_rows.append(["TOTAL", inc["total_expenses"], ""])
    r = _block(ws, r, "Operating expenses", ["Expense", "Amount", "Source"], exp_rows)

    if fin["sales"] > 0:
        r = _block(ws, r, "Financial summary", ["Measure", "Value"], [
            [f"Total sales - {fin['sales_src']}", fin["sales"]],
            [f"Cost of goods sold - {fin['cogs_src']}", fin["cogs"]],
            ["Gross profit", fin["gross"]],
            ["Gross margin %", round(fin["gross_margin"], 1)],
            ["Operating expenses (incl. card fees and fuel)", fin["expenses"]],
            ["Net operating income", fin["noi"]],
            ["Operating margin %", round(fin["op_margin"], 1)],
        ])

    facts = payload.get("doc_facts") or []
    if facts:
        r = _block(ws, r, "Figures read from uploaded documents",
                   ["Document", "File", "Column / line used", "Rows", "Total"],
                   [[DOC_NAME.get(f["cat"], f["cat"]), f["file"], f["col"], f["rows"], f["total"]] for f in facts])

    inv = _inventory_frame(payload)
    if inv is not None:
        nums = numeric_cols(inv)[:12]
        inv_rows = [["Categories", int(inv["Category"].nunique()) if "Category" in inv else len(inv)],
                    ["Line items", len(inv)]]
        inv_rows += [[c if c.lower().startswith("total") else f"Total {c}", float(inv[c].sum(skipna=True))] for c in nums]
        r = _block(ws, r, "Inventory totals", ["Measure", "Value"], inv_rows)
        ws2 = wb.create_sheet(_sheet_name("Inventory", used))
        _write_df(ws2, inv)
        _autowidth(ws2)

    for key, (label, df) in tables.items():
        if len(df) > MAX_SHEET_ROWS:
            df = df.head(MAX_SHEET_ROWS)
        sheet = wb.create_sheet(_sheet_name(label, used))
        _write_df(sheet, df, totals=True)
        _autowidth(sheet)

    if doc_log:
        wd = wb.create_sheet(_sheet_name("Documents", used))
        _header(wd, 1, ["Category", "File", "Size (KB)", "In this report"])
        for i, rec in enumerate(doc_log, start=2):
            for j, v in enumerate(rec, 1):
                _put(wd, i, j, v, fill=_BAD_FILL if "NOT FOUND" in str(rec[-1]) else None)
        _autowidth(wd)

    _income_sheets(wb, store, period, inc, used)
    _autowidth(ws)
    return _to_bytes(wb)


def build_consolidated_workbook(period: str, payloads: dict[str, dict]) -> bytes:
    """All stores for one month: Overview, Expenses matrix, All Inventory, Category Totals."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Overview"
    used = {"overview"}
    ws["A1"] = f"The OZ Store - All stores - {period}"
    ws["A1"].font = Font(name=_FONT, bold=True, size=16, color="065F46")
    ws["A2"] = f"Generated: {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = _BODY

    rows = []
    for store, p in payloads.items():
        cooh = str(p.get("cooh_text", "")).strip()
        fin = financials(p)
        rows.append([
            store,
            f"{checklist_done(p)}/{len(CHECKLIST_ITEMS)}",
            parse_money(cooh) if cooh else "Not entered",
            p.get("cooh_date") or "",
            p.get("chargebacks_text") or "",
            fin["expenses"],
            fin["sales"],
            fin["noi"],
            ", ".join(missing_docs(p)) or "-",
        ])
    if rows:
        money = [r_ for r_ in rows]
        rows.append(["ALL STORES", "", "", "", "",
                     sum(r_[5] for r_ in money), sum(r_[6] for r_ in money), sum(r_[7] for r_ in money), ""])
    _block(ws, 4, "Store overview",
           ["Store", "Checklist", "Cash on hand", "Cash date", "Chargebacks", "Total expenses", "Total sales",
            "Net operating income", "Missing"],
           rows, bad_when=lambda rec: rec[8] not in ("-", ""))
    _autowidth(ws)

    stores = list(payloads)
    wx = wb.create_sheet(_sheet_name("Expenses", used))
    _header(wx, 1, ["Expense"] + stores + ["All stores"])
    for i, (key, label) in enumerate(EXPENSE_FIELDS, start=2):
        vals = [parse_money((payloads[s].get("expenses") or {}).get(key)) for s in stores]
        _put(wx, i, 1, label)
        for j, v in enumerate(vals, 2):
            _put(wx, i, j, v)
        _put(wx, i, len(stores) + 2, sum(vals), bold=True)
    extra_labels = []
    for s_ in stores:
        for x in financials(payloads[s_])["inc"]["extra_expenses"]:
            lab = "Fuel - gas cards" if x["label"].startswith("Fuel") else x["label"].split(" - ")[0] + " - " + x["label"].split(" - ")[1].split(" (")[0]
            if lab not in extra_labels:
                extra_labels.append(lab)
    for k, lab in enumerate(extra_labels):
        row_ = len(EXPENSE_FIELDS) + 2 + k
        _put(wx, row_, 1, lab)
        tot_ = 0.0
        for j, s_ in enumerate(stores, 2):
            v = 0.0
            for x in financials(payloads[s_])["inc"]["extra_expenses"]:
                l2 = "Fuel - gas cards" if x["label"].startswith("Fuel") else x["label"].split(" - ")[0] + " - " + x["label"].split(" - ")[1].split(" (")[0]
                if l2 == lab:
                    v += x["amount"]
            _put(wx, row_, j, v)
            tot_ += v
        _put(wx, row_, len(stores) + 2, tot_, bold=True)
    tr = len(EXPENSE_FIELDS) + 2 + len(extra_labels)
    _put(wx, tr, 1, "TOTAL", bold=True, fill=_TOT_FILL)
    for j, s in enumerate(stores, 2):
        _put(wx, tr, j, financials(payloads[s])["expenses"], bold=True, fill=_TOT_FILL)
    _put(wx, tr, len(stores) + 2, sum(financials(payloads[s])["expenses"] for s in stores), bold=True, fill=_TOT_FILL)
    wx.freeze_panes = "B2"
    _autowidth(wx)

    frames = []
    for s in stores:
        f = _inventory_frame(payloads[s])
        if f is not None:
            f.insert(0, "Store", s)
            frames.append(f)
    if frames:
        allinv = pd.concat(frames, ignore_index=True)
        wa = wb.create_sheet(_sheet_name("All Inventory", used))
        _write_df(wa, allinv, totals=True)
        _autowidth(wa)
        nums = numeric_cols(allinv)
        if "Category" in allinv and nums:
            tot = allinv.groupby("Category", sort=False)[nums].sum()
            rank = {c.casefold(): i for i, c in enumerate(ACCOUNTANT_ORDER)}
            tot = tot.assign(_o=[rank.get(str(c).casefold(), 999) for c in tot.index]) \
                     .sort_values("_o", kind="stable").drop(columns="_o").reset_index()
            wc = wb.create_sheet(_sheet_name("Category Totals", used))
            _write_df(wc, tot, totals=True)
            _autowidth(wc)
    return _to_bytes(wb)


def build_zip(entries: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for arcname, data in entries:
            z.writestr(arcname, data)
    return buf.getvalue()


# ==========================================================================



def describe_columns(df: pd.DataFrame, limit: int = 5) -> list[dict]:
    """Totals of the biggest number columns, for the 'what did we read' table."""
    if df is None or df.empty:
        return []
    df = df.dropna(how="all")
    lc = _first_text_col(df)
    if lc is not None:
        df = df[~df[lc].astype(str).str.match(_TOTAL_ROW)]
    out = []
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c]) and not _TEXT_COLS.search(str(c)):
            v = float(df[c].sum(skipna=True))
            if v != 0:
                out.append({"name": str(c), "total": v})
    return sorted(out, key=lambda d: -abs(d["total"]))[:limit]


# ==========================================================================
# DOCUMENT READERS + INCOME STATEMENT
# (Greenline financials / payment types, Moneris PDFs, OCS invoices, gas card PDFs)
# ==========================================================================
try:
    from pypdf import PdfReader as _PdfReader
except Exception:          # the report still works without it, PDFs are just attached
    _PdfReader = None

_MONEY_RE = r"-?\(?[\d,]*\.\d{2}\)?"


def _mf(s) -> float:
    """'3,259.85', '-1,222.38', '(12.00)', '35.74CR' -> float (CR = credit = negative)."""
    if s is None:
        return 0.0
    t = str(s).strip()
    cr = t.upper().endswith("CR")
    t = re.sub(r"(?i)cr$", "", t)
    v = parse_money(t)
    return -abs(v) if cr else v


def pdf_text(data: bytes) -> str:
    if _PdfReader is None:
        return ""
    try:
        r = _PdfReader(io.BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in r.pages)
    except Exception:
        return ""


def _col(df: pd.DataFrame, *names):
    cols = {str(c).strip().casefold(): c for c in df.columns}
    for n in names:
        if n in cols:
            return cols[n]
    for n in names:
        for k, c in cols.items():
            if k.startswith(n):
                return c
    return None


def _colsum(df, col) -> float:
    if col is None:
        return 0.0
    s = pd.to_numeric(df[col].astype(str).str.replace(r"[$,\s]", "", regex=True), errors="coerce")
    return float(s.sum(skipna=True))


def _daily(df, date_col, val_col):
    if date_col is None or val_col is None:
        return []
    d = pd.to_datetime(df[date_col], errors="coerce")
    v = pd.to_numeric(df[val_col].astype(str).str.replace(r"[$,\s]", "", regex=True), errors="coerce")
    out = [(x.strftime("%Y-%m-%d"), float(y)) for x, y in zip(d, v) if pd.notna(x) and pd.notna(y)]
    return out


def read_greenline_financials(df):
    net = _col(df, "net sales")
    if net is None:
        return None
    date = _col(df, "date")
    m = {
        "gross": _colsum(df, _col(df, "gross sales")), "refunds": _colsum(df, _col(df, "refunds")),
        "discounts": _colsum(df, _col(df, "discounts")), "net": _colsum(df, net),
        "hst": _colsum(df, _col(df, "hst")), "total": _colsum(df, _col(df, "total")),
        "cost": _colsum(df, _col(df, "cost")), "profit": _colsum(df, _col(df, "profit")),
        "giftcards": _colsum(df, _col(df, "gift cards")),
    }
    m["days"] = int(df[net].notna().sum())
    m["daily"] = _daily(df, date, net)
    return m


def read_payment_types(df):
    date = _col(df, "payments.completedate", "date", "completedate")
    cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c]) and c != date]
    if date is None or len(cols) < 2:
        return None
    tender = {str(c).strip().title(): _colsum(df, c) for c in cols}
    return {"tender": {k: v for k, v in tender.items() if abs(v) > 0.005}, "days": int(len(df))}


def read_moneris_statement(text: str):
    if "Sales Summary by Card Type" not in text:
        return None
    g = re.search(r"Grand Total\s+(\d[\d,]*)\s+(" + _MONEY_RE + r")\s+(\d[\d,]*)\s+(" + _MONEY_RE + r")\s+(\d[\d,]*)\s+(" + _MONEY_RE + r")", text)
    if not g:
        return None
    fees = re.search(r"Total fees\s+(" + _MONEY_RE + ")", text)
    dep = re.search(r"Total Deposits\s+\S+\s+\S+\s+(" + _MONEY_RE + ")", text)
    when = re.search(r"Statement Date:\s*([A-Za-z]+ \d{4})", text)
    return {
        "gross": _mf(g.group(2)), "returns": abs(_mf(g.group(4))), "net": _mf(g.group(6)), "items": int(g.group(5).replace(",", "")),
        "fees": abs(_mf(fees.group(1))) if fees else 0.0, "deposits": _mf(dep.group(1)) if dep else 0.0,
        "statement": when.group(1) if when else "",
    }


def read_ocs_invoice(data: bytes):
    try:
        from openpyxl import load_workbook
        ws = load_workbook(io.BytesIO(data), data_only=True, read_only=True).active
    except Exception:
        return None
    vals = {}
    number = date_ = ""
    labels = {"SALES SUBTOTAL AMOUNT": "subtotal", "FREIGHT CHARGES": "freight", "SALES TAX": "tax", "CAD TOTAL": "total"}
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        for i, c in enumerate(cells):
            s = str(c).strip()
            if s.upper() in labels and i + 1 < len(cells):
                vals[labels[s.upper()]] = _mf(cells[i + 1])
            m = re.match(r"Invoice\s+(\S+)", s)
            if m and not number:
                number = m.group(1)
            if not date_ and re.fullmatch(r"\d{4}-\d\d-\d\d", s):
                date_ = s
    if "total" not in vals or "subtotal" not in vals:
        return None
    return {"number": number, "date": date_, "subtotal": vals.get("subtotal", 0.0), "freight": vals.get("freight", 0.0),
            "tax": vals.get("tax", 0.0), "total": vals.get("total", 0.0)}


def read_fuel_statement(text: str):
    if "TOTAL CURRENT CHARGES" in text:                      # monthly statement summary
        cc = re.search(r"TOTAL CURRENT CHARGES\s+(" + _MONEY_RE + ")", text)
        gst = re.search(r"NET GST/HST\s+(" + _MONEY_RE + ")", text)
        lit = re.search(r"TOTAL\s+(\d+)\s+([\d,]+\.\d)\s", text)
        disc = re.search(r"DISCOUNT\s+(" + _MONEY_RE + r"CR)", text)
        return {"kind": "summary", "total": _mf(cc.group(1)) if cc else 0.0, "gst": _mf(gst.group(1)) if gst else 0.0,
                "litres": _mf(lit.group(2)) if lit else 0.0, "discount": abs(_mf(disc.group(1))) if disc else 0.0}
    nets = re.findall(r"NET GST/HST\s+(" + _MONEY_RE + r")\s+NET TOTAL\s+(" + _MONEY_RE + ")", text)
    if nets:
        lit = re.search(r"TOTAL\s+--\s+(\d+)\s+(\d+)", text)
        return {"kind": "details", "total": sum(_mf(n[1]) for n in nets), "gst": sum(_mf(n[0]) for n in nets),
                "litres": float(int(lit.group(1)) + int(lit.group(2))) if lit else 0.0, "cards": len(nets)}
    return None


def analyze_doc(cat: str, name: str, data: bytes, df) -> dict:
    """Read one uploaded file. Always returns {cat,file,kind,lines,status,...}."""
    info = {"cat": cat, "file": name, "kind": "attached", "lines": [], "status": "attached"}
    ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""

    def done(kind, lines, **kw):
        info.update(kind=kind, lines=lines, status="read", **kw)
        return info

    try:
        if ext == "pdf":
            text = pdf_text(data)
            if cat in ("mon1", "mon2"):
                m = read_moneris_statement(text)
                if m:
                    return done("moneris", [["Net sales", m["net"]], ["Gross sales", m["gross"]], ["Returns", m["returns"]],
                                            ["Card fees", m["fees"]], ["Deposited", m["deposits"]]], m=m)
            if cat in ("gas1", "gas2"):
                f = read_fuel_statement(text)
                if f:
                    lines = [["Fuel bought (incl. HST)", f["total"]], ["HST on fuel", f["gst"]], ["Litres", f["litres"]]]
                    return done("fuel_" + f["kind"], lines, f=f)
            return info
        if ext in ("xlsx", "xlsm") and cat == "ocs2":
            inv = read_ocs_invoice(data)
            if inv:
                return done("ocs_invoice", [["Invoice " + inv["number"], inv["date"]], ["Product + freight", inv["subtotal"] + inv["freight"]],
                                            ["HST", inv["tax"]], ["Invoice total", inv["total"]]], inv=inv)
        if df is not None and not df.empty:
            if cat == "green2":
                g = read_greenline_financials(df)
                if g:
                    return done("greenline_fin", [["Net sales", g["net"]], ["Cost of goods sold", g["cost"]], ["Gross sales", g["gross"]],
                                                  ["Refunds", g["refunds"]], ["Discounts", g["discounts"]], ["HST collected", g["hst"]],
                                                  ["Days", g["days"]]], g=g)
            if cat == "green3":
                p = read_payment_types(df)
                if p:
                    return done("payment_types", [[k, v] for k, v in p["tender"].items()], p=p)
            got = extract_total(cat, df)
            if got:
                return done("table_total", [[got["col"], got["total"]]], got=got, rows=int(len(df)))
            info.update(status="nothing", rows=int(len(df)), headers=[str(c) for c in list(df.columns)[:8]], cols=describe_columns(df))
    except Exception:
        pass
    return info


def _inv_totals(payload: dict) -> dict:
    rows = payload.get("inventory_data") or []
    out = {"opening": 0.0, "purchased": 0.0, "sold": 0.0, "closing": 0.0, "cost": 0.0}
    if not rows:
        return out
    heads = list(rows[0].keys())

    def find(*names):
        for n in names:
            for h in heads:
                if h and str(h).strip().casefold() == n:
                    return h
        return None
    hs = {k: find(k) for k in ("opening", "purchased", "sold", "closing")}
    cost_h = next((h for h in heads if h and "cost" in str(h).lower()), None)
    sold_h = hs["sold"] or next((h for h in heads if h and "sold" in str(h).lower()), None)
    for r in rows:
        for k in ("opening", "purchased", "closing"):
            if hs[k]:
                out[k] += _num(r.get(hs[k]))
        if sold_h:
            out["sold"] += _num(r.get(sold_h))
        elif len(r) > 1:
            first = next((k for k in r if k != "Category"), None)
            out["sold"] += _num(r.get(first)) if first else 0.0
        if cost_h:
            out["cost"] += _num(r.get(cost_h))
    return out


def _chk(label, a, b, tol_pct=1.0, tol_abs=25.0, a_name="", b_name=""):
    diff = a - b
    big = max(abs(a), abs(b), 1.0)
    ok = abs(diff) <= max(tol_abs, big * tol_pct / 100)
    return {"label": label, "a": a, "b": b, "diff": diff, "ok": ok, "a_name": a_name, "b_name": b_name}


def build_income(payload: dict) -> dict:
    infos = payload.get("docs_info") or []
    by = lambda kind: [i for i in infos if i["kind"] == kind]
    gl, pay, mon = by("greenline_fin"), by("payment_types"), by("moneris")
    ocs, fs, fd = by("ocs_invoice"), by("fuel_summary"), by("fuel_details")
    inv = _inv_totals(payload)
    inc = {"docs_used": [], "checks": []}

    # ---- revenue
    rev = {"gross": 0.0, "refunds": 0.0, "discounts": 0.0, "net": 0.0, "hst": 0.0, "src": "", "from_docs": False}
    if gl:
        g = [i["g"] for i in gl]
        rev.update(gross=sum(x["gross"] for x in g), refunds=sum(x["refunds"] for x in g), discounts=sum(x["discounts"] for x in g),
                   net=sum(x["net"] for x in g), hst=sum(x["hst"] for x in g), src="Greenline: Financials", from_docs=True)
    elif mon:
        m = [i["m"] for i in mon]
        rev.update(gross=sum(x["gross"] for x in m), refunds=sum(x["returns"] for x in m), net=sum(x["net"] for x in m),
                   src="Moneris statements (card sales only)", from_docs=True)
    else:
        facts = payload.get("doc_facts") or []
        d = summarize_facts(facts)
        if d["sales"] > 0:
            rev.update(net=d["sales"], src=d["sales_src"], from_docs=True)
        elif inv["sold"] > 0:
            rev.update(net=inv["sold"], src="Inventory 'sold' column")
    inc["revenue"] = rev

    # ---- cost of goods
    cogs = {"amount": 0.0, "src": "", "estimated": False}
    if gl and sum(i["g"]["cost"] for i in gl) > 0:
        cogs.update(amount=sum(i["g"]["cost"] for i in gl), src="Greenline: Financials (cost)")
    elif inv["cost"] > 0:
        cogs.update(amount=inv["cost"], src="Inventory cost column")
    elif rev["from_docs"] and inv["sold"] > 0:
        cogs.update(amount=inv["sold"], src="Inventory sold, valued at cost")
    elif rev["net"] > 0:
        cogs.update(amount=rev["net"] * 0.80, src="Estimated at 80% of sales", estimated=True)
    inc["cogs"] = cogs
    gross = rev["net"] - cogs["amount"]
    inc["gross_profit"] = gross
    inc["gross_margin"] = gross / rev["net"] * 100 if rev["net"] > 0 else 0.0

    # ---- expenses
    exp_map = payload.get("expenses") or {}
    lines = [{"label": label, "amount": parse_money(exp_map.get(k)), "src": "Entered by you"} for k, label in EXPENSE_FIELDS]
    lines = [x for x in lines if x["amount"] > 0]
    extra = []
    for i in mon:
        m = i["m"]
        if m["fees"] > 0:
            extra.append({"label": "Card processing fees - " + ("in-store" if i["cat"] == "mon1" else "online") + " (Moneris)",
                          "amount": m["fees"], "src": i["file"]})
    if fs:
        fuel = sum(i["f"]["total"] - i["f"]["gst"] for i in fs)
        lit = sum(i["f"]["litres"] for i in fs)
    else:
        fuel = sum(i["f"]["total"] - i["f"]["gst"] for i in fd)
        lit = sum(i["f"]["litres"] for i in fd)
    if fuel > 0:
        extra.append({"label": "Fuel - gas cards (excl. HST)", "amount": fuel, "src": "Gas card statement", "litres": lit})
    inc["expense_lines"] = lines + extra
    inc["extra_expenses"] = extra
    inc["manual_expenses"] = sum(x["amount"] for x in lines)
    inc["total_expenses"] = inc["manual_expenses"] + sum(x["amount"] for x in extra)
    inc["noi"] = gross - inc["total_expenses"]
    inc["op_margin"] = inc["noi"] / rev["net"] * 100 if rev["net"] > 0 else 0.0

    # ---- supporting schedules
    tender = {}
    for i in pay:
        for k, v in i["p"]["tender"].items():
            tender[k] = tender.get(k, 0.0) + v
    inc["tender"] = tender
    invs = [i["inv"] for i in ocs]
    inc["purchases"] = {"invoices": sorted(invs, key=lambda x: x["date"]), "subtotal": sum(x["subtotal"] for x in invs),
                        "freight": sum(x["freight"] for x in invs), "tax": sum(x["tax"] for x in invs),
                        "total": sum(x["total"] for x in invs)}
    inc["inventory"] = inv
    inc["daily"] = sorted(d for i in gl for d in i["g"]["daily"])
    inc["fuel"] = {"amount": fuel, "litres": lit}
    inc["moneris"] = [{"cat": i["cat"], **i["m"]} for i in mon]

    # ---- cross-checks between the files
    ck = inc["checks"]
    if gl and inv["sold"] > 0:
        ck.append(_chk("Cost of goods: Greenline financials vs inventory 'sold' (at cost)", sum(i["g"]["cost"] for i in gl), inv["sold"],
                       2.0, 100.0, "Greenline cost", "Inventory sold"))
    if tender:
        low = {k.casefold(): v for k, v in tender.items()}
        for i in mon:
            key = "moneris" if i["cat"] == "mon1" else "online"
            if key in low:
                ck.append(_chk(("In-store" if i["cat"] == "mon1" else "Online") + " card sales: Moneris statement vs Greenline payments",
                               i["m"]["net"], low[key], 1.0, 50.0, "Moneris net sales", "Greenline '" + key + "' payments"))
        if gl:
            ck.append(_chk("Total tendered vs Greenline total (sales + HST)", sum(tender.values()), sum(i["g"]["total"] for i in gl),
                           1.0, 50.0, "Payment types", "Greenline total"))
    if ocs and inv["purchased"] > 0:
        c = _chk("Purchases: OCS invoices vs inventory 'purchased'", inc["purchases"]["subtotal"], inv["purchased"], 2.0, 100.0,
                 "OCS invoices (before freight)", "Inventory purchased")
        dates = sorted(x["date"] for x in invs if x["date"])
        if not c["ok"] and dates:
            last = dates[-1]
            late = sum(x["subtotal"] for x in invs if x["date"] == last)
            c["note"] = (f"Invoices dated {last} total ${late:,.2f}. Deliveries invoiced on the last day of the month are often "
                         f"received into inventory the next month.")
        ck.append(c)
    return inc


def financials(payload: dict) -> dict:
    inc = build_income(payload)
    rev = inc["revenue"]
    return {
        "sales": rev["net"], "cogs": inc["cogs"]["amount"], "gross": inc["gross_profit"], "expenses": inc["total_expenses"],
        "noi": inc["noi"], "gross_margin": inc["gross_margin"], "op_margin": inc["op_margin"],
        "cogs_estimated": inc["cogs"]["estimated"], "sales_src": rev["src"], "cogs_src": inc["cogs"]["src"], "inc": inc,
    }


def _income_sheets(wb, store, period, inc, used):
    """Income Statement (+ Daily Sales, OCS Purchases) sheets."""
    rev, cg = inc["revenue"], inc["cogs"]
    net = rev["net"] or 0.0
    ws = wb.create_sheet(_sheet_name("Income Statement", used), 1)
    ws["A1"] = f"Income Statement - {store} - {period}"
    ws["A1"].font = Font(name=_FONT, bold=True, size=14, color="065F46")
    ws["A2"] = "All figures in CAD. HST is shown for information and is not income or expense."
    ws["A2"].font = Font(name=_FONT, size=9, italic=True, color="71717A")
    _header(ws, 4, ["", "Amount", "% of net sales", "Source / note"])
    pct = lambda v: (v / net) if net else None
    row = [5]

    def line(label, amount, note="", bold=False, fill=None, neg=False, pct_on=True):
        r_ = row[0]
        _put(ws, r_, 1, label, bold=bold, fill=fill)
        _put(ws, r_, 2, -amount if neg else amount, bold=bold, fill=fill)
        c = _put(ws, r_, 3, pct(amount) if pct_on else None, bold=bold, fill=fill)
        ws.cell(row=r_, column=3).number_format = "0.0%"
        _put(ws, r_, 4, note, fill=fill)
        row[0] += 1

    def head(t):
        _put(ws, row[0], 1, t, bold=True, fill=_TOT_FILL)
        for c in (2, 3, 4):
            _put(ws, row[0], c, "", fill=_TOT_FILL)
        row[0] += 1

    head("REVENUE")
    if rev["gross"]:
        line("Gross sales", rev["gross"], rev["src"])
        line("Less: refunds", rev["refunds"], "", neg=True)
        line("Less: discounts", rev["discounts"], "", neg=True)
    line("Net sales", rev["net"], rev["src"], bold=True)
    if rev["hst"]:
        line("HST collected (not income)", rev["hst"], "Owed to the government", pct_on=False)
    head("COST OF GOODS SOLD")
    line("Cost of goods sold", cg["amount"], cg["src"], neg=True)
    line("GROSS PROFIT", inc["gross_profit"], "", bold=True, fill=_TOT_FILL)
    head("OPERATING EXPENSES")
    for x in inc["expense_lines"]:
        line(x["label"], x["amount"], x["src"], neg=True)
    line("Total operating expenses", inc["total_expenses"], "", bold=True, neg=True)
    line("NET OPERATING INCOME", inc["noi"], "", bold=True, fill=_TOT_FILL)
    row[0] += 1

    def table(title, heads, rows, bad=None):
        _put(ws, row[0], 1, title, bold=True)
        row[0] += 1
        for j, h in enumerate(heads, 1):
            _put(ws, row[0], j, h, bold=True, fill=_TOT_FILL)
        row[0] += 1
        for rec in rows:
            for j, v in enumerate(rec, 1):
                _put(ws, row[0], j, v, fill=_BAD_FILL if (bad and bad(rec)) else None)
            row[0] += 1
        row[0] += 1

    if inc["tender"]:
        tot = sum(inc["tender"].values()) or 1
        table("SALES BY PAYMENT TYPE (incl. HST)", ["Payment type", "Amount", "Share"],
              [[k, v, v / tot] for k, v in inc["tender"].items()] + [["Total", tot, 1.0]])
        for r_ in range(row[0] - 1 - len(inc["tender"]) - 1, row[0] - 1):
            ws.cell(row=r_, column=3).number_format = "0.0%"
    if inc["moneris"]:
        table("MONERIS CARD STATEMENTS", ["Statement", "Net sales", "Fees", "Fee % of sales", "Deposited"],
              [["In-store" if m["cat"] == "mon1" else "Online", m["net"], m["fees"], (m["fees"] / m["net"]) if m["net"] else 0, m["deposits"]]
               for m in inc["moneris"]])
        for r_ in range(row[0] - 1 - len(inc["moneris"]), row[0] - 1):
            ws.cell(row=r_, column=4).number_format = "0.00%"
    p = inc["purchases"]
    if p["invoices"]:
        table("PURCHASES FROM OCS", ["Item", "Amount"], [
            ["Invoices", len(p["invoices"])], ["Product", p["subtotal"]], ["Freight", p["freight"]],
            ["HST (recoverable)", p["tax"]], ["Invoice total", p["total"]]])
    iv = inc["inventory"]
    if iv["opening"] or iv["closing"] or iv["sold"]:
        table("INVENTORY MOVEMENT (at cost)", ["Measure", "Amount"], [
            ["Opening inventory", iv["opening"]], ["Purchased", iv["purchased"]], ["Sold", iv["sold"]], ["Closing inventory", iv["closing"]]])
    if inc["checks"]:
        table("CROSS-CHECKS BETWEEN YOUR FILES", ["Check", "First figure", "Second figure", "Difference", "Result", "Note"],
              [[c["label"], c["a"], c["b"], c["diff"], "OK" if c["ok"] else "REVIEW", c.get("note", "")] for c in inc["checks"]],
              bad=lambda rec: rec[4] == "REVIEW")
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 44
    ws.column_dimensions["E"].width = 14
    ws.freeze_panes = "A5"

    if inc["daily"]:
        wd = wb.create_sheet(_sheet_name("Daily Sales", used))
        _header(wd, 1, ["Date", "Net sales"])
        for i, (d, v) in enumerate(inc["daily"], start=2):
            _put(wd, i, 1, d)
            _put(wd, i, 2, v)
        n = len(inc["daily"]) + 2
        _put(wd, n, 1, "TOTAL", bold=True, fill=_TOT_FILL)
        _put(wd, n, 2, sum(v for _, v in inc["daily"]), bold=True, fill=_TOT_FILL)
        _put(wd, n + 1, 1, "Daily average", bold=True)
        _put(wd, n + 1, 2, sum(v for _, v in inc["daily"]) / len(inc["daily"]), bold=True)
        _autowidth(wd)
    if p["invoices"]:
        wo = wb.create_sheet(_sheet_name("OCS Purchases", used))
        _header(wo, 1, ["Invoice", "Date", "Product", "Freight", "HST", "Invoice total"])
        for i, x in enumerate(p["invoices"], start=2):
            for j, v in enumerate([x["number"], x["date"], x["subtotal"], x["freight"], x["tax"], x["total"]], 1):
                _put(wo, i, j, v)
        n = len(p["invoices"]) + 2
        _put(wo, n, 1, "TOTAL", bold=True, fill=_TOT_FILL)
        _put(wo, n, 2, "", fill=_TOT_FILL)
        for j, k in enumerate(["subtotal", "freight", "tax", "total"], 3):
            _put(wo, n, j, p[k], bold=True, fill=_TOT_FILL)
        _autowidth(wo)


# ==========================================================================
# OWNER REPORT: one readable page with charts (opens in any browser, prints to PDF)
# ==========================================================================
import html as _html


def _e(x) -> str:
    return _html.escape(str(x if x is not None else ""), quote=True)


def _m(x: float, cents: bool = False) -> str:
    x = float(x or 0)
    s = f"{abs(x):,.2f}" if cents else f"{abs(x):,.0f}"
    return ("-$" if x < 0 else "$") + s


def _pct(x: float) -> str:
    return f"{x:.1f}%"


def _category_sales(payload: dict) -> list[tuple[str, float]]:
    rows = payload.get("inventory_data") or []
    if not rows:
        return []
    headers = list(rows[0].keys())
    sold_h = next((h for h in headers if h and "sold" in str(h).lower()), None)
    if not sold_h:
        return []
    out = [(str(r.get("Category", "")).strip(), _num(r.get(sold_h))) for r in rows]
    return sorted([(c, v) for c, v in out if c and v > 0], key=lambda t: -t[1])


_DONUT_COLORS = ["#065f46", "#10b981", "#34d399", "#f59e0b", "#6366f1", "#ec4899", "#94a3b8"]


def _donut(parts: list[tuple[str, float]]) -> str:
    total = sum(v for _, v in parts) or 1
    circles, offset = [], 0.0
    for i, (_, v) in enumerate(parts):
        p = v / total * 100
        circles.append(
            f'<circle cx="21" cy="21" r="15.9155" fill="none" stroke="{_DONUT_COLORS[i % len(_DONUT_COLORS)]}" '
            f'stroke-width="6" stroke-dasharray="{p:.3f} {100 - p:.3f}" stroke-dashoffset="{-offset:.3f}" '
            f'transform="rotate(-90 21 21)"/>'
        )
        offset += p
    return ('<svg viewBox="0 0 42 42" width="170" height="170" role="img" aria-label="Expense breakdown">'
            '<circle cx="21" cy="21" r="15.9155" fill="none" stroke="#e5e7eb" stroke-width="6"/>'
            + "".join(circles) + "</svg>")


def build_owner_report_html(store: str, period: str, payload: dict) -> str:
    fin = financials(payload)
    inc = fin["inc"]
    sales, cogs, gross, exp, noi = fin["sales"], fin["cogs"], fin["gross"], fin["expenses"], fin["noi"]
    infos = [i for i in (payload.get("docs_info") or []) if i.get("kind") != "attached" or i.get("status") != "attached"]
    exp_items = sorted([(x["label"], x["amount"]) for x in inc["expense_lines"]], key=lambda t: -t[1])
    exp_items = [t for t in exp_items if t[1] > 0]
    cats = _category_sales(payload)
    missing = missing_docs(payload)
    done = len(CHECKLIST_ITEMS) - len(missing)
    cooh = str(payload.get("cooh_text", "")).strip()
    chargebacks = payload.get("chargebacks_text") or "Nothing to report."

    # ---- plain-language summary
    if sales > 0:
        verdict = "made a profit of" if noi >= 0 else "ran at a loss of"
        story = (f"{store} sold <b>{_m(sales)}</b> in {_e(period)}. The products it sold cost about <b>{_m(cogs)}</b>, "
                 f"leaving a gross profit of <b>{_m(gross)}</b> ({_pct(fin['gross_margin'])}). After <b>{_m(exp)}</b> of "
                 f"operating expenses the store {verdict} <b>{_m(abs(noi))}</b> ({_pct(abs(fin['op_margin']))} of sales).")
    else:
        story = ("No sales figures have been uploaded yet for this month. Upload the Greenline financials or the "
                 "inventory report and this page fills in automatically.")
    if exp_items and exp > 0:
        top = exp_items[0]
        story += f" The biggest expense was <b>{_e(top[0])}</b> at {_m(top[1])} ({top[1] / exp * 100:.0f}% of expenses)."
    notes = []
    if fin["cogs_estimated"]:
        notes.append("Cost of goods sold is an estimate (80% of sales) because no product cost was found in the uploads.")
    if missing:
        notes.append(f"{len(missing)} document(s) still missing: " + ", ".join(missing) + ".")

    # ---- where every sales dollar went (stacked bar)
    bar = ""
    if sales > 0:
        c_pct = min(cogs / sales * 100, 100)
        e_pct = min(exp / sales * 100, 100 - c_pct)
        p_pct = max(0.0, 100 - c_pct - e_pct)
        bar = f"""
        <div class="stack" role="img" aria-label="Split of each sales dollar">
          <span style="width:{c_pct:.2f}%;background:#94a3b8">{'Product cost ' + _pct(c_pct) if c_pct > 14 else ''}</span>
          <span style="width:{e_pct:.2f}%;background:#f59e0b">{'Expenses ' + _pct(e_pct) if e_pct > 12 else ''}</span>
          <span style="width:{p_pct:.2f}%;background:#059669">{'Profit ' + _pct(p_pct) if p_pct > 10 else ''}</span>
        </div>
        <div class="legend"><i style="background:#94a3b8"></i>Product cost {_m(cogs)}
          <i style="background:#f59e0b"></i>Operating expenses {_m(exp)}
          <i style="background:#059669"></i>{'Profit ' + _m(noi) if noi >= 0 else 'No profit left (loss ' + _m(abs(noi)) + ')'}</div>"""

    # ---- expense donut
    donut = ""
    if exp_items:
        parts = exp_items[:6]
        rest = sum(v for _, v in exp_items[6:])
        if rest > 0:
            parts.append(("Other", rest))
        legend = "".join(
            f'<li><i style="background:{_DONUT_COLORS[i % len(_DONUT_COLORS)]}"></i>{_e(n)}<b>{_m(v)}</b></li>'
            for i, (n, v) in enumerate(parts))
        donut = f'<div class="donut">{_donut(parts)}<ul>{legend}</ul></div>'
    else:
        donut = '<p class="muted">No operating expenses entered yet.</p>'

    # ---- sales by category bars
    cat_html = ""
    if cats:
        mx = cats[0][1]
        cat_html = "".join(
            f'<div class="hb"><span class="hl">{_e(n)}</span><span class="ht"><span style="width:{v / mx * 100:.1f}%"></span></span>'
            f'<b>{_m(v)}</b></div>' for n, v in cats[:10])

    # ---- documents read
    def _fmt(l):
        v = l[1]
        if isinstance(v, str):
            return _e(l[0]) + ": " + _e(v)
        return _e(l[0]) + ": " + (f"{v:,.1f}" if l[0] in ("Days", "Litres") else _m(v, True))
    seen_ocs = [i for i in infos if i["kind"] == "ocs_invoice"]
    doc_rows = ""
    for i in infos:
        if i["kind"] == "ocs_invoice":
            continue
        doc_rows += (f"<tr><td>{_e(DOC_NAME.get(i['cat'], i['cat']))}</td><td>{_e(i['file'])}</td>"
                     f"<td>{' &middot; '.join(_fmt(l) for l in i['lines'][:5])}</td></tr>")
    if seen_ocs:
        p = inc["purchases"]
        doc_rows += (f"<tr><td>OCS: Invoices</td><td>{len(seen_ocs)} invoice files</td><td>Product + freight: "
                     f"{_m(p['subtotal'] + p['freight'], True)} &middot; HST: {_m(p['tax'], True)} &middot; Invoice total: {_m(p['total'], True)}</td></tr>")
    facts_html = (f"<table><thead><tr><th>Document</th><th>File</th><th>What was read</th></tr></thead><tbody>{doc_rows}</tbody></table>"
                  if doc_rows else '<p class="muted">No documents uploaded yet.</p>')

    checklist = "".join(
        f"<li class='{'ok' if label not in missing else 'no'}'>{'✓' if label not in missing else '✗'} {_e(label)}</li>"
        for _, label in CHECKLIST_ITEMS)

    # ---- income statement table
    rev, cgs = inc["revenue"], inc["cogs"]
    pc = lambda v: (f"{v / rev['net'] * 100:.1f}%" if rev["net"] else "")
    def irow(label, v, bold=False, neg=False, note=""):
        cls = ' class="strong"' if bold else ""
        shown = _m(-v if neg else v, True)
        return f"<tr{cls}><td>{_e(label)}</td><td class='r'>{shown}</td><td class='r'>{pc(v)}</td><td class='muted'>{_e(note)}</td></tr>"
    ir = ""
    if rev["gross"]:
        ir += irow("Gross sales", rev["gross"]) + irow("Less: refunds", rev["refunds"], neg=True) + irow("Less: discounts", rev["discounts"], neg=True)
    ir += irow("Net sales", rev["net"], True, note=rev["src"])
    ir += irow("Cost of goods sold", cgs["amount"], neg=True, note=cgs["src"])
    ir += irow("Gross profit", gross, True)
    for x in inc["expense_lines"]:
        ir += irow(x["label"], x["amount"], neg=True, note=x["src"] if x["src"] != "Entered by you" else "")
    ir += irow("Total operating expenses", exp, True, neg=True)
    ir += irow("Net operating income", noi, True)
    hst_note = (f'<p class="muted">HST collected from customers: {_m(rev["hst"], True)} (owed to the government, not income).</p>'
                if rev["hst"] else "")
    income_html = (f'<table class="inc"><thead><tr><th></th><th class="r">Amount</th><th class="r">% of sales</th><th></th></tr></thead>'
                   f'<tbody>{ir}</tbody></table>{hst_note}')

    # ---- daily sales chart
    daily_html = ""
    dl = inc["daily"]
    if len(dl) >= 2:
        mx = max(v for _, v in dl) or 1
        avg = sum(v for _, v in dl) / len(dl)
        w, h, pad = 600, 150, 16
        bw = (w - 2 * pad) / len(dl)
        bars = "".join(
            f'<rect x="{pad + i * bw + 1:.1f}" y="{h - 22 - (v / mx) * (h - 40):.1f}" width="{max(bw - 2, 1):.1f}" '
            f'height="{(v / mx) * (h - 40):.1f}" rx="2" fill="#10b981"><title>{_e(d)}: {_m(v)}</title></rect>' for i, (d, v) in enumerate(dl))
        ay = h - 22 - (avg / mx) * (h - 40)
        best = max(dl, key=lambda t: t[1])
        worst = min(dl, key=lambda t: t[1])
        daily_html = (f'<svg viewBox="0 0 {w} {h}" width="100%" role="img" aria-label="Daily net sales">{bars}'
                      f'<line x1="{pad}" x2="{w - pad}" y1="{ay:.1f}" y2="{ay:.1f}" stroke="#065f46" stroke-dasharray="4 3"/>'
                      f'<text x="{pad}" y="{h - 6}" font-size="10" fill="#71717a">{_e(dl[0][0][5:])}</text>'
                      f'<text x="{w - pad}" y="{h - 6}" font-size="10" fill="#71717a" text-anchor="end">{_e(dl[-1][0][5:])}</text></svg>'
                      f'<p class="muted">Average day {_m(avg)} &middot; best day {_e(best[0][5:])} {_m(best[1])} &middot; slowest day {_e(worst[0][5:])} {_m(worst[1])}</p>')

    # ---- payment mix donut
    tender_html = ""
    if inc["tender"]:
        tparts = sorted(inc["tender"].items(), key=lambda t: -t[1])
        tl = "".join(f'<li><i style="background:{_DONUT_COLORS[i % len(_DONUT_COLORS)]}"></i>{_e(n)}<b>{_m(v)}</b></li>' for i, (n, v) in enumerate(tparts))
        tender_html = f'<div class="donut">{_donut(tparts)}<ul>{tl}</ul></div><p class="muted">Includes HST. As recorded by the till.</p>'

    # ---- purchases + cross-checks
    p_ = inc["purchases"]
    purchases_html = ""
    if p_["invoices"]:
        purchases_html = (f'<table><tbody><tr><td>OCS invoices</td><td class="r"><b>{len(p_["invoices"])}</b></td></tr>'
                          f'<tr><td>Product</td><td class="r">{_m(p_["subtotal"], True)}</td></tr><tr><td>Freight</td><td class="r">{_m(p_["freight"], True)}</td></tr>'
                          f'<tr><td>HST (recoverable)</td><td class="r">{_m(p_["tax"], True)}</td></tr><tr><td><b>Invoice total</b></td><td class="r"><b>{_m(p_["total"], True)}</b></td></tr></tbody></table>')
    checks_html = "".join(
        f'<li class="{"ok" if c["ok"] else "no"}">{"✓" if c["ok"] else "!"} {_e(c["label"])}<br><span class="muted">{_e(c["a_name"])} {_m(c["a"], True)} vs {_e(c["b_name"])} {_m(c["b"], True)}'
        f'{"" if c["ok"] else " &middot; difference " + _m(c["diff"], True)}</span>'
        f'{"<br><span class=muted>" + _e(c["note"]) + "</span>" if c.get("note") else ""}</li>' for c in inc["checks"])
    checks_html = f'<ul class="chk1">{checks_html}</ul>' if checks_html else ""

    def tile(label, value, sub="", tone=""):
        return f'<div class="tile {tone}"><span>{_e(label)}</span><b>{value}</b><em>{sub}</em></div>'

    tiles = "".join([
        tile("Total sales", _m(sales), _e(fin["sales_src"]) or "&nbsp;"),
        tile("Product cost", _m(cogs), _e(fin["cogs_src"]) or "&nbsp;"),
        tile("Gross profit", _m(gross), _pct(fin["gross_margin"]) + " margin"),
        tile("Operating expenses", _m(exp), f"{len(exp_items)} expense lines"),
        tile("Net operating income", _m(noi), _pct(fin["op_margin"]) + " of sales", "good" if noi >= 0 else "bad"),
    ])

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_e(store)} - {_e(period)} - Owner Report</title>
<style>
:root{{--ink:#18181b;--muted:#71717a;--line:#e4e4e7;--bg:#f4f4f5;--card:#fff;--em:#065f46;--em2:#047857;--soft:#ecfdf5}}
*{{box-sizing:border-box}} body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
.page{{max-width:900px;margin:0 auto;padding:20px 16px 40px}}
header{{background:var(--em);color:#fff;border-radius:16px;padding:22px 24px;display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center}}
header h1{{margin:0;font-size:24px}} header p{{margin:2px 0 0;color:#a7f3d0;font-size:13px}} header .big{{font-size:20px;font-weight:700;text-align:right}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-top:14px;break-inside:avoid}}
h2{{margin:0 0 10px;font-size:16px;color:var(--em)}}
.story{{font-size:16px;line-height:1.6}}
.tiles{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-top:14px}}
.tile{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px;display:flex;flex-direction:column}}
.tile span{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}} .tile b{{font-size:22px;margin:2px 0}} .tile em{{font-style:normal;font-size:12px;color:var(--muted)}}
.tile.good{{background:var(--soft);border-color:#6ee7b7}} .tile.good b{{color:var(--em2)}} .tile.bad{{background:#fef2f2;border-color:#fca5a5}} .tile.bad b{{color:#b91c1c}}
.stack{{display:flex;height:38px;border-radius:10px;overflow:hidden;background:#e5e7eb}} .stack span{{color:#fff;font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;white-space:nowrap}}
.legend{{margin-top:8px;font-size:13px;color:var(--muted)}} .legend i,.donut i{{display:inline-block;width:10px;height:10px;border-radius:3px;margin:0 6px 0 12px}} .legend i:first-child{{margin-left:0}}
.donut{{display:flex;align-items:center;gap:22px;flex-wrap:wrap}} .donut ul{{list-style:none;margin:0;padding:0;flex:1;min-width:220px}}
.donut li{{display:flex;align-items:center;padding:3px 0;border-bottom:1px solid var(--line);font-size:14px}} .donut li i{{margin-left:0}} .donut li b{{margin-left:auto}}
.hb{{display:grid;grid-template-columns:150px 1fr 90px;gap:8px;align-items:center;margin:5px 0;font-size:13px}} .hb .hl{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.ht{{background:#e5e7eb;border-radius:6px;height:14px;overflow:hidden}} .ht span{{display:block;height:100%;background:linear-gradient(90deg,#10b981,#065f46)}} .hb b{{text-align:right}}
table{{width:100%;border-collapse:collapse;font-size:13px}} th,td{{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left}} th{{color:var(--muted);font-weight:600}} .r{{text-align:right}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:14px}} @media(max-width:640px){{.two{{grid-template-columns:1fr}} .hb{{grid-template-columns:100px 1fr 80px}}}}
ul.chk{{list-style:none;margin:0;padding:0;columns:2;font-size:13px}} ul.chk li.ok{{color:var(--em2)}} ul.chk li.no{{color:#b91c1c;font-weight:600}}
.muted{{color:var(--muted);font-size:13px}} .note{{background:#fffbeb;border:1px solid #fcd34d;border-radius:10px;padding:8px 12px;margin-top:10px;font-size:13px}}
table.inc td,table.inc th{{padding:5px 8px}} table.inc tr.strong td{{font-weight:700;background:#ecfdf5}} ul.chk1{{list-style:none;margin:0;padding:0;font-size:13px}} ul.chk1 li{{padding:5px 0;border-bottom:1px solid var(--line)}} ul.chk1 li.ok{{color:var(--em2)}} ul.chk1 li.no{{color:#b45309}}
footer{{margin-top:18px;text-align:center;color:var(--muted);font-size:12px}}
@media print{{body{{background:#fff}} .page{{padding:0}} .card,.tile{{box-shadow:none}} @page{{size:A4;margin:12mm}}}}
</style></head><body><div class="page">
<header><div><h1>🌿 {_e(store)}</h1><p>Monthly owner report · {_e(period)}</p></div>
<div class="big">{'Net operating income<br>' + _m(noi)}</div></header>

<div class="card"><h2>The month in plain English</h2><p class="story">{story}</p>{''.join(f'<div class="note">{_e(n)}</div>' for n in notes)}</div>

<div class="tiles">{tiles}</div>

{f'<div class="card"><h2>Where every sales dollar went</h2>{bar}</div>' if bar else ''}

<div class="card"><h2>Income statement</h2>{income_html}</div>

{f'<div class="card"><h2>Net sales by day</h2>{daily_html}</div>' if daily_html else ''}

<div class="two">
  <div class="card"><h2>Operating expenses</h2>{donut}</div>
  <div class="card"><h2>Cash and chargebacks</h2>
    <table><tbody>
    <tr><td>Cash on hand</td><td class="r"><b>{_m(parse_money(cooh), True) if cooh else 'Not entered'}</b></td></tr>
    <tr><td>Cash counted on</td><td class="r">{_e(payload.get('cooh_date') or 'Not entered')}</td></tr>
    <tr><td>Moneris chargebacks</td><td class="r">{_e(chargebacks)}</td></tr>
    </tbody></table>
    <h2 style="margin-top:14px">Paperwork ({done} of {len(CHECKLIST_ITEMS)} received)</h2>
    <ul class="chk">{checklist}</ul>
  </div>
</div>

{f'<div class="card"><h2>How customers paid</h2>{tender_html}</div>' if tender_html else ''}

{f'<div class="two"><div class="card"><h2>Purchases from OCS</h2>{purchases_html}</div><div class="card"><h2>Do the files agree?</h2>{checks_html}</div></div>' if purchases_html or checks_html else ''}

{f'<div class="card"><h2>Cost of goods by product category</h2>{cat_html}</div>' if cat_html else ''}

<div class="card"><h2>What was read from your uploaded documents</h2>{facts_html}</div>

<footer>Prepared {datetime.now():%B %d, %Y} · The OZ Store · Figures come from the files uploaded for this month. To save as PDF, use your browser's Print option.</footer>
</div></body></html>"""


# ==========================================================================
# THE PAGE (your original Tailwind / emerald interface, unchanged in look)
# ==========================================================================
HTML_CODE = r"""<!DOCTYPE html>
<html lang="en">
<head><style>html,body{background:#f4f4f5 !important;margin:0}html{overflow-y:scroll !important;overflow-x:hidden !important;scrollbar-width:auto;scrollbar-color:#059669 #d4d4d8}::-webkit-scrollbar{width:22px}::-webkit-scrollbar-track{background:#d4d4d8}::-webkit-scrollbar-thumb{background:#059669;border-radius:11px;border:4px solid #d4d4d8;min-height:70px}::-webkit-scrollbar-thumb:hover{background:#047857}</style>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>The OZ Store - Executive Financial Report & Package Manager</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
</head>
<body class="bg-zinc-100 text-zinc-900 flex flex-col items-center py-8 px-4 font-sans">

    <div class="max-w-7xl w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-zinc-200">

        <header class="bg-emerald-800 text-white p-6 md:p-8 flex flex-col md:flex-row justify-between items-center gap-4 border-b border-emerald-900">
            <div class="flex items-center gap-4">
                <div class="bg-white p-2.5 rounded-xl shadow-md flex items-center justify-center">
                    <span class="text-2xl">🌿</span>
                </div>
                <div>
                    <h1 class="text-xl md:text-2xl font-extrabold text-white tracking-tight">The OZ Store</h1>
                    <p class="text-emerald-100/80 text-xs mt-0.5">Executive Financial & Accountant Document Package Manager</p>
                </div>
            </div>

            <div class="flex items-center gap-4">
                <div class="flex items-center gap-3 bg-white text-zinc-900 px-4 py-2 rounded-xl shadow-md border border-emerald-700">
                    <span class="text-xs font-bold text-emerald-800 uppercase tracking-wider">Period:</span>
                    <select id="select-month" onchange="updatePeriod()" class="bg-zinc-50 text-zinc-900 text-xs font-semibold px-2.5 py-1.5 rounded-lg border border-zinc-300 focus:outline-none focus:border-emerald-600">
                        <option value="January">January</option>
                        <option value="February">February</option>
                        <option value="March">March</option>
                        <option value="April">April</option>
                        <option value="May">May</option>
                        <option value="June">June</option>
                        <option value="July">July</option>
                        <option value="August">August</option>
                        <option value="September">September</option>
                        <option value="October" selected>October</option>
                        <option value="November">November</option>
                        <option value="December">December</option>
                    </select>
                    <select id="select-year" onchange="updatePeriod()" class="bg-zinc-50 text-zinc-900 text-xs font-semibold px-2.5 py-1.5 rounded-lg border border-zinc-300 focus:outline-none focus:border-emerald-600">
                        <option value="2025">2025</option>
                        <option value="2026" selected>2026</option>
                        <option value="2027">2027</option>
                        <option value="2028">2028</option>
                    </select>
                </div>
                <button onclick="navigateParent('?auth_user=' + encodeURIComponent(currentUsername) + '&action=sign_out')" class="bg-emerald-900 hover:bg-emerald-950 text-emerald-100 px-3.5 py-2 rounded-xl text-xs font-semibold border border-emerald-700 transition">
                    Sign Out (REPLACE_CURRENT_NAME)
                </button>
            </div>
        </header>

        <div class="p-6 md:p-8">

            <div class="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 pb-4 border-b border-zinc-200">
                <div class="flex items-center gap-2 overflow-x-auto overflow-y-hidden pb-2 scrollbar-thin w-full md:w-auto" id="store-tabs"></div>
                <div class="flex items-center gap-3">
                    REPLACE_ADD_USER_BTN_HTML
                    <div class="text-xs text-zinc-500 font-medium bg-zinc-50 px-3 py-1.5 rounded-lg border border-zinc-200">
                        👤 Logged in as: <span class="font-bold text-emerald-800">REPLACE_CURRENT_NAME</span>
                    </div>
                </div>
            </div>

            <!-- COLLAPSIBLE CREDENTIAL MANAGER (ADMIN ONLY) -->
            <details id="credential-manager-view" class="REPLACE_ADMIN_HIDDEN_CLASS mb-8 bg-zinc-50 rounded-2xl border border-zinc-300 shadow-xs overflow-hidden">
                <summary class="p-4 bg-emerald-50 hover:bg-emerald-100 cursor-pointer font-bold text-emerald-900 text-xs flex items-center justify-between transition">
                    <span class="flex items-center gap-2"><span>🔐</span> Employee & Store Login Credentials Directory (Click to Edit)</span>
                    <span class="text-[11px] text-emerald-700 underline font-normal">Expand / Collapse</span>
                </summary>
                <div class="p-5 border-t border-zinc-200 bg-white space-y-4">
                    <div class="bg-white rounded-xl border border-zinc-200 overflow-hidden shadow-xs">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead class="bg-emerald-50 text-emerald-900 border-b border-emerald-200">
                                <tr>
                                    <th class="p-3 font-semibold">Username</th>
                                    <th class="p-3 font-semibold">Display Name</th>
                                    <th class="p-3 font-semibold">Role</th>
                                    <th class="p-3 font-semibold">Password</th>
                                    <th class="p-3 font-semibold text-right">Actions</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-zinc-100 text-zinc-700">
                                REPLACE_CRED_TABLE_ROWS
                            </tbody>
                        </table>
                    </div>

                    <div id="user-edit-form-box" class="hidden bg-emerald-50/60 p-4 rounded-xl border border-emerald-300 mt-4">
                        <h4 class="font-bold text-emerald-900 text-xs mb-3" id="edit-form-title">Edit User Credential</h4>
                        <form onsubmit="submitUserEdit(event)" class="grid grid-cols-1 md:grid-cols-5 gap-3 text-xs">
                            <input type="hidden" id="edit-orig-username">
                            <div>
                                <label class="block font-semibold mb-1 text-zinc-700">Username</label>
                                <input type="text" id="edit-username" required class="w-full border border-zinc-300 rounded px-2 py-1.5 bg-white">
                            </div>
                            <div>
                                <label class="block font-semibold mb-1 text-zinc-700">Display Name</label>
                                <input type="text" id="edit-name" required class="w-full border border-zinc-300 rounded px-2 py-1.5 bg-white">
                            </div>
                            <div>
                                <label class="block font-semibold mb-1 text-zinc-700">Role</label>
                                <select id="edit-role" class="w-full border border-zinc-300 rounded px-2 py-1.5 bg-white">
                                    <option value="store">Store</option>
                                    <option value="accountant">Accountant</option>
                                    <option value="admin">Admin</option>
                                </select>
                            </div>
                            <div>
                                <label class="block font-semibold mb-1 text-zinc-700">Password</label>
                                <input type="text" id="edit-password" required class="w-full border border-zinc-300 rounded px-2 py-1.5 bg-white">
                            </div>
                            <div class="flex items-end gap-2">
                                <button type="submit" class="bg-emerald-700 hover:bg-emerald-800 text-white font-bold px-3 py-1.5 rounded transition w-full">Save</button>
                                <button type="button" onclick="document.getElementById('user-edit-form-box').classList.add('hidden')" class="bg-zinc-200 text-zinc-700 px-3 py-1.5 rounded transition">Cancel</button>
                            </div>
                        </form>
                    </div>
                </div>
            </details>

            <div class="flex justify-between items-center mb-6 bg-emerald-50/60 p-4 rounded-xl border border-emerald-200 shadow-xs">
                <div>
                    <span class="text-xs uppercase tracking-wider text-emerald-800 font-bold block mb-0.5">Active Workspace</span>
                    <h2 class="text-xl font-extrabold text-zinc-900" id="workspace-store-name">Oz 1</h2>
                </div>
                <div class="flex items-center gap-3 text-xs text-zinc-600">
                    <button type="button" onclick="saveAllData(false)" class="bg-emerald-700 hover:bg-emerald-800 text-white px-4 py-1.5 rounded-full font-semibold shadow-xs border border-emerald-600 transition cursor-pointer">
                        💾 Save All to Cloud
                    </button>
                    <span id="active-period-label" class="bg-white text-emerald-800 border border-emerald-300 px-3 py-1 rounded-full font-semibold shadow-xs">October 2026</span>
                    <span id="store-completion-badge" class="bg-amber-50 text-amber-800 border border-amber-300 px-3 py-1 rounded-full font-semibold shadow-xs">Checklist: 0/10 Completed</span>
                </div>
            </div>

            <!-- SIDE-BY-SIDE SECTIONS -->
            <div class="grid grid-cols-1 xl:grid-cols-12 gap-6 mb-8 REPLACE_ACCOUNTANT_HIDDEN_CLASS">

                <!-- LEFT SIDE: Inventory Report Formatter -->
                <div class="xl:col-span-5 bg-zinc-50 rounded-xl p-5 border border-zinc-200 flex flex-col justify-between shadow-xs">
                    <div>
                        <div class="flex justify-between items-center mb-3">
                            <h3 class="text-base font-bold text-emerald-800 flex items-center gap-2">Inventory Report Formatter</h3>
                            <span id="inv-status-badge" class="text-[11px] text-zinc-600 bg-white border border-zinc-300 px-2 py-0.5 rounded shadow-xs">Not Loaded</span>
                        </div>
                        <p class="text-zinc-600 text-xs mb-4">Upload raw monthly CSV. Rows are sorted and automatically pushed to Greenline Inventory Summary.</p>

                        <div id="drop-zone" class="border-2 border-dashed border-emerald-400 hover:border-emerald-700 rounded-xl p-5 text-center cursor-pointer transition-all bg-white group mb-4 shadow-xs">
                            <input type="file" id="file-input" accept=".csv" class="hidden">
                            <svg class="w-8 h-8 mx-auto text-emerald-700 mb-2 group-hover:scale-110 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                            </svg>
                            <p class="text-xs font-medium text-zinc-700">Drag & drop store CSV here</p>
                            <p class="text-[11px] text-emerald-700 underline mt-0.5">or click to browse</p>
                        </div>
                        <div id="csv-file-info" class="hidden mb-4 bg-emerald-50 border border-emerald-200 rounded-xl p-3 flex items-center justify-between">
                            <div class="flex items-center gap-2">
                                <span class="text-emerald-700">📄</span>
                                <span id="csv-file-name" class="text-xs font-medium text-emerald-800"></span>
                            </div>
                            <button type="button" onclick="clearCSVUpload()" class="bg-red-100 hover:bg-red-200 text-red-700 text-xs font-bold px-3 py-1.5 rounded-lg border border-red-300 transition">
                                🗑️ Delete & Re-upload
                            </button>
                        </div>

                        <div class="bg-white rounded-lg border border-zinc-200 overflow-hidden  mb-4 shadow-xs">
                            <table class="w-full text-left border-collapse text-[11px]">
                                <thead class="bg-emerald-50 text-emerald-900 sticky top-0 z-10 border-b border-emerald-200">
                                    <tr id="table-header"><th class="p-2 font-semibold">Category</th><th class="p-2 font-semibold">Opening</th><th class="p-2 font-semibold">Sold</th><th class="p-2 font-semibold">Closing</th></tr>
                                </thead>
                                <tbody id="table-body" class="divide-y divide-zinc-100 text-zinc-700"><tr><td colspan="4" class="p-4 text-center text-zinc-400">No report uploaded yet.</td></tr></tbody>
                            </table>
                        </div>
                    </div>

                    <div id="action-area" class="hidden">
                        <button id="download-excel-btn" class="w-full bg-emerald-700 hover:bg-emerald-800 text-white font-bold py-2.5 px-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 text-xs">
                            Download Formatted Excel (.xlsx)
                        </button>
                    </div>
                </div>

                <!-- RIGHT SIDE: Monthly Accountant Document Checklist -->
                <div class="xl:col-span-7 bg-zinc-50 rounded-xl p-5 border border-zinc-200 flex flex-col justify-between shadow-xs">
                    <div>
                        <div class="flex justify-between items-center mb-3">
                            <h3 class="text-base font-bold text-emerald-800">Monthly Accountant Package Checklist & Folders</h3>
                        </div>
                        <p class="text-zinc-600 text-xs mb-4">Input cash on hand, browse supporting files, and add notes. Click <b>Save All to Cloud</b> when ready.</p>

                        <div class="space-y-3.5  pr-2" id="checklist-container">

                            <!-- 1. Cash on Hand -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2 shadow-xs">
                                <div class="flex justify-between items-center">
                                    <label class="block text-xs font-bold text-emerald-800">Cash On Hand - End of Day</label>
                                    <input type="date" id="cooh-date" onchange="updateCompletionBadge(); saveLocalOnly();" class="bg-white text-zinc-700 text-[11px] px-2 py-1 rounded border border-zinc-300">
                                </div>
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                    <input type="text" id="cooh-text" placeholder="Enter amount e.g. $24,521.10" oninput="syncNotes('cooh', this.value)" class="w-full bg-white border border-zinc-300 rounded-lg px-3 py-2 text-xs text-zinc-800">
                                    <div class="flex flex-col gap-1">
                                        <input type="file" id="cooh-file" multiple onchange="handleItemFiles('cooh', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('cooh')" class="flex-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-medium py-2 px-3 rounded-lg border border-emerald-300 transition truncate">📎 Browse & Upload Document</button>
                                            <button type="button" onclick="clearUploadField('cooh')" class="bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold px-2 py-2 rounded-lg border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                </div>
                                <div id="files-cooh" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files uploaded</div>
                            </div>

                            <!-- 2. Brinks Book -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2 shadow-xs">
                                <div class="flex items-center justify-between">
                                    <label class="flex items-center gap-2 text-xs font-bold text-emerald-800 cursor-pointer">
                                        <input type="checkbox" id="brinks-chk" onchange="updateCompletionBadge(); saveLocalOnly();" class="accent-emerald-600 w-4 h-4 rounded"> Brinks Book (Folder)
                                    </label>
                                    <input type="date" id="brinks-date" onchange="updateCompletionBadge(); saveLocalOnly();" class="bg-white text-zinc-700 text-[11px] px-2 py-1 rounded border border-zinc-300">
                                </div>
                                <div class="pl-5 space-y-2">
                                    <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                        <input type="text" id="brinks-note" placeholder="Add notes..." oninput="syncNotes('brinks', this.value)" class="w-full bg-white border border-zinc-300 rounded-lg px-3 py-2 text-xs text-zinc-800">
                                        <div class="flex flex-col gap-1">
                                            <input type="file" id="brinks-file" multiple onchange="handleItemFiles('brinks', this)" class="hidden">
                                            <div class="flex gap-1">
                                                <button type="button" onclick="openUploadPanel('brinks')" class="flex-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-medium py-2 px-3 rounded-lg border border-emerald-300 transition truncate">📎 Browse & Upload Photos</button>
                                                <button type="button" onclick="clearUploadField('brinks')" class="bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold px-2 py-2 rounded-lg border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                            </div>
                                        </div>
                                    </div>
                                    <div id="files-brinks" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files uploaded</div>
                                </div>
                            </div>

                            <!-- 3. Greenline Reports -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2.5 shadow-xs">
                                <span class="text-xs font-bold text-emerald-800 block">Greenline Reports</span>
                                <div class="pl-3 space-y-2.5 bg-zinc-50 p-2.5 rounded-lg border border-zinc-200">
                                    <div class="flex justify-between items-center">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green1-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Inventory Summary (Auto-pushed from Formatter)</label>
                                        <input type="file" id="green1-file" multiple onchange="handleItemFiles('green1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('green1')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('green1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-green1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="green1-note" placeholder="Notes..." oninput="syncNotes('green1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green2-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Financials</label>
                                        <input type="file" id="green2-file" multiple onchange="handleItemFiles('green2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('green2')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('green2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-green2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="green2-note" placeholder="Notes..." oninput="syncNotes('green2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green3-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Payment Types</label>
                                        <input type="file" id="green3-file" multiple onchange="handleItemFiles('green3', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('green3')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('green3')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-green3" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="green3-note" placeholder="Notes..." oninput="syncNotes('green3', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">
                                </div>
                            </div>

                            <!-- 4. Moneris Reports -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2.5 shadow-xs">
                                <span class="text-xs font-bold text-emerald-800 block">Moneris Reports</span>
                                <div class="pl-3 space-y-2.5 bg-zinc-50 p-2.5 rounded-lg border border-zinc-200">
                                    <div class="flex justify-between items-center">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="mon1-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Counter Sales</label>
                                        <input type="file" id="mon1-file" multiple onchange="handleItemFiles('mon1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('mon1')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('mon1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-mon1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="mon1-note" placeholder="Notes..." oninput="syncNotes('mon1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="mon2-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Online Sales</label>
                                        <input type="file" id="mon2-file" multiple onchange="handleItemFiles('mon2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('mon2')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('mon2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-mon2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="mon2-note" placeholder="Notes..." oninput="syncNotes('mon2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">
                                </div>
                            </div>

                            <!-- 5. OCS Documents -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2.5 shadow-xs">
                                <span class="text-xs font-bold text-emerald-800 block">OCS Documents</span>
                                <div class="pl-3 space-y-2.5 bg-zinc-50 p-2.5 rounded-lg border border-zinc-200">
                                    <div class="flex justify-between items-center">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="ocs1-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Account Statement</label>
                                        <input type="file" id="ocs1-file" multiple onchange="handleItemFiles('ocs1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('ocs1')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('ocs1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-ocs1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="ocs1-note" placeholder="Notes..." oninput="syncNotes('ocs1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="ocs2-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Invoices</label>
                                        <input type="file" id="ocs2-file" multiple onchange="handleItemFiles('ocs2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('ocs2')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('ocs2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-ocs2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="ocs2-note" placeholder="Notes..." oninput="syncNotes('ocs2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">
                                </div>
                            </div>

                            <!-- 6. Moneris Chargebacks -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2 shadow-xs">
                                <span class="text-xs font-bold text-emerald-800 block">Moneris Chargebacks</span>
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                    <input type="text" id="chargebacks-text" value="Nothing to report." oninput="syncNotes('chargebacks', this.value)" class="w-full border border-zinc-300 rounded-lg px-3 py-2 text-xs">
                                    <div class="flex flex-col gap-1">
                                        <input type="file" id="chargebacks-file" multiple onchange="handleItemFiles('chargebacks', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('chargebacks')" class="flex-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 text-xs font-medium py-1.5 px-3 rounded-lg border border-zinc-300 transition truncate">📎 Browse & Upload</button>
                                            <button type="button" onclick="clearUploadField('chargebacks')" class="bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold px-2 py-1.5 rounded-lg border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                </div>
                                <div id="files-chargebacks" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600 bg-zinc-50 p-1.5 rounded border border-zinc-200">No files uploaded</div>
                            </div>

                            <!-- 7. Gas Cards -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2.5 shadow-xs">
                                <span class="text-xs font-bold text-emerald-800 block">Gas Cards</span>
                                <div class="pl-3 space-y-2.5 bg-zinc-50 p-2.5 rounded-lg border border-zinc-200">
                                    <div class="flex justify-between items-center">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="gas1-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Details Report</label>
                                        <input type="file" id="gas1-file" multiple onchange="handleItemFiles('gas1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('gas1')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('gas1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-gas1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="gas1-note" placeholder="Notes..." oninput="syncNotes('gas1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="gas2-chk" onchange="updateCompletionBadge(); saveLocalOnly();"> Summary Report</label>
                                        <input type="file" id="gas2-file" multiple onchange="handleItemFiles('gas2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="openUploadPanel('gas2')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('gas2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-gas2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="gas2-note" placeholder="Notes..." oninput="syncNotes('gas2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">
                                </div>
                            </div>

                            <details class="md:col-span-2 lg:col-span-3 bg-white border border-zinc-200 rounded-xl overflow-hidden" id="receipts-section">
                                <summary class="cursor-pointer select-none px-4 py-3 text-sm font-bold text-emerald-900 bg-emerald-50 hover:bg-emerald-100">🧾 Credit Card Receipts (physical receipts - photos or scans, #1 to #50)</summary>
                                <div class="p-4 space-y-3">
                                    <p class="text-xs text-zinc-500">Press Browse on any slot and pick one or many receipt photos or scans; they fill the next empty numbered slots in order (use the × on a chip to free a slot). They are kept with this store and month and included in the accountant package.</p>
                                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #1</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt1')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt1" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #2</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt2')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt2" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #3</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt3')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt3')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt3" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #4</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt4')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt4')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt4" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #5</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt5')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt5')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt5" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #6</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt6')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt6')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt6" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #7</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt7')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt7')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt7" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #8</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt8')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt8')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt8" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #9</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt9')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt9')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt9" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #10</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt10')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt10')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt10" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #11</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt11')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt11')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt11" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #12</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt12')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt12')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt12" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #13</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt13')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt13')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt13" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #14</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt14')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt14')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt14" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #15</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt15')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt15')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt15" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #16</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt16')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt16')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt16" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #17</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt17')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt17')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt17" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #18</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt18')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt18')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt18" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #19</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt19')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt19')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt19" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #20</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt20')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt20')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt20" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #21</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt21')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt21')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt21" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #22</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt22')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt22')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt22" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #23</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt23')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt23')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt23" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #24</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt24')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt24')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt24" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #25</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt25')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt25')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt25" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #26</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt26')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt26')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt26" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #27</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt27')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt27')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt27" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #28</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt28')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt28')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt28" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #29</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt29')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt29')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt29" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #30</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt30')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt30')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt30" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #31</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt31')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt31')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt31" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #32</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt32')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt32')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt32" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #33</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt33')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt33')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt33" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #34</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt34')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt34')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt34" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #35</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt35')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt35')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt35" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #36</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt36')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt36')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt36" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #37</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt37')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt37')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt37" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #38</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt38')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt38')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt38" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #39</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt39')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt39')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt39" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #40</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt40')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt40')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt40" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #41</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt41')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt41')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt41" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #42</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt42')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt42')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt42" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #43</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt43')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt43')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt43" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #44</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt44')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt44')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt44" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #45</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt45')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt45')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt45" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #46</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt46')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt46')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt46" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #47</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt47')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt47')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt47" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #48</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt48')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt48')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt48" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #49</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt49')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt49')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt49" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                        <div class="border border-zinc-200 rounded-lg p-2.5 bg-zinc-50 space-y-1">
                                            <div class="flex justify-between items-center">
                                                <span class="text-xs font-bold text-zinc-700">Receipt #50</span>
                                                <div class="flex gap-1">
                                                    <button type="button" onclick="openUploadPanel('rcpt50')" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                                    <button type="button" onclick="clearUploadField('rcpt50')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Remove this receipt">🗑️</button>
                                                </div>
                                            </div>
                                            <div id="files-rcpt50" class="flex flex-wrap gap-1.5 text-[11px] text-zinc-600">No file</div>
                                        </div>
                                    </div>
                                </div>
                            </details>

                        </div>
                    </div>
                </div>

            </div>

            <!-- ================= ACCOUNTANT REVIEW PORTAL ================= -->
            <div class="bg-emerald-950 text-white rounded-2xl p-6 border border-emerald-800 shadow-lg">
                <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-4 pb-3 border-b border-emerald-800 gap-3">
                    <div>
                        <h3 class="text-lg font-bold text-emerald-300">Accountant Review Portal & Monthly Record Viewer</h3>
                        <p class="text-emerald-200/80 text-xs mt-0.5">Accountant view for <span id="portal-store-label" class="text-white font-semibold">Oz 1</span> – <span id="portal-period-label" class="text-white font-semibold">October 2026</span></p>
                    </div>
                    <div class="flex items-center gap-2">
                        <button onclick="downloadMasterZip()" class="bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-semibold px-3.5 py-2 rounded-xl border border-emerald-500 transition">
                            📦 Download All as ZIP
                        </button>
                        <button onclick="window.print()" class="bg-emerald-900 hover:bg-emerald-800 text-emerald-100 text-xs font-semibold px-3.5 py-2 rounded-xl border border-emerald-700 transition">
                            Print / Save PDF Summary
                        </button>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                    <div class="bg-emerald-900/60 p-4 rounded-xl border border-emerald-800">
                        <span class="text-xs font-semibold text-emerald-300 block mb-1">Cash On Hand</span>
                        <span id="portal-cooh-val" class="text-base font-bold text-white">Not entered</span>
                    </div>
                    <div class="bg-emerald-900/60 p-4 rounded-xl border border-emerald-800">
                        <span class="text-xs font-semibold text-emerald-300 block mb-1">Inventory Formatted Report</span>
                        <span id="portal-inv-val" class="text-xs font-medium text-emerald-200">No CSV report uploaded</span>
                    </div>
                    <div class="bg-emerald-900/60 p-4 rounded-xl border border-emerald-800">
                        <span class="text-xs font-semibold text-emerald-300 block mb-1">Checklist Status</span>
                        <span id="portal-chk-val" class="text-xs font-bold text-amber-300">0 / 10 Items Checked</span>
                    </div>
                </div>

                <div class="bg-emerald-900/40 rounded-xl border border-emerald-800 overflow-x-auto overflow-y-hidden">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-emerald-900 text-emerald-300 border-b border-emerald-800 sticky top-0">
                            <tr>
                                <th class="p-3 font-semibold">Document / Folder Category</th>
                                <th class="p-3 font-semibold">Status</th>
                                <th class="p-3 font-semibold">Document Date</th>
                                <th class="p-3 font-semibold">Attached File Names</th>
                                <th class="p-3 font-semibold">Notes & Details</th>
                            </tr>
                        </thead>
                        <tbody id="portal-table-body" class="divide-y divide-emerald-800/60 text-emerald-100">
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Cash On Hand - End of Day</td><td class="p-3"><span class="bg-emerald-800 text-emerald-200 px-2 py-0.5 rounded text-[11px]">Entered ✓</span></td><td class="p-3 text-emerald-300/50 italic" id="portal-date-cooh">No date</td><td class="p-3" id="table-files-cooh">No files</td><td class="p-3 text-emerald-100" id="note-cooh">Not entered</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Brinks Book (Photos highlighted)</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-brinks">Pending</span></td><td class="p-3 text-emerald-300/50 italic" id="portal-date-brinks">No date</td><td class="p-3" id="table-files-brinks">No files</td><td class="p-3 text-emerald-100" id="note-brinks">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Greenline: Inventory Summary</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-green1">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-green1">No files</td><td class="p-3 text-emerald-100" id="note-green1">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Greenline: Financials</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-green2">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-green2">No files</td><td class="p-3 text-emerald-100" id="note-green2">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Greenline: Payment Types</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-green3">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-green3">No files</td><td class="p-3 text-emerald-100" id="note-green3">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Moneris: Counter Sales Statement</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-mon1">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-mon1">No files</td><td class="p-3 text-emerald-100" id="note-mon1">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Moneris: Online Sales Statement</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-mon2">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-mon2">No files</td><td class="p-3 text-emerald-100" id="note-mon2">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">OCS: Account Statement</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-ocs1">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-ocs1">No files</td><td class="p-3 text-emerald-100" id="note-ocs1">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">OCS: Invoices (Multiple)</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-ocs2">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-ocs2">No files</td><td class="p-3 text-emerald-100" id="note-ocs2">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Moneris Chargebacks</td><td class="p-3"><span class="bg-emerald-800 text-emerald-200 px-2 py-0.5 rounded text-[11px]">Entered ✓</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-chargebacks">No files</td><td class="p-3 text-emerald-100" id="note-chargebacks">Nothing to report.</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Gas Cards: Details Report</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-gas1">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-gas1">No files</td><td class="p-3 text-emerald-100" id="note-gas1">No notes</td></tr>
                            <tr class="hover:bg-emerald-900/30"><td class="p-3 font-medium text-white">Gas Cards: Summary Report</td><td class="p-3"><span class="bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]" id="status-gas2">Pending</span></td><td class="p-3 text-emerald-300/50 italic">No date</td><td class="p-3" id="table-files-gas2">No files</td><td class="p-3 text-emerald-100" id="note-gas2">No notes</td></tr>
                        </tbody>
                    </table>
                </div>

                <!-- ================= MONTHLY EXPENSE REPORT ================= -->
                <div class="mt-6 bg-emerald-900/40 rounded-xl border border-emerald-800 p-5" id="expense-report-section">
                    <h4 class="text-base font-bold text-emerald-300 mb-1">Monthly Expenses &amp; Report</h4>
                    <p class="text-xs text-emerald-200/70 mb-4">Enter this month's expenses for the selected store, then press Generate Report. Sales, cost of goods, card fees and fuel are read from the documents you uploaded.</p>
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Rent</label>
                            <input type="text" inputmode="decimal" id="exp-rent" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Payroll</label>
                            <input type="text" inputmode="decimal" id="exp-payroll" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Utilities</label>
                            <input type="text" inputmode="decimal" id="exp-utilities" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Insurance</label>
                            <input type="text" inputmode="decimal" id="exp-insurance" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">System Fees</label>
                            <input type="text" inputmode="decimal" id="exp-systemfees" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Repairs</label>
                            <input type="text" inputmode="decimal" id="exp-repairs" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block mb-1">Other</label>
                            <input type="text" inputmode="decimal" id="exp-other" placeholder="0.00" oninput="onExpenseInput()" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1.5 text-sm">
                        </div>
                    </div>
                    <div class="flex flex-wrap gap-2 mb-2">
                        <button type="button" onclick="generateExpenseReport(true)" class="bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-bold px-4 py-2 rounded-xl border border-emerald-400 transition">📊 Generate Report</button>
                        <button type="button" onclick="downloadDetailedReport()" class="bg-white hover:bg-emerald-50 text-emerald-800 text-sm font-bold px-4 py-2 rounded-xl border border-emerald-300 transition">⬇️ Download Detailed Accounting Report</button>
                    </div>
                    <p class="text-[11px] text-emerald-200/60">Your entries are kept on this device as you type and are saved to the cloud when you press Save All to Cloud or Download. The detailed report is an Excel file built from every file uploaded.</p>
                    <div id="expense-report" class="mt-4"></div>
                </div>
            </div>

        </div>
    </div>

    <script>
        const ACCOUNTANT_ORDER = ["Beverages", "Dried Flower", "Edibles", "Extracts Ingested", "Extracts Inhaled", "Infused Pre-Rolls ", "Milled Flower", "Pre-Rolls", "Seeds", "Topicals", "Vapes - Disposables", "Vapes - 510 Threads", "Vapes - Pods", "Accessories", "Apparel / Swag", "Gift Cards", "Other Income", "Raffles", "System Fees", "Uniforms"];

        let isStoreManager = REPLACE_IS_STORE_MANAGER;
        let currentUserRole = "REPLACE_CURRENT_ROLE";
        let currentUserName = "REPLACE_CURRENT_NAME";
        let currentUsername = "REPLACE_CURRENT_USERNAME";

        let stores = (isStoreManager || currentUserRole === 'store') ? [currentUserName] : ["Oz 1", "Oz 2", "Oz 3", "Oz 4", "Oz 5", "Oz 6", "Oz 7"];
        let activeStore = "REPLACE_ACTIVE_STORE";
        if (!stores.includes(activeStore)) activeStore = stores[0];
        let selectedMonth = "REPLACE_MONTH";
        let selectedYear = "REPLACE_YEAR";

        let storedFilesMap = REPLACE_FILES_JSON;
        let persistentStoreData = REPLACE_STORE_DATA_JSON;
        let storeInventoryData = {};
        let docFacts = REPLACE_DOC_FACTS;

        function esc(s) {
            return String(s === undefined || s === null ? '' : s)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
        }

        function fmtSize(n) {
            n = Number(n) || 0;
            if (n >= 1048576) return (n / 1048576).toFixed(1) + ' MB';
            if (n >= 1024) return Math.round(n / 1024) + ' KB';
            return n ? n + ' B' : '';
        }

        function ctxQS(extra) {
            return '?auth_user=' + encodeURIComponent(currentUsername) +
                '&store=' + encodeURIComponent(activeStore) +
                '&month=' + encodeURIComponent(selectedMonth) +
                '&year=' + encodeURIComponent(selectedYear) + (extra || '');
        }

        function navigateContext(store, month, year) {
            saveLocalOnly();
            navigateParent('?auth_user=' + encodeURIComponent(currentUsername) +
                '&store=' + encodeURIComponent(store) +
                '&month=' + encodeURIComponent(month) +
                '&year=' + encodeURIComponent(year));
        }

        function openUploadPanel(cat) {
            // Opens the computer's file picker through an invisible uploader (nothing is shown on screen).
            saveLocalOnly();
            if (String(cat).indexOf('rcpt') === 0) cat = 'receipts';
            try {
                const d = window.parent.document;
                let inp = d.querySelector('.st-key-up_' + cat + ' input[type=file]');
                if (!inp) {
                    alert('The file picker is still loading. Please try again in a moment.');
                    return;
                }
                inp.click();
            } catch (e) {
                alert('Could not open the file picker: ' + e);
            }
        }

        function navigateParent(queryString) {
            try {
                window.parent.location.search = queryString;
            } catch (e) {
                window.location.href = queryString;
            }
        }

        function initApp() {
            document.getElementById('select-month').value = selectedMonth;
            document.getElementById('select-year').value = selectedYear;
            document.getElementById('active-period-label').textContent = `${selectedMonth} ${selectedYear}`;
            document.getElementById('portal-period-label').textContent = `${selectedMonth} ${selectedYear}`;
            renderTabs();
            loadStorePeriodData();
            setupCSVParser();
            document.getElementById('workspace-store-name').textContent = activeStore;
            document.getElementById('portal-store-label').textContent = activeStore;
        }

        function renderTabs() {
            const storeTabsEl = document.getElementById('store-tabs');
            if(!storeTabsEl) return;
            storeTabsEl.innerHTML = stores.map(store => `
                <button onclick="switchStore(this.dataset.store)" data-store="${esc(store)}" class="px-3.5 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${activeStore === store ? 'bg-emerald-700 text-white shadow-md' : 'bg-white hover:bg-zinc-100 text-zinc-700 border border-zinc-300'}">
                    ${esc(store)}
                </button>
            `).join('');
        }

        function switchStore(store) {
            if (store === activeStore) return;
            navigateContext(store, selectedMonth, selectedYear);
        }

        function updatePeriod() {
            const m = document.getElementById('select-month').value;
            const y = document.getElementById('select-year').value;
            if (m === selectedMonth && y === selectedYear) return;
            navigateContext(activeStore, m, y);
        }

        function getSPK() {
            return `${activeStore}_${selectedMonth}_${selectedYear}`;
        }

        function loadStorePeriodData() {
            const spk = getSPK();

            document.getElementById('cooh-text').value = "";
            document.getElementById('cooh-date').value = "";
            document.getElementById('brinks-date').value = "";
            document.getElementById('chargebacks-text').value = "Nothing to report.";

            const checkboxes = ['brinks', 'green1', 'green2', 'green3', 'mon1', 'mon2', 'ocs1', 'ocs2', 'gas1', 'gas2'];
            checkboxes.forEach(k => {
                const chk = document.getElementById(`${k}-chk`);
                if(chk) chk.checked = false;
                const note = document.getElementById(`${k}-note`);
                if(note) note.value = "";
                syncNotes(k, "", false);
            });

            let data = persistentStoreData[spk];
            try {
                const local = localStorage.getItem('oz_data_' + spk);
                if (local) {
                    const parsedLocal = JSON.parse(local);
                    if (!data || parsedLocal._updated > (data._updated || 0)) {
                        data = parsedLocal;
                    }
                }
            } catch(e) {}

            if (data) {
                if(data.cooh_text !== undefined) {
                    document.getElementById('cooh-text').value = data.cooh_text;
                    syncNotes('cooh', data.cooh_text, false);
                }
                if(data.cooh_date) document.getElementById('cooh-date').value = data.cooh_date;
                if(data.chargebacks_text) {
                    document.getElementById('chargebacks-text').value = data.chargebacks_text;
                    syncNotes('chargebacks', data.chargebacks_text, false);
                }
                if(data.checklist_json) {
                    const chkData = data.checklist_json;
                    if(chkData.checkboxes) {
                        Object.keys(chkData.checkboxes).forEach(k => {
                            const chk = document.getElementById(`${k}-chk`);
                            if(chk) chk.checked = chkData.checkboxes[k];
                        });
                    }
                    if(chkData.notes) {
                        Object.keys(chkData.notes).forEach(k => {
                            const nEl = document.getElementById(`${k}-note`);
                            if(nEl) {
                                nEl.value = chkData.notes[k];
                                syncNotes(k, chkData.notes[k], false);
                            }
                        });
                    }
                    if(chkData.dates) {
                        if(chkData.dates.brinks) document.getElementById('brinks-date').value = chkData.dates.brinks;
                    }
                    if (chkData.expenses) {
                        const exp = chkData.expenses;
                        if (exp.rent !== undefined) { const el = document.getElementById('exp-rent'); if (el) el.value = exp.rent; }
                        if (exp.payroll !== undefined) { const el = document.getElementById('exp-payroll'); if (el) el.value = exp.payroll; }
                        if (exp.utilities !== undefined) { const el = document.getElementById('exp-utilities'); if (el) el.value = exp.utilities; }
                        if (exp.insurance !== undefined) { const el = document.getElementById('exp-insurance'); if (el) el.value = exp.insurance; }
                        if (exp.systemfees !== undefined) { const el = document.getElementById('exp-systemfees'); if (el) el.value = exp.systemfees; }
                        if (exp.repairs !== undefined) { const el = document.getElementById('exp-repairs'); if (el) el.value = exp.repairs; }
                        if (exp.other !== undefined) { const el = document.getElementById('exp-other'); if (el) el.value = exp.other; }
                    }
                }
                const invObj = data.inventory_data || data.inventory_json;
                if(invObj) {
                    storeInventoryData[spk] = {
                        headers: invObj.headers,
                        rows: invObj.rows,
                        fileName: data.inventory_filename || 'Uploaded_Inventory.csv'
                    };
                } else {
                    delete storeInventoryData[spk];
                }
            }

            renderInventoryTable();
            renderStoredFiles();
            updateCompletionBadge();
            updateFinancialSummary();
        }

        function setupCSVParser() {
            const dropZone = document.getElementById('drop-zone');
            const fileInput = document.getElementById('file-input');
            if(!dropZone || !fileInput) return;

            dropZone.onclick = () => fileInput.click();
            dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('border-emerald-600', 'bg-emerald-50/50'); };
            dropZone.ondragleave = () => dropZone.classList.remove('border-emerald-600', 'bg-emerald-50/50');
            dropZone.ondrop = (e) => {
                e.preventDefault();
                dropZone.classList.remove('border-emerald-600', 'bg-emerald-50/50');
                if (e.dataTransfer.files.length) handleCSVFile(e.dataTransfer.files[0]);
            };
            fileInput.onchange = (e) => {
                if (e.target.files.length) handleCSVFile(e.target.files[0]);
                fileInput.value = '';
            };
        }

        function handleCSVFile(file) {
            Papa.parse(file, {
                header: true,
                skipEmptyLines: true,
                complete: function(results) {
                    processParsedCSV(results.meta.fields, results.data, file.name);
                }
            });
        }

        function processParsedCSV(fields, data, fileName) {
            const parentRows = data.filter(row => {
                const cat = row['Category'] || '';
                return !cat.trim().startsWith('>');
            });

            const sortedRows = [];
            ACCOUNTANT_ORDER.forEach(targetCat => {
                const found = parentRows.find(r => (r['Category'] || '').trim() === targetCat.trim());
                if (found) {
                    sortedRows.push(found);
                } else {
                    const emptyRow = { Category: targetCat };
                    fields.forEach(f => { if (f !== 'Category') emptyRow[f] = 0.00; });
                    sortedRows.push(emptyRow);
                }
            });

            const spk = getSPK();
            storeInventoryData[spk] = { headers: fields, rows: sortedRows, fileName: fileName };

            const green1Chk = document.getElementById('green1-chk');
            if (green1Chk) green1Chk.checked = true;

            renderInventoryTable();
            showCSVFileInfo(fileName);
            renderStoredFiles();
            updateCompletionBadge();
            updateFinancialSummary();
            saveAllData(true);
        }

        function renderInventoryTable() {
            const spk = getSPK();
            const inv = storeInventoryData[spk];
            const tableHeader = document.getElementById('table-header');
            const tableBody = document.getElementById('table-body');
            const invStatusBadge = document.getElementById('inv-status-badge');
            const actionArea = document.getElementById('action-area');
            const portalInvVal = document.getElementById('portal-inv-val');
            const infoBox = document.getElementById('csv-file-info');

            if (!inv) {
                tableHeader.innerHTML = `<th class="p-2 font-semibold">Category</th><th class="p-2 font-semibold">Opening</th><th class="p-2 font-semibold">Sold</th><th class="p-2 font-semibold">Closing</th>`;
                tableBody.innerHTML = `<tr><td colspan="4" class="p-4 text-center text-zinc-400">No report uploaded yet.</td></tr>`;
                actionArea.classList.add('hidden');
                invStatusBadge.textContent = "Not Loaded";
                invStatusBadge.className = "text-[11px] text-zinc-600 bg-white border border-zinc-300 px-2 py-0.5 rounded shadow-xs";
                portalInvVal.innerHTML = '<span class="text-emerald-200/60">No CSV report uploaded</span>';
                if(infoBox) infoBox.classList.add('hidden');
                updateFinancialSummary();
                return;
            }

            const headers = inv.headers;
            const rows = inv.rows;
            const displayHeaders = headers.includes('Opening') ? ['Category', 'Opening', 'Sold', 'Closing'] : [headers[0], headers[1], headers[2]];

            tableHeader.innerHTML = displayHeaders.map(h => `<th class="p-2 font-semibold">${esc(h)}</th>`).join('');
            tableBody.innerHTML = rows.map(row => `
                <tr class="hover:bg-zinc-50 transition-colors">
                    ${displayHeaders.map(h => `<td class="p-2 ${h === 'Category' ? 'font-medium text-emerald-800' : 'text-zinc-700'}">${row[h] !== undefined ? esc(row[h]) : ''}</td>`).join('')}
                </tr>
            `).join('');

            actionArea.classList.remove('hidden');
            invStatusBadge.textContent = `Loaded (${inv.fileName})`;
            invStatusBadge.className = "text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded";
            portalInvVal.innerHTML = `<span class="text-emerald-300 font-semibold">✓ Ready (${esc(inv.fileName)})</span>`;

            showCSVFileInfo(inv.fileName);

            document.getElementById('download-excel-btn').onclick = function() {
                const worksheetData = [headers, ...rows.map(row => headers.map(h => row[h]))];
                const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, `${activeStore} Inventory`);
                XLSX.writeFile(workbook, `${activeStore}_${selectedMonth}_${selectedYear}_Formatted_Report.xlsx`);
            };
            updateFinancialSummary();
        }

        function showCSVFileInfo(fileName) {
            const infoBox = document.getElementById('csv-file-info');
            const nameSpan = document.getElementById('csv-file-name');
            if (infoBox && nameSpan) {
                nameSpan.textContent = fileName;
                infoBox.classList.remove('hidden');
            }
        }

        function clearCSVUpload() {
            if (confirm('Are you sure you want to delete this CSV upload?')) {
                const spk = getSPK();
                delete storeInventoryData[spk];
                const green1Chk = document.getElementById('green1-chk');
                if (green1Chk) green1Chk.checked = false;
                renderInventoryTable();
                renderStoredFiles();
                updateCompletionBadge();
                saveAllData(true);
            }
        }

        function syncNotes(key, val, triggerSave = true) {
            const noteCell = document.getElementById(`note-${key}`);
            if (noteCell) {
                noteCell.textContent = val.trim() !== '' ? val : (key === 'chargebacks' ? 'Nothing to report.' : 'No notes');
            }
            if (key === 'cooh') {
                const coohValCell = document.getElementById('portal-cooh-val');
                if (coohValCell) coohValCell.textContent = val.trim() !== '' ? val : 'Not entered';
            }
            if (triggerSave) saveLocalOnly();
        }

        function buildPayload() {
            const spk = getSPK();
            const checkboxes = {
                'brinks': document.getElementById('brinks-chk')?.checked || false,
                'green1': document.getElementById('green1-chk')?.checked || false,
                'green2': document.getElementById('green2-chk')?.checked || false,
                'green3': document.getElementById('green3-chk')?.checked || false,
                'mon1': document.getElementById('mon1-chk')?.checked || false,
                'mon2': document.getElementById('mon2-chk')?.checked || false,
                'ocs1': document.getElementById('ocs1-chk')?.checked || false,
                'ocs2': document.getElementById('ocs2-chk')?.checked || false,
                'gas1': document.getElementById('gas1-chk')?.checked || false,
                'gas2': document.getElementById('gas2-chk')?.checked || false
            };

            const notes = {
                'cooh': document.getElementById('cooh-text')?.value || '',
                'brinks': document.getElementById('brinks-note')?.value || '',
                'green1': document.getElementById('green1-note')?.value || '',
                'green2': document.getElementById('green2-note')?.value || '',
                'green3': document.getElementById('green3-note')?.value || '',
                'mon1': document.getElementById('mon1-note')?.value || '',
                'mon2': document.getElementById('mon2-note')?.value || '',
                'ocs1': document.getElementById('ocs1-note')?.value || '',
                'ocs2': document.getElementById('ocs2-note')?.value || '',
                'chargebacks': document.getElementById('chargebacks-text')?.value || 'Nothing to report.',
                'gas1': document.getElementById('gas1-note')?.value || '',
                'gas2': document.getElementById('gas2-note')?.value || ''
            };

            const dates = {
                'cooh': document.getElementById('cooh-date')?.value || '',
                'brinks': document.getElementById('brinks-date')?.value || ''
            };

            const expenses = {rent: document.getElementById('exp-rent')?.value || '', payroll: document.getElementById('exp-payroll')?.value || '', utilities: document.getElementById('exp-utilities')?.value || '', insurance: document.getElementById('exp-insurance')?.value || '', systemfees: document.getElementById('exp-systemfees')?.value || '', repairs: document.getElementById('exp-repairs')?.value || '', other: document.getElementById('exp-other')?.value || ''};

            const checklistData = {
                checkboxes: checkboxes,
                notes: notes,
                dates: dates,
                expenses: expenses,
                files: {}
            };
            const inv = storeInventoryData[spk] || null;

            return {
                _updated: Date.now(),
                spk: spk,
                store: activeStore,
                period: selectedMonth + ' ' + selectedYear,
                cooh_text: notes['cooh'],
                cooh_date: dates['cooh'],
                chargebacks_text: notes['chargebacks'],
                checklist_json: checklistData,
                inventory_data: inv ? { headers: inv.headers, rows: inv.rows } : null,
                inventory_filename: inv ? inv.fileName : ""
            };
        }

        function saveLocalOnly() {
            const spk = getSPK();
            const payload = buildPayload();
            persistentStoreData[spk] = payload;
            try {
                localStorage.setItem('oz_data_' + spk, JSON.stringify(payload));
            } catch(e) {}
            const badge = document.getElementById('store-completion-badge');
            if (badge) {
                const checked = Object.values(payload.checklist_json.checkboxes).filter(Boolean).length;
                badge.textContent = `✓ Draft Saved | Checklist: ${checked}/10`;
            }
        }

        function saveAllData(isSilent = false) {
            const payload = buildPayload();
            const spk = payload.spk;
            persistentStoreData[spk] = payload;
            try {
                localStorage.setItem('oz_data_' + spk, JSON.stringify(payload));
            } catch(e) {}

            const payloadEncoded = encodeURIComponent(JSON.stringify(payload));
            if (payloadEncoded.length > 60000) {
                alert('This store has too much data to save in one go. Please remove unneeded rows from the inventory and try again.');
                return;
            }
            navigateParent(ctxQS('&action=save_all&payload=' + payloadEncoded));
        }

        function updateCompletionBadge() {
            const checkboxes = ['brinks', 'green1', 'green2', 'green3', 'mon1', 'mon2', 'ocs1', 'ocs2', 'gas1', 'gas2'];
            let checked = 0;
            checkboxes.forEach(key => {
                const chk = document.getElementById(`${key}-chk`);
                if (chk && chk.checked) {
                    checked++;
                    const statusEl = document.getElementById(`status-${key}`);
                    if(statusEl) {
                        statusEl.textContent = "Complete ✓";
                        statusEl.className = "bg-emerald-800 text-emerald-200 px-2 py-0.5 rounded text-[11px]";
                    }
                } else {
                    const statusEl = document.getElementById(`status-${key}`);
                    if(statusEl) {
                        statusEl.textContent = "Pending";
                        statusEl.className = "bg-emerald-900 text-emerald-300/60 px-2 py-0.5 rounded text-[11px]";
                    }
                }
            });

            const badge = document.getElementById('store-completion-badge');
            if (badge) badge.textContent = `Checklist: ${checked}/10 Completed`;

            const portalChk = document.getElementById('portal-chk-val');
            if (portalChk) portalChk.textContent = `${checked} / 10 Items Checked`;

            const coohDate = document.getElementById('cooh-date')?.value;
            const brinksDate = document.getElementById('brinks-date')?.value;
            if(document.getElementById('portal-date-cooh')) document.getElementById('portal-date-cooh').textContent = coohDate || 'No date';
            if(document.getElementById('portal-date-brinks')) document.getElementById('portal-date-brinks').textContent = brinksDate || 'No date';
        }

        function handleItemFiles(cat, inputElem) {
            if (inputElem) inputElem.value = '';
            openUploadPanel(cat);
        }

        function handleRemoveFile(cat, fileId) {
            if (confirm('Remove this file permanently?')) {
                saveLocalOnly();
                navigateParent(ctxQS('&action=delete_file&cat=' + encodeURIComponent(cat) + '&file_id=' + encodeURIComponent(fileId)));
            }
        }

        function clearUploadField(cat) {
            if (confirm('Clear all files for this category?')) {
                saveLocalOnly();
                navigateParent(ctxQS('&action=clear_files&cat=' + encodeURIComponent(cat)));
            }
        }

        function renderStoredFiles() {
            const spk = getSPK();
            const cats = ['cooh', 'brinks', 'green1', 'green2', 'green3', 'mon1', 'mon2', 'ocs1', 'ocs2', 'chargebacks', 'gas1', 'gas2', 'receipts', ...Array.from({length: 50}, (_, i) => 'rcpt' + (i + 1))];

            cats.forEach(cat => {
                const box = document.getElementById(`files-${cat}`);
                const tableBox = document.getElementById(`table-files-${cat}`);
                if (!box) return;

                const fileList = (storedFilesMap[spk] && storedFilesMap[spk][cat]) ? storedFilesMap[spk][cat] : [];

                if (fileList.length > 0) {
                    const chip = (f, cls, xcls) => {
                        const size = fmtSize(f.size);
                        return `<span class="inline-block ${cls} px-1.5 py-0.5 rounded mr-1 mb-1 font-medium">📎 ${esc(f.filename)}${size ? ' <span class="opacity-60">(' + size + ')</span>' : ''}` +
                            ` <button type="button" data-cat="${esc(cat)}" data-id="${esc(f.id)}" onclick="handleRemoveFile(this.dataset.cat, this.dataset.id)" class="${xcls} font-bold ml-1">×</button></span>`;
                    };
                    box.innerHTML = fileList.map(f => chip(f, 'bg-emerald-100 text-emerald-800', 'text-red-600 hover:text-red-800')).join('');
                    if (tableBox) {
                        tableBox.innerHTML = fileList.map(f => chip(f, 'bg-emerald-800 text-emerald-200', 'text-red-400 hover:text-red-300')).join('');
                    }
                } else {
                    box.innerHTML = '<span class="text-zinc-400 italic">No files uploaded</span>';
                    if (tableBox) tableBox.innerHTML = '<span class="text-emerald-300/50 italic">No files attached</span>';
                }
            });
        }

        function openOwnerReport() {
            const payload = buildPayload();
            const encoded = encodeURIComponent(JSON.stringify(payload));
            if (encoded.length > 60000) {
                alert('Too much data to save before building the report. Press Save All to Cloud first.');
                return;
            }
            persistentStoreData[payload.spk] = payload;
            try { localStorage.setItem('oz_data_' + payload.spk, JSON.stringify(payload)); } catch(e) {}
            navigateParent(ctxQS('&action=save_all&then=owner&payload=' + encoded));
        }

        function downloadMasterZip() {
            // Save what is on screen first, then ask the server to build the real package
            // (Excel report + every uploaded document) from the saved data.
            const payload = buildPayload();
            const encoded = encodeURIComponent(JSON.stringify(payload));
            if (encoded.length > 60000) {
                alert('Too much data to save before building the package. Press Save All to Cloud first.');
                return;
            }
            persistentStoreData[payload.spk] = payload;
            try { localStorage.setItem('oz_data_' + payload.spk, JSON.stringify(payload)); } catch(e) {}
            navigateParent(ctxQS('&action=save_all&then=package&payload=' + encoded));
        }

        function openAddUserModal() {
            document.getElementById('edit-form-title').textContent = "Add New User Credential";
            document.getElementById('edit-orig-username').value = "";
            document.getElementById('edit-username').value = "";
            document.getElementById('edit-name').value = "";
            document.getElementById('edit-role').value = "store";
            document.getElementById('edit-password').value = "";
            document.getElementById('user-edit-form-box').classList.remove('hidden');
        }

        function editUser(username, name, role, password) {
            document.getElementById('edit-form-title').textContent = `Edit User: ${username}`;
            document.getElementById('edit-orig-username').value = username;
            document.getElementById('edit-username').value = username;
            document.getElementById('edit-name').value = name;
            document.getElementById('edit-role').value = role;
            document.getElementById('edit-password').value = password;
            document.getElementById('user-edit-form-box').classList.remove('hidden');
        }

        function submitUserEdit(event) {
            event.preventDefault();
            const orig = document.getElementById('edit-orig-username').value;
            const user = document.getElementById('edit-username').value;
            const name = document.getElementById('edit-name').value;
            const role = document.getElementById('edit-role').value;
            const pass = document.getElementById('edit-password').value;

            navigateParent(`?auth_user=${encodeURIComponent(currentUsername)}&action=update_user&orig_username=${encodeURIComponent(orig)}&username=${encodeURIComponent(user)}&name=${encodeURIComponent(name)}&role=${encodeURIComponent(role)}&password=${encodeURIComponent(pass)}`);
        }

        function deleteUser(username) {
            if (confirm(`Are you sure you want to delete user "${username}"?`)) {
                navigateParent(`?auth_user=${encodeURIComponent(currentUsername)}&action=delete_user&username=${encodeURIComponent(username)}`);
            }
        }

        function renderDocFacts(formatMoney) {
            const box = document.getElementById('doc-facts-box');
            if (!box) return;
            const docs = (docFacts && docFacts.docs) ? docFacts.docs : [];
            if (!docs.length) {
                box.innerHTML = ((docFacts && docFacts.error) ? '<p class="text-xs text-red-300 font-bold mb-2">Problem reading your documents: ' + esc(docFacts.error) + '</p>' : '') + '<p class="text-xs text-emerald-200/70">No documents uploaded for this store and month yet. Use the Browse buttons in the checklist to add the Greenline, Moneris, OCS and gas card reports. Their figures will appear here and feed the numbers above.</p>';
                return;
            }
            const val = (l) => (typeof l[1] === 'string') ? esc(l[1]) : ((l[0] === 'Days' || l[0] === 'Litres') ? Number(l[1]).toLocaleString(undefined, {maximumFractionDigits: 1}) : formatMoney(Number(l[1]) || 0));
            const result = (d) => {
                if (d.status === 'read' && d.lines && d.lines.length) {
                    return d.lines.map(l => '<span class="inline-block mr-3"><span class="text-emerald-300">' + esc(l[0]) + ':</span> <span class="font-bold text-white">' + val(l) + '</span></span>').join('');
                }
                if (d.status === 'unreadable') return '<span class="text-red-300 font-bold">Could not be opened from storage. Upload it again (use Browse, then Remove files that are missing from storage).</span><div class="text-[11px] text-red-200 mt-0.5">' + esc(d.error || '') + '</div>';
                if (d.status === 'attached') return '<span class="text-amber-300">Attached to the package (a photo, or numbers could not be read from this file).</span>';
                const cols = (d.cols || []).map(c => esc(c.name) + ' ' + formatMoney(Number(c.total) || 0)).join(' &middot; ');
                return '<span class="text-amber-300">Read ' + d.rows + ' rows but no sales/total line was recognised.</span><div class="text-[11px] text-emerald-300/70 mt-0.5">Columns: ' + esc((d.headers || []).join(', ')) + (cols ? '<br>Totals: ' + cols : '') + '</div>';
            };
            const extra = (docFacts && docFacts.extra) ? docFacts.extra : [];
            const errHtml = (docFacts && docFacts.error) ? '<p class="text-xs text-red-300 font-bold mb-2">Problem reading your documents: ' + esc(docFacts.error) + '</p>' : '';
            box.innerHTML = errHtml + '<h5 class="text-sm font-bold text-emerald-300 mb-3">Figures read from your uploaded documents</h5>' +
                '<div class="overflow-x-auto overflow-y-hidden rounded-lg border border-emerald-800"><table class="w-full text-left text-xs text-emerald-100">' +
                '<thead class="bg-emerald-800/60 text-emerald-300"><tr><th class="p-2">Document</th><th class="p-2">File</th><th class="p-2">What was read</th></tr></thead><tbody>' +
                docs.map(d => '<tr class="border-t border-emerald-800 align-top"><td class="p-2 font-medium text-white">' + esc(d.label) + '</td><td class="p-2">' + esc(d.file) + '</td><td class="p-2">' + result(d) + '</td></tr>').join('') +
                '</tbody></table></div>' +
                (extra.length ? '<p class="text-[11px] text-emerald-200/80 mt-2">Also counted as operating expenses from your documents: ' + extra.map(x => esc(x.label) + ' ' + formatMoney(Number(x.total) || 0)).join(' &middot; ') + '.</p>' : '') +
                '<p class="text-[11px] text-emerald-200/70 mt-2">Sales and cost of goods come from Greenline Financials. Card fees come from the Moneris statements and fuel from the gas card statements. The expense fields below are typed in by you. Press the Owner Report button for the full income statement.</p>';
        }

        function updateFinancialSummary() {
            try {
                const spk = getSPK();
                const inv = storeInventoryData[spk];
                let totalSales = 0;
                let totalCogs = 0;

                if (inv && inv.rows && inv.headers) {
                    const soldHeader = inv.headers.find(h => h && h.toLowerCase().includes('sold'));
                    const costHeader = inv.headers.find(h => h && (h.toLowerCase().includes('cost') || h.toLowerCase().includes('avg cost')));

                    inv.rows.forEach(row => {
                        let soldVal = 0;
                        if (soldHeader && row[soldHeader] !== undefined) {
                            soldVal = parseFloat(String(row[soldHeader]).replace(/[^0-9.-]/g, '')) || 0;
                        } else {
                            const keys = Object.keys(row).filter(k => k !== 'Category');
                            if (keys.length > 0) soldVal = parseFloat(String(row[keys[0]]).replace(/[^0-9.-]/g, '')) || 0;
                        }
                        totalSales += soldVal;

                        if (costHeader && row[costHeader] !== undefined) {
                            const costVal = parseFloat(String(row[costHeader]).replace(/[^0-9.-]/g, '')) || 0;
                            totalCogs += costVal;
                        }
                    });
                }

                const invSoldTotal = totalSales, invCostTotal = totalCogs;
                let salesSrc = invSoldTotal > 0 ? "Inventory sold column" : "";
                let cogsSrc = invCostTotal > 0 ? "Inventory cost column" : "";
                if (docFacts && docFacts.sales > 0) {
                    totalSales = docFacts.sales; salesSrc = docFacts.sales_src;
                    if (docFacts.cogs > 0) { totalCogs = docFacts.cogs; cogsSrc = docFacts.cogs_src; }
                    else if (invCostTotal === 0 && invSoldTotal > 0) { totalCogs = invSoldTotal; cogsSrc = 'Inventory sold, valued at cost'; }
                }
                if (totalCogs === 0 && totalSales > 0) { cogsSrc = "Estimated at 80% of sales"; }
                if (totalCogs === 0 && totalSales > 0) {
                    totalCogs = totalSales * 0.80;
                }

                const parseNum = (id) => {
                    const el = document.getElementById(id);
                    if (!el) return 0;
                    return parseFloat(String(el.value).replace(/[^0-9.-]/g, '')) || 0;
                };

                const expRent = parseNum('exp-rent');
                const expSal = parseNum('exp-salaries');
                const expUtil = parseNum('exp-utilities');
                const expAdv = parseNum('exp-advertising');
                const expMkt = parseNum('exp-marketing');
                const expData = parseNum('exp-datadeals');
                const expIns = parseNum('exp-insurance');
                const expGar = parseNum('exp-garbage');
                const expSec = parseNum('exp-security');
                const expLic = parseNum('exp-licenses');

                const extraExp = (docFacts && docFacts.extra ? docFacts.extra : []).reduce((a, x) => a + (Number(x.total) || 0), 0);
                const totalExp = expRent + expSal + expUtil + expAdv + expMkt + expData + expIns + expGar + expSec + expLic + extraExp;

                const grossProfit = totalSales - totalCogs;
                const grossMargin = totalSales > 0 ? (grossProfit / totalSales) * 100 : 0;

                const netOperatingIncome = grossProfit - totalExp;
                const operatingMargin = totalSales > 0 ? (netOperatingIncome / totalSales) * 100 : 0;

                const formatMoney = (n) => '$' + n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                const formatPct = (n) => n.toFixed(1) + '%';

                const salesEl = document.getElementById('fin-total-sales');
                const cogsEl = document.getElementById('fin-cogs');
                const grossMEl = document.getElementById('fin-gross-margin');
                const netEl = document.getElementById('fin-net');
                const netMEl = document.getElementById('fin-net-margin');
                const totExpEl = document.getElementById('fin-total-expenses');

                if (salesEl) salesEl.textContent = formatMoney(totalSales);
                if (cogsEl) cogsEl.textContent = formatMoney(totalCogs);
                if (grossMEl) grossMEl.textContent = 'Gross Margin: ' + formatPct(grossMargin);
                if (netEl) {
                    netEl.textContent = formatMoney(netOperatingIncome);
                    netEl.className = netOperatingIncome >= 0 ? 'text-lg font-bold text-white' : 'text-lg font-bold text-red-300';
                }
                if (netMEl) {
                    netMEl.textContent = formatPct(operatingMargin);
                    netMEl.className = operatingMargin >= 0 ? 'text-lg font-bold text-amber-300' : 'text-lg font-bold text-red-300';
                }
                if (totExpEl) totExpEl.textContent = formatMoney(totalExp);
                const srcLabel = (el, text, fallback) => {
                    const lab = el ? el.previousElementSibling : null;
                    if (lab) lab.textContent = text ? fallback + ' - ' + text : fallback;
                };
                srcLabel(salesEl, salesSrc, 'Total Sales');
                srcLabel(cogsEl, cogsSrc, 'Cost of Goods Sold (COGS)');
                renderDocFacts(formatMoney);
            } catch(e) {
                console.error("updateFinancialSummary error:", e);
            }
        }


        // ===== Expense report (redesigned) =====
        const EXP_LABELS = [["rent","Rent"],["payroll","Payroll"],["utilities","Utilities"],["insurance","Insurance"],["systemfees","System Fees"],["repairs","Repairs"],["other","Other"]];
        let expenseReportShown = false;
        try { expenseReportShown = sessionStorage.getItem('oz_exp_report') === '1'; } catch(e) {}

        function moneyFmt(n) { n = Number(n) || 0; return (n < 0 ? '-$' : '$') + Math.abs(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
        function pctFmt(n) { return (Number(n) || 0).toFixed(1) + '%'; }

        function computeFinancials() {
            const spk = getSPK();
            const inv = storeInventoryData[spk];
            let sales = 0, cogs = 0;
            if (inv && inv.rows && inv.headers) {
                const soldHeader = inv.headers.find(h => h && h.toLowerCase().includes('sold'));
                const costHeader = inv.headers.find(h => h && h.toLowerCase().includes('cost'));
                inv.rows.forEach(row => {
                    let sv = 0;
                    if (soldHeader && row[soldHeader] !== undefined) sv = parseFloat(String(row[soldHeader]).replace(/[^0-9.-]/g, '')) || 0;
                    sales += sv;
                    if (costHeader && row[costHeader] !== undefined) cogs += parseFloat(String(row[costHeader]).replace(/[^0-9.-]/g, '')) || 0;
                });
            }
            const invSold = sales, invCost = cogs;
            let salesSrc = invSold > 0 ? 'Inventory sold column' : '', cogsSrc = invCost > 0 ? 'Inventory cost column' : '';
            if (docFacts && docFacts.sales > 0) {
                sales = docFacts.sales; salesSrc = docFacts.sales_src;
                if (docFacts.cogs > 0) { cogs = docFacts.cogs; cogsSrc = docFacts.cogs_src; }
                else if (invCost === 0 && invSold > 0) { cogs = invSold; cogsSrc = 'Inventory sold, valued at cost'; }
            }
            if (cogs === 0 && sales > 0) { cogs = sales * 0.80; cogsSrc = 'Estimated at 80% of sales (upload Greenline Financials for the real figure)'; }
            const num = id => { const el = document.getElementById(id); return el ? (parseFloat(String(el.value).replace(/[^0-9.-]/g, '')) || 0) : 0; };
            const lines = EXP_LABELS.map(([k, label]) => ({ label: label, amount: num('exp-' + k), src: 'Entered' }));
            const extra = ((docFacts && docFacts.extra) ? docFacts.extra : []).map(x => ({ label: x.label, amount: Number(x.total) || 0, src: 'From documents' }));
            const all = lines.concat(extra);
            const totalExp = all.reduce((a, x) => a + x.amount, 0);
            const gross = sales - cogs, noi = gross - totalExp;
            return { sales, cogs, gross, grossMargin: sales > 0 ? gross / sales * 100 : 0, lines: all, totalExp,
                     expPct: sales > 0 ? totalExp / sales * 100 : 0, noi, opMargin: sales > 0 ? noi / sales * 100 : 0, salesSrc, cogsSrc };
        }


        const INFO_COLORS = ['#059669','#d97706','#2563eb','#7c3aed','#db2777','#0891b2','#65a30d','#dc2626','#475569','#ca8a04'];
        function infographicHtml(f) {
            const sales = f.sales;
            const tile = (label, val, sub, cls) => '<div class="rounded-lg border border-zinc-200 bg-zinc-50 p-3"><div class="text-[11px] text-zinc-500">' + label + '</div><div class="text-lg font-bold ' + cls + '">' + val + '</div><div class="text-[11px] text-zinc-500">' + sub + '</div></div>';
            const tiles = '<div class="grid grid-cols-2 md:grid-cols-4 gap-2 mb-4">' +
                tile('Sales', moneyFmt(f.sales), 'for the month', 'text-zinc-900') +
                tile('Gross profit', moneyFmt(f.gross), pctFmt(f.grossMargin) + ' margin', f.gross < 0 ? 'text-red-600' : 'text-emerald-700') +
                tile('Total expenses', moneyFmt(f.totalExp), pctFmt(f.expPct) + ' of sales', 'text-amber-700') +
                tile('Net operating income', moneyFmt(f.noi), pctFmt(f.opMargin) + ' margin', f.noi < 0 ? 'text-red-600' : 'text-emerald-700') + '</div>';
            // 1) where each sales dollar went
            let bar = '';
            if (sales > 0) {
                const cp = Math.max(0, f.cogs / sales * 100), ep = Math.max(0, f.totalExp / sales * 100);
                let pp = Math.max(0, 100 - cp - ep);
                const tot = cp + ep + pp;
                const seg = (p, color, label) => p > 0 ? '<div style="width:' + (p / tot * 100) + '%;background:' + color + '" class="h-full flex items-center justify-center text-[11px] font-bold text-white overflow-hidden whitespace-nowrap" title="' + label + '">' + (p >= 8 ? p.toFixed(0) + '%' : '') + '</div>' : '';
                bar = '<div class="mb-4"><div class="text-xs font-bold text-zinc-700 mb-1">Where each $1 of sales went</div>' +
                    '<div class="flex h-8 rounded-lg overflow-hidden border border-zinc-200">' + seg(cp, '#d97706', 'Cost of goods') + seg(ep, '#2563eb', 'Expenses') + seg(pp, '#059669', 'Profit') + '</div>' +
                    '<div class="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-[11px] text-zinc-600">' +
                    '<span><span class="inline-block w-2.5 h-2.5 rounded-sm mr-1" style="background:#d97706"></span>Cost of goods ' + cp.toFixed(1) + '%</span>' +
                    '<span><span class="inline-block w-2.5 h-2.5 rounded-sm mr-1" style="background:#2563eb"></span>Expenses ' + ep.toFixed(1) + '%</span>' +
                    '<span><span class="inline-block w-2.5 h-2.5 rounded-sm mr-1" style="background:#059669"></span>' + (f.noi >= 0 ? 'Profit ' + (f.noi / sales * 100).toFixed(1) + '%' : 'Loss ' + (-f.noi / sales * 100).toFixed(1) + '% (expenses and cost exceed sales)') + '</span></div></div>';
            }
            // 2) expense breakdown donut
            const parts = f.lines.filter(x => x.amount > 0);
            let donut = '';
            if (parts.length) {
                const R = 40, C = 2 * Math.PI * R, total = parts.reduce((a, x) => a + x.amount, 0);
                let off = 0;
                const arcs = parts.map((x, i) => {
                    const len = x.amount / total * C;
                    const s = '<circle cx="60" cy="60" r="' + R + '" fill="none" stroke="' + INFO_COLORS[i % INFO_COLORS.length] + '" stroke-width="22" stroke-dasharray="' + len + ' ' + (C - len) + '" stroke-dashoffset="' + (-off) + '" transform="rotate(-90 60 60)"><title>' + esc(x.label) + ' ' + moneyFmt(x.amount) + '</title></circle>';
                    off += len; return s;
                }).join('');
                const legend = parts.map((x, i) => '<div class="flex items-center justify-between gap-3 text-xs"><span class="flex items-center"><span class="inline-block w-2.5 h-2.5 rounded-sm mr-1.5" style="background:' + INFO_COLORS[i % INFO_COLORS.length] + '"></span>' + esc(x.label) + '</span><span class="text-zinc-600 whitespace-nowrap">' + moneyFmt(x.amount) + ' &middot; ' + (x.amount / total * 100).toFixed(0) + '%</span></div>').join('');
                donut = '<div class="mb-4"><div class="text-xs font-bold text-zinc-700 mb-2">Where the expenses went</div><div class="flex flex-wrap items-center gap-4">' +
                    '<svg viewBox="0 0 120 120" width="150" height="150" role="img" aria-label="Expense breakdown">' + arcs +
                    '<text x="60" y="57" text-anchor="middle" font-size="8" fill="#71717a">Total</text><text x="60" y="68" text-anchor="middle" font-size="9" font-weight="bold" fill="#18181b">' + moneyFmt(total) + '</text></svg>' +
                    '<div class="flex-1 min-w-[220px] space-y-1">' + legend + '</div></div></div>';
            }
            return tiles + bar + donut;
        }

        function generateExpenseReport(userClick) {
            const box = document.getElementById('expense-report');
            if (!box) return;
            expenseReportShown = true;
            try { sessionStorage.setItem('oz_exp_report', '1'); } catch(e) {}
            const f = computeFinancials();
            const row = (label, val, cls, sub) => '<tr class="border-t border-zinc-200"><td class="py-1.5 pr-3 ' + (cls || '') + '">' + label +
                (sub ? '<div class="text-[11px] font-normal text-zinc-500">' + esc(sub) + '</div>' : '') + '</td><td class="py-1.5 text-right whitespace-nowrap ' + (cls || '') + '">' + val + '</td></tr>';
            const neg = v => v < 0 ? 'text-red-600' : 'text-emerald-800';
            const expRows = f.lines.filter(x => x.amount !== 0 || x.src === 'Entered').map(x => row(esc(x.label), moneyFmt(x.amount), 'text-zinc-800', x.src === 'From documents' ? 'from your uploaded documents' : '')).join('');
            const noData = f.sales === 0 ? '<p class="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded p-2 mb-3">No sales figures were found for this store and month yet. Upload the Greenline Financials report (and the others) in the checklist above, tick them and press Save All to Cloud.</p>' : '';
            box.innerHTML = '<div class="bg-white text-zinc-900 rounded-xl p-5 shadow">' +
                '<h5 class="text-base font-bold text-emerald-900">Expense Report - ' + esc(activeStore) + ' &middot; ' + esc(selectedMonth) + ' ' + esc(selectedYear) + '</h5>' +
                '<p class="text-xs text-zinc-500 mb-3">Simple monthly summary with margins</p>' + noData + infographicHtml(f) +
                '<table class="w-full text-sm"><tbody>' +
                row('Sales', moneyFmt(f.sales), 'font-bold text-zinc-900', f.salesSrc) +
                row('Cost of goods sold', moneyFmt(-f.cogs), 'text-zinc-800', f.cogsSrc) +
                row('Gross profit', moneyFmt(f.gross), 'font-bold ' + neg(f.gross)) +
                row('Gross margin', pctFmt(f.grossMargin), 'text-zinc-600') +
                '<tr><td colspan="2" class="pt-4 pb-1 text-xs font-bold uppercase tracking-wide text-emerald-800">Expenses</td></tr>' + expRows +
                row('Total expenses', moneyFmt(f.totalExp), 'font-bold text-zinc-900') +
                row('Expenses as % of sales', pctFmt(f.expPct), 'text-zinc-600') +
                '<tr class="border-t-2 border-emerald-700"><td class="pt-2 font-bold text-base ' + neg(f.noi) + '">Net operating income</td><td class="pt-2 text-right font-bold text-base ' + neg(f.noi) + '">' + moneyFmt(f.noi) + '</td></tr>' +
                row('Operating margin', pctFmt(f.opMargin), 'font-bold ' + neg(f.opMargin)) +
                '</tbody></table>' +
                '<details class="mt-4"><summary class="cursor-pointer text-xs font-bold text-emerald-800">Figures read from your uploaded documents</summary><div id="doc-facts-box" class="mt-2 bg-emerald-950 rounded-lg p-3 text-emerald-100"></div></details>' +
                '<p class="text-[11px] text-zinc-500 mt-3">Use the Download Detailed Accounting Report button for the full Excel report built from every uploaded file.</p></div>';
            try { renderDocFacts(moneyFmt); } catch(e) {}
            if (userClick) {
                saveLocalOnly();
                setTimeout(() => { try { box.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); } catch(e) {} }, 100);
            }
        }

        function onExpenseInput() {
            saveLocalOnly();
            if (expenseReportShown) generateExpenseReport(false);
        }

        function updateFinancialSummary() {
            if (expenseReportShown) { try { generateExpenseReport(false); } catch(e) { console.error(e); } }
        }

        function downloadDetailedReport() {
            const payload = buildPayload();
            const encoded = encodeURIComponent(JSON.stringify(payload));
            if (encoded.length > 60000) {
                alert('Too much data to save before building the report. Press Save All to Cloud first.');
                return;
            }
            persistentStoreData[payload.spk] = payload;
            try { localStorage.setItem('oz_data_' + payload.spk, JSON.stringify(payload)); } catch(e) {}
            navigateParent(ctxQS('&action=save_all&then=excel&payload=' + encoded));
        }

        if (expenseReportShown) {
            setTimeout(() => {
                try { generateExpenseReport(false); } catch(e) {}
                try {
                    if (sessionStorage.getItem('oz_exp_scroll') === '1') {
                        sessionStorage.removeItem('oz_exp_scroll');
                        setTimeout(() => { const b = document.getElementById('expense-report'); if (b) b.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 400);
                    }
                } catch(e) {}
            }, 300);
        }

        initApp();
    </script>
<script>
(function(){
  // The page scrolls inside its own window, which is sized to fill the browser. Everything around it is locked,
  // so there is exactly one scrollbar.
  function fit(){
    try{
      var fe = window.frameElement, p = window.parent;
      if (!fe) return;
      var d = p.document;
      var sc = [d.querySelector('[data-testid="stMain"]'), d.querySelector('section.main'), d.querySelector('[data-testid="stAppViewContainer"]')];
      sc.forEach(function(e){ if (e && e.scrollTop) e.scrollTop = 0; });
      if (d.scrollingElement && d.scrollingElement.scrollTop) d.scrollingElement.scrollTop = 0;
      var top = fe.getBoundingClientRect().top;
      var h = Math.max(320, Math.floor(p.innerHeight - top - 2));
      if (Math.abs((parseInt(fe.style.height) || 0) - h) > 2) fe.style.height = h + 'px';
    }catch(e){}
  }
  window.addEventListener('pagehide', function(){
    try{ sessionStorage.setItem('oz_scroll', String(window.scrollY || document.documentElement.scrollTop || 0)); }catch(e){}
  });
  (function restore(){
    var y = 0;
    try{ y = parseInt(sessionStorage.getItem('oz_scroll') || '0') || 0; }catch(e){}
    if (!y) return;
    var tries = 0;
    var t = setInterval(function(){
      tries++;
      window.scrollTo(0, y);
      var cur = window.scrollY || document.documentElement.scrollTop || 0;
      if (Math.abs(cur - y) < 3 || tries > 25){ clearInterval(t); try{ sessionStorage.removeItem('oz_scroll'); }catch(e){} }
    }, 120);
  })();
  window.addEventListener('load', fit);
  try { window.parent.addEventListener('resize', fit); } catch(e) {}
  window.addEventListener('resize', fit);
  setInterval(fit, 700);
  [100, 400, 1000, 2500].forEach(function(t){ setTimeout(fit, t); });
  fit();
})();
</script></body>
</html>"""


# ============================================================================
# STREAMLIT APP  (original HTML interface + real file storage + report engine)
# ============================================================================
import hashlib
import hmac
import html as htmllib
import inspect
import json
import logging
import os
import secrets as pysecrets
import sqlite3
import tempfile
import uuid
from contextlib import closing
from datetime import date
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

# Suppress deprecation warning for components.html
logging.getLogger("streamlit.runtime.scriptrunner_utils.script_run_context").setLevel(logging.ERROR)
log = logging.getLogger("oz_store")

# --- Page config ---
st.set_page_config(
    page_title="The OZ Store - Persistent Accountant Package Manager",
    page_icon="🌿",
    layout="wide",
)

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding: 1.25rem 1.5rem 2rem 1.5rem; max-width: 100%; }
    [data-testid="stExpander"], [data-testid="stAlert"] { margin: 0 0 0.75rem 0; }
    .stButton>button, .stDownloadButton>button { background-color: #047857; color: white; border-radius: 12px; font-weight: bold; border: none; }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #065f46; color: white; }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


def _stretch():
    """Full-width widgets: newer Streamlit uses width='stretch', older use_container_width."""
    try:
        if "width" in inspect.signature(st.button).parameters:
            return {"width": "stretch"}
    except Exception:
        pass
    return {"use_container_width": True}


STRETCH = _stretch()

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
DROPBOX_CREDENTIALS_FILE = "dropbox_credentials.json"
DROPBOX_FOLDER_PATH = "/OZ_Store_Uploads"
MASTER_FILE = "master_store_state.json"
USERS_FILE = "users.json"
DB_FILE = "oz_store.db"
APP_VERSION = "2026-10-24 r6"
USERS_RESET_MARK = "users_reset_" + APP_VERSION.split()[0]
# NOTE: not called "uploads": a leftover *file* with that name crashed the old app
LOCAL_FILES_DIR = Path(tempfile.gettempdir()) / "oz_store_files"
MAX_UPLOAD_MB = 25

STORES = [f"Oz {i}" for i in range(1, 8)]
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
YEARS = ["2025", "2026", "2027", "2028"]
ROLES = ("store", "accountant", "admin")

UPLOAD_CATS = [
    ("cooh", "Cash On Hand - End of Day"),
    ("brinks", "Brinks Book"),
    ("green1", "Greenline: Inventory Summary"),
    ("green2", "Greenline: Financials"),
    ("green3", "Greenline: Payment Types"),
    ("mon1", "Moneris: Counter Sales"),
    ("mon2", "Moneris: Online Sales"),
    ("ocs1", "OCS: Account Statement"),
    ("ocs2", "OCS: Invoices"),
    ("chargebacks", "Moneris Chargebacks"),
    ("gas1", "Gas Cards: Details"),
    ("gas2", "Gas Cards: Summary"),
    ("receipts", "Credit Card Receipts"),
] + [(f"rcpt{_n}", f"Credit Card Receipt #{_n}") for _n in range(1, 51)]
DOC_LABELS = dict(UPLOAD_CATS)

DEFAULT_USERS = [
    ("admin", "Master Admin", "admin", "admin2026"),
    ("leanne", "Leanne (Accountant)", "accountant", "leanne2026"),
    ("oz1", "Oz 1", "store", "oz1pass2026"),
    ("oz2", "Oz 2", "store", "oz2pass2026"),
    ("oz3", "Oz 3", "store", "oz3pass2026"),
    ("oz4", "Oz 4", "store", "oz4pass2026"),
    ("oz5", "Oz 5", "store", "oz5pass2026"),
    ("oz6", "Oz 6", "store", "oz6pass2026"),
    ("oz7", "Oz 7", "store", "oz7pass2026"),
]


def secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# ----------------------------------------------------------------------------
# Dropbox
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _make_dropbox_client(key, sec, rt, token, cred_file_token):
    import dropbox
    if key and sec and rt:
        return dropbox.Dropbox(oauth2_refresh_token=rt, app_key=key, app_secret=sec)
    t = token or cred_file_token
    return dropbox.Dropbox(t) if t else None


def get_dropbox_client():
    """Refresh-token login (never expires) if configured, else the access token."""
    try:
        import dropbox
    except ImportError:
        return None
    try:
        key, sec, rt = secret("DROPBOX_APP_KEY"), secret("DROPBOX_APP_SECRET"), secret("DROPBOX_REFRESH_TOKEN")
        token = secret("DROPBOX_ACCESS_TOKEN")
        cred_tok = None
        if not (key and sec and rt) and not token and os.path.exists(DROPBOX_CREDENTIALS_FILE):
            with open(DROPBOX_CREDENTIALS_FILE, "r") as f:
                cred_tok = json.load(f).get("access_token")
        return _make_dropbox_client(key, sec, rt, token, cred_tok)
    except Exception:
        log.exception("Dropbox client could not be created")
    return None


def _dbx_put(dbx, path: str, data: bytes):
    import dropbox
    dbx.files_upload(data, path, mode=dropbox.files.WriteMode.overwrite)


def _dbx_get(dbx, path: str) -> bytes:
    _, res = dbx.files_download(path)
    return res.content


def _dbx_missing(exc) -> bool:
    try:
        return exc.error.is_path() and exc.error.get_path().is_not_found()
    except Exception:
        return False


# ----------------------------------------------------------------------------
# Database (same tables and columns as before, so existing data keeps working)
# ----------------------------------------------------------------------------
def _connect():
    return closing(sqlite3.connect(DB_FILE))


def _loads(s, default):
    try:
        v = json.loads(s) if s else default
        return default if v is None else v
    except Exception:
        return default


def load_users_from_dropbox():
    dbx = get_dropbox_client()
    if not dbx:
        return None
    try:
        users = json.loads(_dbx_get(dbx, f"{DROPBOX_FOLDER_PATH}/{USERS_FILE}"))
        ok = [tuple(u) for u in users if isinstance(u, list) and len(u) == 4]
        return ok or None
    except Exception:
        return None


def sync_users_to_dropbox():
    dbx = get_dropbox_client()
    if not dbx:
        return
    try:
        with _connect() as conn:
            rows = conn.execute("SELECT username, name, role, password FROM users").fetchall()
        _dbx_put(dbx, f"{DROPBOX_FOLDER_PATH}/{USERS_FILE}", json.dumps([list(r) for r in rows]).encode())
    except Exception:
        log.exception("Could not back up users to Dropbox")


def init_db():
    try:
        LOCAL_FILES_DIR.mkdir(parents=True, exist_ok=True)
    except OSError:
        log.exception("Could not create local file folder")
    with _connect() as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT, password TEXT)")
        cur.execute(
            "CREATE TABLE IF NOT EXISTS store_data (store_period_key TEXT PRIMARY KEY, store TEXT, period TEXT, "
            "cooh_text TEXT, cooh_date TEXT, chargebacks_text TEXT, checklist_json TEXT, inventory_json TEXT, "
            "inventory_filename TEXT)"
        )
        cur.execute(
            "CREATE TABLE IF NOT EXISTS store_files (id INTEGER PRIMARY KEY AUTOINCREMENT, store_period_key TEXT, "
            "category TEXT, filename TEXT, filepath TEXT)"
        )
        # Add any column an older version of the app didn't have
        existing = {c[1] for c in cur.execute("PRAGMA table_info(store_data)")}
        for col in ["store", "period", "cooh_text", "cooh_date", "chargebacks_text",
                    "checklist_json", "inventory_json", "inventory_filename"]:
            if col not in existing:
                cur.execute(f"ALTER TABLE store_data ADD COLUMN {col} TEXT")
        # One-time repair: an old database may hold users/passwords nobody can log in with.
        cur.execute("CREATE TABLE IF NOT EXISTS app_meta (k TEXT PRIMARY KEY, v TEXT)")
        done = cur.execute("SELECT v FROM app_meta WHERE k = ?", (USERS_RESET_MARK,)).fetchone()
        if not done:
            cur.execute("DELETE FROM users")
            cur.executemany("INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?)", DEFAULT_USERS)
            cur.execute("INSERT OR REPLACE INTO app_meta VALUES (?, '1')", (USERS_RESET_MARK,))
        elif cur.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            cur.executemany("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?)",
                            load_users_from_dropbox() or DEFAULT_USERS)
        conn.commit()


def get_user(username):
    with _connect() as conn:
        return conn.execute("SELECT name, role, password FROM users WHERE username = ?",
                            ((username or "").lower().strip(),)).fetchone()


def normalize_inventory(inv):
    """The page expects {headers, rows}. Accept the other shapes older versions saved."""
    if isinstance(inv, dict) and isinstance(inv.get("rows"), list):
        rows = [r for r in inv["rows"] if isinstance(r, dict)]
        headers = inv.get("headers") or (list(rows[0].keys()) if rows else [])
        return {"headers": headers, "rows": rows} if rows else None
    if isinstance(inv, list) and inv and isinstance(inv[0], dict):
        return {"headers": list(inv[0].keys()), "rows": inv}
    return None


def normalize_remote(d: dict) -> dict:
    """Bring any saved format (including the short Streamlit-widget versions) to the page format."""
    cj = d.get("checklist_json")
    if not isinstance(cj, dict):
        notes = dict(d.get("notes") or {})
        cj = {"checkboxes": d.get("checklist") or {}, "notes": notes,
              "dates": {}, "expenses": d.get("expenses") or {}, "files": {}}
        cj["notes"].setdefault("cooh", d.get("cooh_text", ""))
        cj["notes"].setdefault("chargebacks", d.get("chargebacks_text", ""))
    cj.setdefault("files", {})
    if not isinstance(cj["files"], dict):
        cj["files"] = {}
    return {
        "cooh_text": d.get("cooh_text") or "",
        "cooh_date": d.get("cooh_date") or "",
        "chargebacks_text": d.get("chargebacks_text") or "Nothing to report.",
        "checklist_json": cj,
        "inventory_data": normalize_inventory(d.get("inventory_data") or d.get("inventory_json")),
        "inventory_filename": d.get("inventory_filename") or "",
        "_updated": d.get("_updated") or cj.get("_updated") or 0,
    }


_REC_COLS = ("store_period_key, cooh_text, cooh_date, chargebacks_text, checklist_json, "
             "inventory_json, inventory_filename")


def _row_to_rec(r) -> dict:
    cj = _loads(r[4], {})
    return normalize_remote({
        "cooh_text": r[1], "cooh_date": r[2], "chargebacks_text": r[3], "checklist_json": cj,
        "inventory_data": _loads(r[5], None), "inventory_filename": r[6],
        "_updated": cj.get("_updated", 0) if isinstance(cj, dict) else 0,
    })


def db_all() -> dict:
    with _connect() as conn:
        rows = conn.execute(f"SELECT {_REC_COLS} FROM store_data").fetchall()
    return {r[0]: _row_to_rec(r) for r in rows}


def get_record(spk: str):
    with _connect() as conn:
        r = conn.execute(f"SELECT {_REC_COLS} FROM store_data WHERE store_period_key = ?", (spk,)).fetchone()
    return _row_to_rec(r) if r else None


def blank_record() -> dict:
    return normalize_remote({"checklist_json": {"checkboxes": {}, "notes": {}, "dates": {}, "expenses": {}, "files": {}}})


def save_record(spk: str, rec: dict):
    store, _, rest = spk.partition("_")
    cj = dict(rec.get("checklist_json") or {})
    cj["_updated"] = rec.get("_updated") or cj.get("_updated") or 0
    inv = normalize_inventory(rec.get("inventory_data"))
    with _connect() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO store_data (store_period_key, store, period, cooh_text, cooh_date, "
            "chargebacks_text, checklist_json, inventory_json, inventory_filename) VALUES (?,?,?,?,?,?,?,?,?)",
            (spk, store, rest.replace("_", " "), rec.get("cooh_text", ""), rec.get("cooh_date", ""),
             rec.get("chargebacks_text", "Nothing to report."), json.dumps(cj),
             json.dumps(inv) if inv else None, rec.get("inventory_filename", "")),
        )
        conn.commit()


def sync_all_to_dropbox():
    """Write the master JSON that the phone reads. Returns (ok, error text)."""
    dbx = get_dropbox_client()
    if not dbx:
        return False, "Dropbox isn't configured in Streamlit secrets, so data is only on this server."
    try:
        _dbx_put(dbx, f"{DROPBOX_FOLDER_PATH}/{MASTER_FILE}",
                 json.dumps(db_all(), indent=2, default=str).encode("utf-8"))
        return True, ""
    except Exception as e:
        log.exception("Dropbox sync failed")
        return False, str(e)


@st.cache_data(ttl=120, show_spinner=False)
def restore_missing_from_dropbox():
    """After a reboot the server database is empty: pull anything missing back from Dropbox."""
    dbx = get_dropbox_client()
    if not dbx:
        return ""
    try:
        remote = json.loads(_dbx_get(dbx, f"{DROPBOX_FOLDER_PATH}/{MASTER_FILE}"))
    except Exception:
        return ""
    if not isinstance(remote, dict):
        return ""
    with _connect() as conn:
        have = {r[0] for r in conn.execute("SELECT store_period_key FROM store_data")}
    for spk, d in remote.items():
        if spk not in have and isinstance(d, dict):
            try:
                save_record(spk, normalize_remote(d))
            except Exception:
                log.exception("Could not restore %s", spk)
    return "ok"


# ----------------------------------------------------------------------------
# Uploaded documents: stored for real (Dropbox, or this server as a fallback)
# ----------------------------------------------------------------------------
def safe_name(n: str) -> str:
    return (re.sub(r"[^\w.\- ]", "_", n or "").strip() or "file")[:120]


def store_file(spk: str, cat: str, name: str, data: bytes):
    """Returns (meta, warning). meta is what we keep in the checklist record."""
    fid = uuid.uuid4().hex[:10]
    fname = f"{fid}_{safe_name(name)}"
    meta = {"id": fid, "filename": name, "size": len(data), "uploaded": datetime.now().isoformat(timespec="seconds")}
    warn = None
    dbx = get_dropbox_client()
    if dbx:
        path = f"{DROPBOX_FOLDER_PATH}/files/{spk}/{cat}/{fname}"
        try:
            _dbx_put(dbx, path, data)
            return {**meta, "backend": "dropbox", "path": path}, None
        except Exception as e:
            log.exception("Dropbox upload failed")
            warn = f"Dropbox upload failed ({e}); kept on the server only."
    p = LOCAL_FILES_DIR / spk / cat / fname
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    return {**meta, "backend": "local", "path": str(p)}, warn


@st.cache_data(ttl=600, show_spinner=False, max_entries=128)
def _dbx_bytes(path: str) -> bytes:
    dbx = get_dropbox_client()
    if not dbx:
        raise RuntimeError("Dropbox not configured")
    return _dbx_get(dbx, path)


_READ_ERR: dict = {}


def read_file_bytes(meta: dict):
    try:
        if meta.get("backend") == "dropbox":
            return _dbx_bytes(meta["path"])
        p = Path(meta.get("path", ""))
        return p.read_bytes() if p.is_file() else None
    except Exception as e:
        log.exception("Could not read %s", meta.get("filename"))
        _READ_ERR[meta.get("id")] = f"{type(e).__name__}: {e}"
        return None


def delete_file_bytes(meta: dict):
    try:
        if meta.get("backend") == "dropbox":
            dbx = get_dropbox_client()
            if dbx:
                dbx.files_delete_v2(meta["path"])
        else:
            Path(meta.get("path", "")).unlink(missing_ok=True)
    except Exception:
        log.exception("Could not delete %s", meta.get("filename"))


def add_files_to_record(spk: str, cat: str, metas: list):
    rec = get_record(spk) or blank_record()
    files = rec["checklist_json"].setdefault("files", {})
    files.setdefault(cat, []).extend(metas)
    save_record(spk, rec)


# ----------------------------------------------------------------------------
# Report engine glue: saved record -> engine payload -> Excel + ZIP
# ----------------------------------------------------------------------------
def rec_to_engine(rec: dict) -> dict:
    cj = rec.get("checklist_json") or {}
    inv = normalize_inventory(rec.get("inventory_data"))
    recs = None
    if inv:
        headers = inv["headers"]
        recs = [{h: row.get(h) for h in headers} for row in inv["rows"]]
    files = {}
    for cat, lst in (cj.get("files") or {}).items():
        if lst:
            files[cat] = [{"id": f.get("id"), "name": f.get("filename", ""), "size": f.get("size", 0), "meta": f}
                          for f in lst]
    return {
        "checklist": cj.get("checkboxes") or {}, "notes": cj.get("notes") or {},
        "cooh_text": rec.get("cooh_text", ""), "cooh_date": rec.get("cooh_date", ""),
        "chargebacks_text": rec.get("chargebacks_text", ""), "expenses": cj.get("expenses") or {},
        "inventory_data": recs, "inventory_filename": rec.get("inventory_filename", ""), "files": files,
    }


_IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".heic", ".heif", ".bmp", ".tif", ".tiff")


def gather_documents(payload: dict, light: bool = False):
    """light=True is used for the page: photos and receipts are not downloaded (they are only attached)."""
    docs, unreadable = [], []
    for cat, lst in (payload.get("files") or {}).items():
        for f in lst:
            if light and (cat.startswith(("rcpt", "receipts")) or str(f.get("name", "")).lower().endswith(_IMG_EXT)):
                continue
            data = read_file_bytes(f["meta"])
            label = DOC_LABELS.get(cat, cat)
            if data is None:
                unreadable.append((label, f["name"], _READ_ERR.get(f.get("id"), "file not found in storage")))
            else:
                docs.append({"key": f"{cat}:{f['id']}", "label": label, "name": f["name"], "data": data,
                             "auto": bool(f["meta"].get("auto"))})
    return docs, unreadable


@st.cache_data(show_spinner=False, max_entries=64)
def parse_table(data: bytes, name: str):
    try:
        df = read_table(name, data)
        return _coerce_numeric(df.copy()) if df is not None else None
    except Exception:
        log.exception("Could not parse %s", name)
        return None


@st.cache_data(show_spinner=False, max_entries=64)
def analyze_cached(cat: str, name: str, data: bytes):
    return analyze_doc(cat, name, data, parse_table(data, name))


def compute_facts(docs: list, want_tables: bool = True):
    """Read every uploaded file. -> (sheet tables, generic table totals, one info dict per file)."""
    tables, facts, infos = {}, [], []
    for d in docs:
        if d.get("auto") or d["key"].startswith(("receipts:", "rcpt")):
            continue
        cat = d["key"].split(":")[0]
        info = dict(analyze_cached(cat, d["name"], d["data"]))
        info["key"], info["label"] = d["key"], d["label"]
        infos.append(info)
        if want_tables and info["kind"] != "ocs_invoice":
            df = parse_table(d["data"], d["name"])
            if df is not None and not df.empty:
                tables[d["key"]] = (d["label"], df)
        if info["kind"] == "table_total":
            facts.append({"cat": cat, "file": d["name"], **info["got"]})
    return tables, facts, infos


def describe_docs(infos: list) -> list:
    """One line per uploaded file for the page (OCS invoices are folded into one line)."""
    out, ocs = [], [i for i in infos if i["kind"] == "ocs_invoice"]
    for i in infos:
        if i["kind"] == "ocs_invoice":
            continue
        out.append({"label": i["label"], "file": i["file"], "status": i["status"], "lines": i["lines"],
                    "rows": i.get("rows", 0), "headers": i.get("headers", []), "cols": i.get("cols", [])})
    if ocs:
        sub = sum(i["inv"]["subtotal"] + i["inv"]["freight"] for i in ocs)
        out.append({"label": "OCS: Invoices", "file": f"{len(ocs)} invoice files", "status": "read", "rows": 0, "headers": [], "cols": [],
                    "lines": [["Product + freight", sub], ["HST", sum(i["inv"]["tax"] for i in ocs)],
                              ["Invoice total", sum(i["inv"]["total"] for i in ocs)]]})
    return out


def _brief(info: dict) -> str:
    parts = []
    for l in info["lines"][:4]:
        v = l[1]
        parts.append(f"{l[0]}: {v}" if isinstance(v, str) else f"{l[0]}: {v:,.2f}")
    return "Read - " + "; ".join(parts)


def package_entries(store: str, period: str, payload: dict):
    docs, unreadable = gather_documents(payload)
    tables, facts, infos = compute_facts(docs)
    payload["doc_facts"], payload["docs_info"] = facts, infos
    by_key = {i["key"]: i for i in infos}
    doc_log = []
    for d in docs:
        kb = round(len(d["data"]) / 1024, 1)
        info = by_key.get(d["key"])
        if d.get("auto"):
            note = "Auto-filed formatted inventory (see Inventory sheet)"
        elif info and info["status"] == "read" and d["key"] not in tables:
            note = _brief(info) + " (see Income Statement)"
        elif d["key"] in tables:
            note = f"Own sheet ({len(tables[d['key']][1])} rows)" + ("; " + _brief(info) if info and info["status"] == "read" else "")
        else:
            note = "Attached only (photo, or numbers could not be read)"
        doc_log.append([d["label"], d["name"], kb, note])
    for label, name, *_e in unreadable:
        doc_log.append([label, name, "", "NOT FOUND in storage - please upload again"])
    xlsx = build_store_workbook(store, period, payload, tables, doc_log)
    owner_html = build_owner_report_html(store, period, payload)
    inv_rows = payload.get("inventory_data")
    base = safe_name(f"{store} - {period}")
    entries, seen = [(f"{base}/{base} - Report.xlsx", xlsx),
                     (f"{base}/{base} - Owner Report.html", owner_html.encode("utf-8"))], set()
    for d in docs:
        arc = f"{base}/Documents/{safe_name(d['label'])} - {safe_name(d['name'])}"
        n = 2
        while arc in seen:
            arc = f"{base}/Documents/{n} - {safe_name(d['label'])} - {safe_name(d['name'])}"
            n += 1
        seen.add(arc)
        entries.append((arc, d["data"]))
    return xlsx, entries, unreadable


def build_store_package(store: str, month: str, year: str):
    spk = f"{store}_{month}_{year}"
    rec = get_record(spk)
    if not rec:
        return None
    xlsx, entries, unreadable = package_entries(store, f"{month} {year}", rec_to_engine(rec))
    return {"spk": spk, "xlsx": xlsx, "zip": build_zip(entries), "n_docs": len(entries) - 1,
            "unreadable": unreadable, "at": datetime.now().strftime("%H:%M:%S")}


def build_owner_report(store: str, month: str, year: str):
    rec = get_record(f"{store}_{month}_{year}")
    if not rec:
        return None
    payload = rec_to_engine(rec)
    docs, _bad = gather_documents(payload)
    _t, payload["doc_facts"], payload["docs_info"] = compute_facts(docs)
    return build_owner_report_html(store, f"{month} {year}", payload)


def build_all_package(month: str, year: str):
    period, payloads, entries = f"{month} {year}", {}, []
    for s in STORES:
        rec = get_record(f"{s}_{month}_{year}")
        if rec:
            payloads[s] = rec_to_engine(rec)
            entries += package_entries(s, period, payloads[s])[1]
    if not payloads:
        return None
    summary = build_consolidated_workbook(period, payloads)
    entries.insert(0, (f"All Stores - {period} - Summary.xlsx", summary))
    return {"key": f"{month}_{year}", "xlsx": summary, "zip": build_zip(entries), "stores": len(payloads)}


# ----------------------------------------------------------------------------
# Login tokens (the page keeps you signed in through the URL; this makes it unforgeable)
# ----------------------------------------------------------------------------
@st.cache_resource
def _auth_key() -> bytes:
    return str(secret("AUTH_SECRET") or pysecrets.token_hex(32)).encode()


def make_token(username: str, password: str) -> str:
    sig = hmac.new(_auth_key(), f"{username}:{password}".encode(), hashlib.sha256).hexdigest()[:32]
    return f"{username}~{sig}"


def user_from_token(token: str):
    """-> (username, name, role) or None"""
    if not token or "~" not in token:
        return None
    username, _, sig = token.partition("~")
    rec = get_user(username)
    if not rec:
        return None
    if hmac.compare_digest(make_token(username.lower().strip(), rec[2]), f"{username.lower().strip()}~{sig}"):
        return username.lower().strip(), rec[0], rec[1]
    return None


# ----------------------------------------------------------------------------
# Boot
# ----------------------------------------------------------------------------
def apply_global_css():
    st.markdown("<style>.stApp,[data-testid='stAppViewContainer']{background:#f4f4f5 !important}"
                "[data-stale='true'],[data-stale='true'] *{opacity:1 !important;filter:none !important;transition:none !important}"
                "[data-testid='stStatusWidget'],[data-testid='stDecoration'],[data-testid='stAppDeployButton']{display:none !important}"
                "[data-testid='stHeader']{background:transparent !important}"
                "iframe{background:#f4f4f5}</style>", unsafe_allow_html=True)


def lock_page_scroll():
    """Once signed in, only the app window scrolls: every layer around it is locked."""
    st.markdown("<style>html,body,.stApp,[data-testid='stApp'],[data-testid='stAppViewContainer'],[data-testid='stMain'],"
                "section.main,[data-testid='stMainBlockContainer'],.block-container{overflow:hidden !important}</style>",
                unsafe_allow_html=True)


init_db()
restore_missing_from_dropbox()
apply_global_css()

for _k, _v in {"logged_in": False, "username": None, "name": None, "role": None}.items():
    st.session_state.setdefault(_k, _v)

query_params = st.query_params


def _log_in(username, name, role):
    st.session_state.logged_in = True
    st.session_state.username = username
    st.session_state.name = name
    st.session_state.role = role


def _log_out():
    for k in ("logged_in", "username", "name", "role", "ctx", "pkg", "all_pkg", "owner"):
        st.session_state.pop(k, None)


if not st.session_state.logged_in and "auth_user" in query_params:
    _u = user_from_token(query_params.get("auth_user", ""))
    if _u:
        _log_in(*_u)

# ----------------------------------------------------------------------------
# Login screen
# ----------------------------------------------------------------------------
if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background-color: #065f46; padding: 30px; border-radius: 16px; color: white; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
                <div style="font-size: 40px; margin-bottom: 10px;">🌿</div>
                <h2 style="margin: 0; font-size: 22px;">The OZ Store Portal</h2>
                <p style="font-size: 12px; color: #a7f3d0; margin-top: 5px;">Secure Persistent Database Manager</p>
            </div>
        """,
            unsafe_allow_html=True,
        )

        st.caption(f"App version {APP_VERSION}")
        with st.form("login_form"):
            username_input = st.text_input("Username")
            password_input = st.text_input("Password", type="password")
            submit = st.form_submit_button("Sign In", **STRETCH)

            if submit:
                user_record = get_user(username_input.strip())
                if user_record and hmac.compare_digest(str(user_record[2]).encode(), password_input.strip().encode()):
                    clean_u = username_input.lower().strip()
                    _log_in(clean_u, user_record[0], user_record[1])
                    st.query_params.clear()
                    st.query_params["auth_user"] = make_token(clean_u, user_record[2])
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
    st.stop()

role = st.session_state.role
is_staff = role in ("admin", "accountant")


# ----------------------------------------------------------------------------
# Which store / month / year is on screen (set by the page's tabs and period menu)
# ----------------------------------------------------------------------------
def allowed_stores():
    return [st.session_state.name] if role == "store" else STORES


def init_ctx():
    today = date.today()
    ctx = st.session_state.get("ctx") or {
        "store": None, "month": MONTHS[today.month - 1],
        "year": str(today.year) if str(today.year) in YEARS else YEARS[1],
    }
    allowed = allowed_stores()
    s, m, y = query_params.get("store"), query_params.get("month"), query_params.get("year")
    if s in allowed:
        ctx["store"] = s
    if m in MONTHS:
        ctx["month"] = m
    if y in YEARS:
        ctx["year"] = y
    if ctx["store"] not in allowed:
        ctx["store"] = allowed[0]
    st.session_state.ctx = ctx
    return ctx


ctx = init_ctx()
spk = f"{ctx['store']}_{ctx['month']}_{ctx['year']}"


def current_token():
    rec = get_user(st.session_state.username)
    return make_token(st.session_state.username, rec[2]) if rec else None


def reset_url_and_rerun():
    """Drop the action from the URL but keep you signed in on the same store/period."""
    tok = current_token()
    if not tok:
        _log_out()
        st.query_params.clear()
        st.rerun()
    c = st.session_state.ctx
    params = {"auth_user": tok, "store": c["store"], "month": c["month"], "year": c["year"]}
    st.query_params.clear()
    if hasattr(st.query_params, "from_dict"):
        st.query_params.from_dict(params)
    else:
        for k, v in params.items():
            st.query_params[k] = v
    st.rerun()


def flash(kind: str, text: str):
    st.session_state.flash = (kind, text)


# ----------------------------------------------------------------------------
# Actions sent by the page (these arrive as URL parameters)
# ----------------------------------------------------------------------------
def sync_auto_inventory(target: str):
    """File the formatted inventory report (Excel) in Greenline: Inventory Summary, replacing the last auto one."""
    rec = get_record(target)
    if not rec:
        return
    files = rec["checklist_json"].setdefault("files", {})
    g1 = files.get("green1", [])
    for f in [f for f in g1 if f.get("auto")]:
        delete_file_bytes(f)
    g1 = [f for f in g1 if not f.get("auto")]
    inv = rec.get("inventory_data")
    if inv and inv.get("rows"):
        df = _inventory_frame({"inventory_data": [{h: r.get(h) for h in inv["headers"]} for r in inv["rows"]]})
        store, _, rest = target.partition("_")
        meta, _warn = store_file(target, "green1", f"{store}_{rest}_Formatted_Inventory.xlsx",
                                 inventory_workbook(df, f"{store} Inventory"))
        meta["auto"] = True
        g1.append(meta)
    files["green1"] = g1
    save_record(target, rec)


def act_save_all():
    raw = query_params.get("payload", "")
    if not raw or len(raw) > 400_000:
        return flash("error", "Nothing was saved: the data was empty or too large.")
    try:
        payload = json.loads(raw)
    except ValueError:
        try:
            payload = json.loads(urllib.parse.unquote(raw))
        except ValueError:
            return flash("error", "Nothing was saved: the data could not be read.")
    store = payload.get("store", "")
    month, _, year = str(payload.get("period", "")).partition(" ")
    if store not in allowed_stores() or month not in MONTHS or year not in YEARS:
        return flash("error", "Nothing was saved: that store or period isn't allowed for your account.")
    target = f"{store}_{month}_{year}"
    old = get_record(target) or blank_record()
    cj = payload.get("checklist_json") if isinstance(payload.get("checklist_json"), dict) else {}
    cj["files"] = old["checklist_json"].get("files", {})  # files are managed on the server only
    rec = normalize_remote({
        "cooh_text": payload.get("cooh_text", ""), "cooh_date": payload.get("cooh_date", ""),
        "chargebacks_text": payload.get("chargebacks_text", "Nothing to report."), "checklist_json": cj,
        "inventory_data": payload.get("inventory_data") or payload.get("inventory_json"), "inventory_filename": payload.get("inventory_filename", ""),
        "_updated": payload.get("_updated") or 0,
    })
    save_record(target, rec)
    try:
        sync_auto_inventory(target)
    except Exception:
        log.exception("Could not file the formatted inventory report")
    ok, err = sync_all_to_dropbox()
    if query_params.get("then") == "package":
        st.session_state.auto_pkg = True
    elif query_params.get("then") == "excel":
        st.session_state.auto_pkg = True
        st.session_state.dl_which = "xlsx"
    elif query_params.get("then") == "owner":
        st.session_state.scroll_owner = True
    if ok:
        flash("success", f"Saved {store} - {month} {year} and synced to Dropbox.")
    else:
        flash("warning", f"Saved on the server, but Dropbox sync failed: {err}")


def act_update_user():
    if role != "admin":
        return flash("error", "Only the admin can change users.")
    orig = query_params.get("orig_username", "").lower().strip()
    user = query_params.get("username", "").lower().strip()
    name = query_params.get("name", "").strip()
    new_role = query_params.get("role", "")
    password = query_params.get("password", "")
    if not re.fullmatch(r"[a-z0-9_.-]{2,30}", user) or not name or len(name) > 60 or not password or len(password) > 100:
        return flash("error", "User not saved: check the username (letters/numbers), display name and password.")
    if new_role not in ROLES:
        return flash("error", "User not saved: unknown role.")
    if new_role == "store" and name not in STORES:
        return flash("error", "User not saved: a store login's display name must be exactly one of " + ", ".join(STORES) + ".")
    with _connect() as conn:
        if orig and orig != user:
            conn.execute("DELETE FROM users WHERE username = ?", (orig,))
        conn.execute("INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?)", (user, name, new_role, password))
        conn.commit()
    if orig and orig == st.session_state.username and orig != user:
        st.session_state.username = user
    sync_users_to_dropbox()
    flash("success", f"User {user} saved.")


def act_delete_user():
    if role != "admin":
        return flash("error", "Only the admin can change users.")
    user = query_params.get("username", "").lower().strip()
    if user == st.session_state.username:
        return flash("error", "You can't delete the account you're signed in with.")
    with _connect() as conn:
        conn.execute("DELETE FROM users WHERE username = ?", (user,))
        conn.commit()
    sync_users_to_dropbox()
    flash("success", f"User {user} deleted.")


def act_delete_file():
    cat, fid = query_params.get("cat", ""), query_params.get("file_id", "")
    rec = get_record(spk)
    files = ((rec or {}).get("checklist_json") or {}).get("files", {})
    keep, gone = [], []
    for f in files.get(cat, []):
        (gone if f.get("id") == fid else keep).append(f)
    if not gone:
        return flash("warning", "That file was already removed.")
    for f in gone:
        delete_file_bytes(f)
    files[cat] = keep
    save_record(spk, rec)
    sync_all_to_dropbox()
    flash("success", f"Removed {gone[0].get('filename', 'file')}.")


def act_clear_files():
    cat = query_params.get("cat", "")
    rec = get_record(spk)
    files = ((rec or {}).get("checklist_json") or {}).get("files", {})
    for f in files.get(cat, []):
        delete_file_bytes(f)
    if files.get(cat):
        files[cat] = []
        save_record(spk, rec)
        sync_all_to_dropbox()
    flash("success", f"Cleared all files in {DOC_LABELS.get(cat, cat)}.")


def handle_action(action: str):
    if action == "sign_out":
        _log_out()
        st.query_params.clear()
        st.rerun()
    elif action == "save_all":
        act_save_all()
    elif action == "update_user":
        act_update_user()
    elif action == "delete_user":
        act_delete_user()
    elif action == "delete_file":
        act_delete_file()
    elif action == "clear_files":
        act_clear_files()
    elif action == "upload_panel":
        cat = query_params.get("cat", "")
        pass
    elif action == "package":
        st.session_state.auto_pkg = True
    reset_url_and_rerun()


if "action" in query_params:
    handle_action(query_params.get("action", ""))

# ----------------------------------------------------------------------------
# Documents & reports panel (real uploads + the report generator)
# ----------------------------------------------------------------------------
def process_hidden_uploads(chosen):
    """Files chosen with the page's own Browse buttons arrive here (the uploaders are invisible)."""
    ver = st.session_state.setdefault("upl_ver", 0)
    for cat, _label in UPLOAD_CATS:
        files = chosen.get(cat)
        if not files:
            continue
        ok_n, problems, metas = 0, [], []
        free_slots = []
        if cat == "receipts":
            _have = ((get_record(spk) or {}).get("checklist_json") or {}).get("files", {})
            free_slots = [f"rcpt{n}" for n in range(1, 51) if not _have.get(f"rcpt{n}")]
        by_cat = {}
        for f in files:
            if f.size > MAX_UPLOAD_MB * 1024 * 1024:
                problems.append(f"{f.name} is over {MAX_UPLOAD_MB} MB.")
                continue
            target = cat
            if cat == "receipts":
                if not free_slots:
                    problems.append(f"{f.name}: all 50 receipt slots are full.")
                    continue
                target = free_slots.pop(0)
            meta, warn = store_file(spk, target, f.name, f.getvalue())
            metas.append(meta)
            by_cat.setdefault(target, []).append(meta)
            ok_n += 1
            if warn:
                problems.append(warn)
        if metas:
            for _t, _m in by_cat.items():
                add_files_to_record(spk, _t, _m)
            if read_file_bytes(metas[0]) is None:
                problems.append("The file was saved but could not be read back from storage: "
                                + _READ_ERR.get(metas[0].get("id"), "unknown error"))
            synced, err = sync_all_to_dropbox()
            if not synced:
                problems.append(f"Dropbox sync failed: {err}")
        st.session_state.upl_ver = ver + 1
        if problems:
            flash("warning", f"{ok_n} file(s) saved to {DOC_LABELS[cat]}. " + " ".join(dict.fromkeys(problems)))
        else:
            flash("success", f"{ok_n} file(s) saved to {DOC_LABELS[cat]}. Tick its box in the checklist and press Save All.")
        st.rerun()


def render_hidden_tools():
    """Invisible controls driven by the page's buttons. Nothing here is ever shown on screen."""
    store, month, year = ctx["store"], ctx["month"], ctx["year"]
    ver = st.session_state.setdefault("upl_ver", 0)

    st.markdown("<style>.st-key-oz_hidden_tools{position:fixed !important;left:-10000px !important;top:0 !important;"
                "width:1px !important;height:1px !important;overflow:hidden !important;opacity:0 !important;pointer-events:none !important}"
                "</style>", unsafe_allow_html=True)
    with st.container(key="oz_hidden_tools"):
        chosen = {}
        for cat, _label in UPLOAD_CATS:
            if cat.startswith("rcpt"):
                continue
            with st.container(key=f"up_{cat}"):
                chosen[cat] = st.file_uploader(DOC_LABELS[cat], accept_multiple_files=True, key=f"upl_{cat}_{ver}",
                                               label_visibility="collapsed")
        process_hidden_uploads(chosen)
        if st.session_state.pop("auto_pkg", False):
            pkg = build_store_package(store, month, year)
            if pkg is not None:
                st.session_state.pkg = pkg
                st.session_state.dl_now = True
            else:
                flash("info", "Nothing saved yet for this store and month. Press Save All to Cloud first.")
        pkg = st.session_state.get("pkg")
        if pkg and pkg["spk"] == spk:
            base = safe_name(f"{store} - {month} {year}")
            st.download_button("zip", pkg["zip"], file_name=f"{base} - Package.zip", mime="application/zip", key="dl_zip")
            st.download_button("xlsx", pkg["xlsx"], file_name=f"{base} - Detailed Accounting Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
            if pkg.get("unreadable"):
                names = ", ".join(n for _l, n, *_e in pkg["unreadable"])
                st.session_state.flash_keep = ("warning", f"Could not read from storage: {names}. Please upload again.")
    js = ""
    if st.session_state.pop("dl_now", False):
        _which = "dl_xlsx" if st.session_state.pop("dl_which", "zip") == "xlsx" else "dl_zip"
        js += ("setTimeout(function(){try{var b=window.parent.document.querySelector('.st-key-" + _which + " button');if(b)b.click();}catch(e){}},700);")
    if st.session_state.pop("scroll_owner", False):
        js += ("function go(){try{var f=window.parent.document.querySelectorAll('iframe');if(f.length>1)"
               "f[f.length-1].scrollIntoView({behavior:'smooth',block:'start'});}catch(e){}}setTimeout(go,900);setTimeout(go,2000);")
    if js:
        components.html("<script>" + js + "</script>", height=0)



# ----------------------------------------------------------------------------
# Build the page
# ----------------------------------------------------------------------------
def build_html_page() -> str:
    rec = get_record(spk)
    files_json = {}
    store_data_json = {}
    if rec:
        cj = rec["checklist_json"]
        files_json[spk] = {cat: [{"id": f.get("id"), "filename": f.get("filename"), "size": f.get("size", 0)}
                                 for f in lst] for cat, lst in (cj.get("files") or {}).items() if lst}
        store_data_json[spk] = rec

    page_facts = {"sales": 0, "sales_src": "", "cogs": 0, "cogs_src": "", "extra": [], "docs": []}
    if rec:
        sig = "v2:" + ",".join(sorted(str(f.get("id")) for lst in ((rec["checklist_json"].get("files")) or {}).values() for f in lst))
        cached_facts = None
        try:
            with _connect() as conn:
                row = conn.execute("SELECT v FROM app_meta WHERE k = ?", (f"facts:{spk}",)).fetchone()
            if row:
                c = json.loads(row[0])
                if c.get("sig") == sig:
                    cached_facts = c.get("facts")
        except Exception:
            log.exception("Could not read stored document totals")
        if cached_facts:
            page_facts = cached_facts
        else:
            try:
                payload = rec_to_engine(rec)
                docs, _bad = gather_documents(payload, light=True)
                _t, payload["doc_facts"], payload["docs_info"] = compute_facts(docs, want_tables=False)
                inc = build_income(payload)
                from_greenline = inc["cogs"]["src"].startswith("Greenline")
                page_facts = {
                    "sales": inc["revenue"]["net"] if inc["revenue"]["from_docs"] else 0, "sales_src": inc["revenue"]["src"],
                    "cogs": inc["cogs"]["amount"] if from_greenline else 0, "cogs_src": inc["cogs"]["src"],
                    "extra": [{"label": x["label"], "total": x["amount"]} for x in inc["extra_expenses"]],
                    "docs": describe_docs(payload["docs_info"]) + [
                        {"label": lab, "file": nm, "status": "unreadable", "lines": [], "rows": 0, "headers": [], "cols": [], "error": err}
                        for lab, nm, err in _bad],
                }
                if not _bad:
                    try:
                        with _connect() as conn:
                            conn.execute("INSERT OR REPLACE INTO app_meta VALUES (?, ?)",
                                         (f"facts:{spk}", json.dumps({"sig": sig, "facts": page_facts}, default=str)))
                    except Exception:
                        log.exception("Could not store document totals")
            except Exception as ex:
                log.exception("Could not read totals from the uploaded documents")
                page_facts["error"] = f"{type(ex).__name__}: {ex}"

    cred_rows = ""
    if role == "admin":
        with _connect() as conn:
            users = conn.execute("SELECT username, name, role, password FROM users ORDER BY role, username").fetchall()
        for u, n, r, p in users:
            args = htmllib.escape(json.dumps([u, n, r, p]), quote=True)
            cred_rows += f"""
        <tr>
            <td class="p-3 font-mono font-medium text-emerald-800">{htmllib.escape(u)}</td>
            <td class="p-3">{htmllib.escape(n)}</td>
            <td class="p-3 uppercase text-[10px] font-bold"><span class="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded">{htmllib.escape(r)}</span></td>
            <td class="p-3 font-mono">{htmllib.escape(p)}</td>
            <td class="p-3 text-right">
                <button onclick="editUser(...{args})" class="text-emerald-700 font-bold underline mr-2">Edit</button>
                <button onclick="deleteUser({htmllib.escape(json.dumps(u), quote=True)})" class="text-red-600 font-bold underline">Delete</button>
            </td>
        </tr>
    """

    add_user_btn = (
        '<button onclick="openAddUserModal()" class="bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-xs">+ Add New User</button>'
        if role == "admin" else ""
    )
    e = lambda s: htmllib.escape(str(s), quote=True)
    return (
        HTML_CODE.replace("REPLACE_CURRENT_NAME", e(st.session_state.name))
        .replace("REPLACE_CURRENT_USERNAME", e(current_token() or ""))
        .replace("REPLACE_CURRENT_ROLE", e(role))
        .replace("REPLACE_ADMIN_HIDDEN_CLASS", "" if role == "admin" else "hidden")
        .replace("REPLACE_CRED_TABLE_ROWS", cred_rows)
        .replace("REPLACE_ACCOUNTANT_HIDDEN_CLASS", "hidden" if role == "accountant" else "")
        .replace("REPLACE_ADD_USER_BTN_HTML", add_user_btn)
        .replace("REPLACE_IS_STORE_MANAGER", "true" if role == "store" else "false")
        .replace("REPLACE_ACTIVE_STORE", e(ctx["store"]))
        .replace("REPLACE_MONTH", e(ctx["month"]))
        .replace("REPLACE_YEAR", e(ctx["year"]))
        .replace("REPLACE_DOC_FACTS", json.dumps(page_facts).replace("</", "<\\/"))
        .replace("REPLACE_FILES_JSON", json.dumps(files_json).replace("</", "<\\/"))
        .replace("REPLACE_STORE_DATA_JSON", json.dumps(store_data_json, default=str).replace("</", "<\\/"))
    )


# Make sure the sign-in token is in the URL (so a refresh keeps you signed in)
_tok = current_token()
if _tok and query_params.get("auth_user") != _tok:
    st.query_params["auth_user"] = _tok

# Nothing is drawn above the page any more: messages are toasts, and the upload panel sits at the very bottom.
_flash = st.session_state.pop("flash", None)
if _flash:
    try:
        st.toast(_flash[1], icon={"success": "✅", "warning": "⚠️", "error": "❌"}.get(_flash[0], "ℹ️"))
    except Exception:
        pass
    if _flash[0] in ("warning", "error"):
        st.session_state.flash_keep = _flash
    else:
        st.session_state.pop("flash_keep", None)

lock_page_scroll()
render_hidden_tools()
components.html(build_html_page(), height=900, scrolling=False)

_keep = st.session_state.get("flash_keep")
if _keep:
    getattr(st, _keep[0], st.info)(_keep[1])
st.caption(f"App version {APP_VERSION}")
