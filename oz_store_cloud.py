"""The OZ Store - Cloud Accountant Package Manager (single file)."""
from __future__ import annotations

import io
import re
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
    ("rent", "Store Rent & Lease"), ("salaries", "Employee Salaries"),
    ("utilities", "Utilities"), ("advertising", "Advertising"), ("marketing", "Marketing"),
    ("datadeals", "Data Deals & Analytics"), ("insurance", "Insurance"),
    ("garbage", "Garbage Disposal"), ("security", "Security & POS"),
    ("licenses", "Licenses & Compliance"),
]

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
    return sum(parse_money(v) for v in (payload.get("expenses") or {}).values())


def missing_docs(payload: dict) -> list[str]:
    chk = payload.get("checklist") or {}
    return [label for key, label in CHECKLIST_ITEMS if not chk.get(key)]


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
    return pd.DataFrame(recs) if recs else None


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
    store: str, period: str, payload: dict, tables: dict[str, tuple[str, pd.DataFrame]]
) -> bytes:
    """One store, one month: Summary + Inventory + one sheet per tabular document."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    used = {"summary"}

    ws["A1"] = f"The OZ Store - {store}"
    ws["A1"].font = Font(name=_FONT, bold=True, size=16, color="065F46")
    ws["A2"] = f"Period: {period}"
    ws["A3"] = f"Generated: {datetime.now():%Y-%m-%d %H:%M}"
    for a in ("A2", "A3"):
        ws[a].font = _BODY
    missing = missing_docs(payload)
    ws["A4"] = ("Missing documents: " + ", ".join(missing)) if missing else "All documents received"
    ws["A4"].font = Font(name=_FONT, bold=True, size=10, color="B91C1C" if missing else "047857")

    files = payload.get("files") or {}
    notes = payload.get("notes") or {}
    chk = payload.get("checklist") or {}
    rows = [
        [label, "Received" if chk.get(key) else "MISSING", (files.get(key) or {}).get("name", ""), notes.get(key, "")]
        for key, label in CHECKLIST_ITEMS
    ]
    r = _block(ws, 6, "Document checklist", ["Document", "Status", "File", "Note"], rows,
               bad_when=lambda rec: rec[1] == "MISSING")

    cooh_text = payload.get("cooh_text", "")
    cooh_amt = parse_money(cooh_text) if str(cooh_text).strip() else None
    r = _block(ws, r, "Cash & chargebacks", ["Item", "Value"], [
        ["Cash on hand", cooh_amt if cooh_amt is not None else "Not entered"],
        ["Cash on hand date", payload.get("cooh_date") or "Not entered"],
        ["Moneris chargebacks", payload.get("chargebacks_text") or "Nothing to report."],
    ])

    exp = payload.get("expenses") or {}
    exp_rows = [[label, parse_money(exp.get(key))] for key, label in EXPENSE_FIELDS]
    exp_rows.append(["TOTAL", total_expenses(payload)])
    r = _block(ws, r, "Operating expenses", ["Expense", "Amount"], exp_rows)

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
        sheet = wb.create_sheet(_sheet_name(label, used))
        _write_df(sheet, df, totals=True)
        _autowidth(sheet)

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
        rows.append([
            store,
            f"{checklist_done(p)}/{len(CHECKLIST_ITEMS)}",
            parse_money(cooh) if cooh else "Not entered",
            p.get("cooh_date") or "",
            p.get("chargebacks_text") or "",
            total_expenses(p),
            ", ".join(missing_docs(p)) or "-",
        ])
    _block(ws, 4, "Store overview",
           ["Store", "Checklist", "Cash on hand", "Cash date", "Chargebacks", "Total expenses", "Missing"],
           rows, bad_when=lambda rec: rec[6] != "-")
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
    tr = len(EXPENSE_FIELDS) + 2
    _put(wx, tr, 1, "TOTAL", bold=True, fill=_TOT_FILL)
    for j, s in enumerate(stores, 2):
        _put(wx, tr, j, total_expenses(payloads[s]), bold=True, fill=_TOT_FILL)
    _put(wx, tr, len(stores) + 2, sum(total_expenses(payloads[s]) for s in stores), bold=True, fill=_TOT_FILL)
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
# STREAMLIT APP
# ==========================================================================
import hashlib
import hmac
import inspect
import json
import logging
import os
import re
import secrets as pysecrets
import sqlite3
import tempfile
from contextlib import closing
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import dropbox
    from dropbox.exceptions import ApiError
    from dropbox.files import WriteMode
except ImportError:  # app still works (local only) without the package
    dropbox = None
    ApiError = Exception
    WriteMode = None

log = logging.getLogger("oz_store")


def _stretch():
    try:
        if "width" in inspect.signature(st.button).parameters:
            return {"width": "stretch"}
    except Exception:
        pass
    return {"use_container_width": True}


STRETCH = _stretch()

st.set_page_config(page_title="The OZ Store - Cloud Accountant Package Manager", page_icon="🌿", layout="wide")
st.markdown(
    """
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    h1, h2, h3 { color: #065f46; }
    .stButton>button, .stDownloadButton>button { background-color: #047857; color: white; border-radius: 12px; font-weight: bold; border: none; }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #065f46; color: white; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
DROPBOX_FOLDER = "/OZ_Store_Uploads"
MASTER_PATH = f"{DROPBOX_FOLDER}/master_store_state.json"
DB_FILE = "oz_store_v2.db"      # fresh file: ignores any old-layout oz_store.db
LEGACY_DB = "oz_store.db"
LOCAL_FILES = Path(tempfile.gettempdir()) / "oz_uploads"
MAX_UPLOAD_MB = 25
DOC_TYPES = ["csv", "xlsx", "xls", "pdf", "png", "jpg", "jpeg"]
STORES = [f"Oz {i}" for i in range(1, 8)]
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
DOC_LABELS = dict(CHECKLIST_ITEMS)
DOC_LABELS["inventory"] = "Inventory CSV (raw)"

DEFAULT_USERS = {
    "admin": ("Master Admin", "admin", "admin2026"),
    "leanne": ("Leanne (Accountant)", "accountant", "leanne2026"),
    **{f"oz{i}": (f"Oz {i}", "store", f"oz{i}pass2026") for i in range(1, 8)},
}


def secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# --------------------------------------------------------------------------
# Passwords (PBKDF2). Old plaintext rows keep working and are upgraded at login.
# --------------------------------------------------------------------------
def hash_pw(pw: str, salt: str | None = None) -> str:
    salt = salt or pysecrets.token_hex(8)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 100_000)
    return f"pbkdf2${salt}${dk.hex()}"


def check_pw(pw: str, stored: str) -> bool:
    if stored and stored.startswith("pbkdf2$"):
        return hmac.compare_digest(hash_pw(pw, stored.split("$")[1]), stored)
    return hmac.compare_digest(pw.encode(), (stored or "").encode())


# --------------------------------------------------------------------------
# Local SQLite (users + local cache of the data; Dropbox is the source of truth)
# --------------------------------------------------------------------------
LEGACY_COLS = ["cooh_text", "cooh_date", "chargebacks_text", "checklist_json",
               "inventory_json", "inventory_filename"]


def _conn():
    return closing(sqlite3.connect(DB_FILE))


@st.cache_resource
def init_db():
    """Create tables, add any column an older version of the app didn't have, seed users."""
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT, password TEXT)")
        cur.execute("CREATE TABLE IF NOT EXISTS store_data (store_period_key TEXT PRIMARY KEY, store TEXT, period TEXT, data_json TEXT)")
        cur.execute("PRAGMA table_info(store_data)")
        existing = {c[1] for c in cur.fetchall()}
        for col in ["store", "period", "data_json"] + LEGACY_COLS:
            if col not in existing:
                cur.execute(f"ALTER TABLE store_data ADD COLUMN {col} TEXT")

        configured = secret("USERS")  # optional: survives reboots, see notes
        if configured:
            for u, info in configured.items():
                cur.execute("INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?)",
                            (u.lower(), info["name"], info["role"], hash_pw(info["password"])))
        else:
            have = {r[0] for r in cur.execute("SELECT username FROM users")}
            for u, (name, role, pw) in DEFAULT_USERS.items():
                if u not in have:
                    cur.execute("INSERT INTO users VALUES (?, ?, ?, ?)", (u, name, role, hash_pw(pw)))
        conn.commit()
        empty = conn.execute("SELECT 1 FROM store_data LIMIT 1").fetchone() is None
    if empty and os.path.exists(LEGACY_DB):  # one-time import from the old database file
        legacy = db_read_all(LEGACY_DB)
        if legacy:
            db_mirror(legacy)
    return True


def get_user(username: str):
    with _conn() as conn:
        return conn.execute("SELECT name, role, password FROM users WHERE username = ?",
                            (username.lower().strip(),)).fetchone()


def _legacy_row(r):
    chk = json.loads(r[5]) if r[5] else {}
    return {
        "cooh_text": r[2] or "", "cooh_date": r[3] or "",
        "chargebacks_text": r[4] or "Nothing to report.",
        "checklist": chk.get("checkboxes", {}), "notes": chk.get("notes", {}),
        "inventory_data": json.loads(r[6]) if r[6] else None,
        "inventory_filename": r[7] or "",
    }


def db_read_all(path: str = DB_FILE) -> dict:
    out = {}
    want = ["store_period_key", "data_json"] + LEGACY_COLS
    try:
        with closing(sqlite3.connect(path)) as conn:
            cols = {c[1] for c in conn.execute("PRAGMA table_info(store_data)")}
            if not cols:
                return out
            sel = ", ".join(c if c in cols else "NULL" for c in want)
            rows = conn.execute(f"SELECT {sel} FROM store_data").fetchall()
    except sqlite3.Error:
        log.exception("Could not read %s", path)
        return out
    for r in rows:
        try:
            out[r[0]] = json.loads(r[1]) if r[1] else _legacy_row(r)
        except Exception:
            log.exception("Skipping unreadable row %s", r[0])
    return out


def db_mirror(master: dict):
    rows = []
    for key, d in master.items():
        store, _, rest = key.partition("_")
        rows.append((key, store, rest.replace("_", " "), json.dumps(d, default=str)))
    with _conn() as conn:
        conn.executemany("INSERT OR REPLACE INTO store_data (store_period_key, store, period, data_json) VALUES (?, ?, ?, ?)", rows)
        conn.commit()


def db_save(key: str, store: str, period: str, payload: dict):
    db_mirror({key: payload})


# --------------------------------------------------------------------------
# Dropbox
# --------------------------------------------------------------------------
class SyncError(Exception):
    pass


@st.cache_resource
def get_dbx():
    """Prefers a refresh-token login (never expires). Falls back to a plain access token."""
    if dropbox is None:
        return None
    try:
        key, sec, rt = secret("DROPBOX_APP_KEY"), secret("DROPBOX_APP_SECRET"), secret("DROPBOX_REFRESH_TOKEN")
        if key and sec and rt:
            return dropbox.Dropbox(oauth2_refresh_token=rt, app_key=key, app_secret=sec)
        token = secret("DROPBOX_ACCESS_TOKEN")
        return dropbox.Dropbox(token) if token else None
    except Exception:
        log.exception("Dropbox init failed")
        return None


def _not_found(e) -> bool:
    try:
        return e.error.is_path() and e.error.get_path().is_not_found()
    except Exception:
        return False


def dbx_read_master():
    """dict, {} if the file doesn't exist yet, None if Dropbox isn't configured."""
    dbx = get_dbx()
    if dbx is None:
        return None
    try:
        _, res = dbx.files_download(MASTER_PATH)
        return json.loads(res.content)
    except ApiError as e:
        if _not_found(e):
            return {}
        raise SyncError(str(e)) from e
    except Exception as e:
        raise SyncError(str(e)) from e


def dbx_write_master(data: dict):
    dbx = get_dbx()
    if dbx is None:
        raise SyncError("Dropbox is not configured in Streamlit secrets.")
    try:
        dbx.files_upload(json.dumps(data, default=str).encode(), MASTER_PATH, mode=WriteMode.overwrite)
    except Exception as e:
        raise SyncError(str(e)) from e


@st.cache_data(ttl=60, show_spinner="Loading from cloud…")
def fetch_master():
    """(data, source, error). Dropbox first, local SQLite cache if Dropbox is unreachable/empty."""
    err = ""
    try:
        remote = dbx_read_master()
    except SyncError as e:
        remote, err = None, str(e)
    if remote:
        db_mirror(remote)
        return remote, "dropbox", ""
    return db_read_all(), "local", err


# --------------------------------------------------------------------------
# Uploaded documents (Dropbox, with a local fallback)
# --------------------------------------------------------------------------
def safe_name(n: str) -> str:
    return re.sub(r"[^\w.\- ]", "_", n)


def store_file(spk: str, key: str, name: str, data: bytes) -> dict:
    fname = f"{key}_{safe_name(name)}"
    meta = {"name": name, "size": len(data), "uploaded": datetime.now().isoformat(timespec="seconds")}
    dbx = get_dbx()
    if dbx:
        path = f"{DROPBOX_FOLDER}/files/{spk}/{fname}"
        try:
            dbx.files_upload(data, path, mode=WriteMode.overwrite)
            return {**meta, "backend": "dropbox", "path": path}
        except Exception:
            log.exception("Dropbox upload failed for %s", fname)
    p = LOCAL_FILES / spk / fname
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    return {**meta, "backend": "local", "path": str(p)}


@st.cache_data(ttl=600, max_entries=24, show_spinner=False)
def _dbx_download(path: str) -> bytes:
    _, res = get_dbx().files_download(path)
    return res.content


def load_file_bytes(meta: dict):
    try:
        if meta.get("backend") == "dropbox":
            return _dbx_download(meta["path"])
        p = Path(meta["path"])
        return p.read_bytes() if p.exists() else None
    except Exception:
        log.exception("Could not load %s", meta.get("name"))
        return None


@st.cache_data(show_spinner=False, max_entries=32)
def parse_inventory(data: bytes, name: str):
    df = read_table(name, data)
    if df is None:
        raise ValueError("Please upload a CSV or Excel file.")
    return clean_inventory(df)


@st.cache_data(show_spinner=False, max_entries=64)
def parse_table(data: bytes, name: str):
    try:
        return read_table(name, data)
    except Exception:
        log.exception("Could not parse %s", name)
        return None


@st.cache_data(show_spinner=False, max_entries=8)
def inv_xlsx(df: pd.DataFrame, title: str) -> bytes:
    return inventory_workbook(df, title)


def inv_records(df: pd.DataFrame):
    return df.astype(object).where(df.notna(), None).to_dict("records")


# --------------------------------------------------------------------------
# Package building
# --------------------------------------------------------------------------
def gather_documents(payload: dict, pend: dict | None = None) -> list[dict]:
    """Every uploaded document for a payload: saved ones from storage + unsaved ones in memory."""
    pend = pend or {}
    docs = []
    for key, meta in (payload.get("files") or {}).items():
        if key in pend:
            continue
        data = load_file_bytes(meta) if meta.get("path") else None
        if data:
            docs.append({"key": key, "label": DOC_LABELS.get(key, key), "name": meta["name"], "data": data})
    for key, (name, data) in pend.items():
        docs.append({"key": key, "label": DOC_LABELS.get(key, key), "name": name, "data": data})
    return docs


def package_entries(store: str, period: str, payload: dict, docs: list[dict]):
    tables = {}
    for d in docs:
        if d["key"] == "inventory":
            continue
        df = parse_table(d["data"], d["name"])
        if df is not None and not df.empty:
            tables[d["key"]] = (d["label"], df)
    xlsx = build_store_workbook(store, period, payload, tables)
    base = safe_name(f"{store} - {period}")
    entries = [(f"{base}/{base} - Report.xlsx", xlsx)]
    entries += [(f"{base}/Documents/{safe_name(d['label'])} - {safe_name(d['name'])}", d["data"]) for d in docs]
    return xlsx, entries


# --------------------------------------------------------------------------
# Session / login
# --------------------------------------------------------------------------
init_db()
for k, v in {"logged_in": False, "username": None, "name": None, "role": None,
             "pending": {}, "upver": 0, "default_pw": False}.items():
    st.session_state.setdefault(k, v)

if not st.session_state.logged_in:
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background-color:#065f46;padding:30px;border-radius:16px;color:white;text-align:center;box-shadow:0 10px 25px rgba(0,0,0,0.1);">
                <div style="font-size:40px;margin-bottom:10px;">🌿</div>
                <h2 style="margin:0;font-size:22px;color:white;">The OZ Store Portal</h2>
                <p style="font-size:12px;color:#a7f3d0;margin-top:5px;">Secure Persistent Database Manager</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            username_input = st.text_input("Username")
            password_input = st.text_input("Password", type="password")
            if st.form_submit_button("Sign In", **STRETCH):
                rec = get_user(username_input)
                if rec and check_pw(password_input, rec[2]):
                    uname = username_input.lower().strip()
                    if not rec[2].startswith("pbkdf2$"):  # upgrade old plaintext password
                        with _conn() as conn:
                            conn.execute("UPDATE users SET password = ? WHERE username = ?", (hash_pw(password_input), uname))
                            conn.commit()
                    st.session_state.update(
                        logged_in=True, username=uname, name=rec[0], role=rec[1],
                        default_pw=(uname in DEFAULT_USERS and DEFAULT_USERS[uname][2] == password_input),
                    )
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
    st.stop()

role = st.session_state.role
is_staff = role in ("admin", "accountant")

# --------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------
st.markdown(
    """
    <div style="background-color: #065f46; padding: 20px; border-radius: 16px; color: white; display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h2 style="margin:0; color: white;">🌿 The OZ Store</h2>
            <p style="margin:0; font-size: 12px; color: #a7f3d0;">Executive Financial & Accountant Package Manager</p>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

col_h1, col_h2 = st.columns([4, 1])
with col_h2:
    if st.button("Sign Out", **STRETCH):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

st.write("")

if role == "admin" and st.session_state.default_pw:
    st.warning("You're still using the default admin password. Define `[USERS]` in Streamlit secrets "
               "so passwords survive reboots (see setup notes).")

if flash := st.session_state.pop("flash", None):
    getattr(st, flash[0])(flash[1])

# --------------------------------------------------------------------------
# Store / period
# --------------------------------------------------------------------------
today = date.today()
YEARS = [str(y) for y in range(2025, today.year + 3)]
c_s, c_m, c_y = st.columns(3)
with c_s:
    if role == "store":
        active_store = st.session_state.name
        if active_store not in STORES:
            st.error("This account isn't linked to a valid store.")
            st.stop()
        st.markdown(f"**Store:** {active_store}")
    else:
        active_store = st.selectbox("Store", STORES)
with c_m:
    month = st.selectbox("Month", MONTHS, index=today.month - 1)
with c_y:
    year = st.selectbox("Year", YEARS, index=YEARS.index(str(today.year)))

spk = f"{active_store}_{month}_{year}"
period_label = f"{month} {year}"

master, source, sync_err = fetch_master()
stored = master.get(spk, {})
saved_files = stored.get("files", {})
saved_chk, saved_notes = stored.get("checklist", {}), stored.get("notes", {})
pend = st.session_state.pending.setdefault(spk, {})
ver = st.session_state.upver

if get_dbx() is None:
    st.info("Dropbox isn't configured: data lives only on this server and **resets when the app reboots**.")
elif source == "local":
    st.warning(f"Couldn't reach Dropbox, showing the local copy. {sync_err}")

st.markdown(f"### Active Workspace: **{active_store}** — {period_label}")


def accept_upload(key: str, f) -> bool:
    """Register an uploaded file as pending; auto-tick its checklist box when it's new."""
    if f.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(f"{f.name} is over {MAX_UPLOAD_MB} MB.")
        return False
    pend[key] = (f.name, f.getvalue())
    seen = f"seen_{spk}_{key}"
    fid = getattr(f, "file_id", None) or (f.name, f.size)
    if st.session_state.get(seen) != fid:
        st.session_state[seen] = fid
        st.session_state[f"chk_{spk}_{'green1' if key == 'inventory' else key}"] = True
    return True


# --------------------------------------------------------------------------
# Main layout
# --------------------------------------------------------------------------
col_left, col_right = st.columns([1, 1])
expenses_box = st.container()   # positioned here, filled last
docs_box = st.container()       # positioned under expenses, filled FIRST so uploads can tick the checklist

with docs_box:
    with st.expander("📎 Attach supporting documents (optional) — attaching a file ticks its checklist box"):
        up_cols = st.columns(2)
        for i, (key, label) in enumerate(CHECKLIST_ITEMS):
            with up_cols[i % 2]:
                f = st.file_uploader(label, type=DOC_TYPES, key=f"up_{spk}_{key}_{ver}")
                if f is not None:
                    accept_upload(key, f)
                else:
                    pend.pop(key, None)

with col_left:
    st.subheader("Inventory Report Formatter")
    up = st.file_uploader("Upload raw monthly CSV", type=["csv", "xlsx", "xls"], key=f"up_{spk}_inventory_{ver}")
    if up is not None:
        accept_upload("inventory", up)
    else:
        pend.pop("inventory", None)

    inventory_df, inv_notes, inv_name = None, [], stored.get("inventory_filename", "")
    if "inventory" in pend:
        name, data = pend["inventory"]
        try:
            inventory_df, inv_notes = parse_inventory(data, name)
            inv_name = name
            st.success(f"Successfully formatted: {name}")
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
            pend.pop("inventory", None)
    elif stored.get("inventory_data"):
        inventory_df = pd.DataFrame(stored["inventory_data"])

    for n in inv_notes:
        st.warning(n)
    if inventory_df is not None:
        st.caption(f"Loaded File: **{inv_name}**")
        st.dataframe(inventory_df.head(15), **STRETCH)
        if len(inventory_df) > 15:
            st.caption(f"Showing the first 15 of {len(inventory_df)} rows. The download has all of them.")
        st.download_button("⬇️ Download formatted inventory (.xlsx)", inv_xlsx(inventory_df, "Inventory"),
                           file_name=f"{safe_name(active_store)} - Inventory - {period_label}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No inventory CSV uploaded for this period yet.")

with col_right:
    st.subheader("Monthly Accountant Document Checklist")
    chk_states, note_states = {}, {}
    for key, label in CHECKLIST_ITEMS:
        c1, c2 = st.columns([1, 1.2])
        ck = f"chk_{spk}_{key}"
        if ck not in st.session_state:
            st.session_state[ck] = bool(saved_chk.get(key)) or key in saved_files
        with c1:
            chk_states[key] = st.checkbox(label, key=ck)
            if key in pend:
                st.caption(f"📎 {pend[key][0]} (unsaved)")
            elif key in saved_files:
                st.caption(f"📎 {saved_files[key]['name']}")
        with c2:
            note_states[key] = st.text_input(f"Note ({label})", value=saved_notes.get(key, ""), key=f"note_{spk}_{key}")

    st.markdown("#### Cash on Hand & Chargebacks")
    cooh_text = st.text_input("Cash on Hand Amount", value=stored.get("cooh_text", ""), key=f"cooh_{spk}")
    try:
        parsed_date = datetime.strptime(stored["cooh_date"], "%Y-%m-%d").date() if stored.get("cooh_date") else None
    except ValueError:
        parsed_date = None
    cooh_date = st.date_input("Cash on Hand Date", value=parsed_date, key=f"cooh_date_{spk}")
    chargebacks_text = st.text_input("Moneris Chargebacks", value=stored.get("chargebacks_text", "Nothing to report."),
                                     key=f"cb_{spk}")

# --- OPERATING EXPENSES ---
with expenses_box:
    st.subheader("Operating Expenses")
    exp_saved = stored.get("expenses", {})
    exp_cols = st.columns(5)
    expenses = {}
    for i, (key, label) in enumerate(EXPENSE_FIELDS):
        with exp_cols[i % 5]:
            expenses[key] = st.text_input(label, value=exp_saved.get(key, "0.00"), key=f"exp_{spk}_{key}")


def current_payload() -> dict:
    files = dict(saved_files)
    for k, (n, d) in pend.items():
        files[k] = {"name": n, "size": len(d), "unsaved": True}
    return {
        "checklist": chk_states, "notes": note_states,
        "cooh_text": cooh_text, "cooh_date": str(cooh_date) if cooh_date else "",
        "chargebacks_text": chargebacks_text, "expenses": expenses,
        "inventory_data": inv_records(inventory_df) if inventory_df is not None else None,
        "inventory_filename": inv_name, "files": files,
        "saved_by": st.session_state.username, "saved_at": datetime.now().isoformat(timespec="seconds"),
    }


# --------------------------------------------------------------------------
# Save
# --------------------------------------------------------------------------
if st.button("💾 Save All to Cloud & Dropbox", type="primary", **STRETCH):
    with st.spinner("Saving…"):
        payload = current_payload()
        fell_back = False
        for k, (n, d) in list(pend.items()):
            meta = store_file(spk, k, n, d)
            fell_back |= meta["backend"] == "local" and get_dbx() is not None
            payload["files"][k] = meta
        db_save(spk, active_store, period_label, payload)
        cloud_error = None
        try:
            base = dbx_read_master()  # merge into the latest copy so other stores' saves aren't overwritten
            if base is None:
                raise SyncError("Dropbox isn't configured in Streamlit secrets.")
            if not base:
                base = db_read_all()
            base[spk] = payload
            dbx_write_master(base)
        except SyncError as e:
            cloud_error = str(e)
        fetch_master.clear()
        _dbx_download.clear()
        pend.clear()
        st.session_state.upver += 1  # resets the file uploaders
    if cloud_error:
        st.session_state.flash = ("warning", f"Saved on the server only. Dropbox sync failed: {cloud_error}")
    elif fell_back:
        st.session_state.flash = ("warning", "Data synced, but some files couldn't be uploaded to Dropbox and were kept on the server only.")
    else:
        st.session_state.flash = ("success", "Saved and synced to Dropbox. Your phone will now reflect these updates.")
    st.rerun()

# --------------------------------------------------------------------------
# Report generator
# --------------------------------------------------------------------------
st.divider()
st.subheader("📦 Accountant Package")
st.caption("Builds an Excel report (summary, checklist, expenses, formatted inventory and a sheet for every "
           "uploaded CSV/Excel document) plus a ZIP with all the original files. Uses what's on screen, saved or not.")
if st.button("🔄 Refresh from cloud", help="Re-read the latest saved data (e.g. after another store saved)"):
    fetch_master.clear()
    _dbx_download.clear()
    for k in [k for k in st.session_state if k.startswith(("chk_", "note_", "exp_", "cooh_", "cb_", "seen_"))]:
        del st.session_state[k]
    st.rerun()
if st.button("⚙️ Generate report package", **STRETCH):
    with st.spinner("Reading documents and building the report…"):
        payload = current_payload()
        docs = gather_documents(payload, pend)
        xlsx, entries = package_entries(active_store, period_label, payload, docs)
        st.session_state.pkg = {"spk": spk, "xlsx": xlsx, "zip": build_zip(entries),
                                "docs": len(docs), "at": datetime.now().strftime("%H:%M:%S")}
pkg = st.session_state.get("pkg")
if pkg and pkg["spk"] == spk:
    base_name = safe_name(f"{active_store} - {period_label}")
    st.success(f"Package ready ({pkg['docs']} document(s) included, built {pkg['at']}).")
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Excel report", pkg["xlsx"], file_name=f"{base_name} - Report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       **STRETCH)
    d2.download_button("⬇️ Full package (ZIP)", pkg["zip"], file_name=f"{base_name} - Package.zip",
                       mime="application/zip", **STRETCH)

# --------------------------------------------------------------------------
# Accountant / admin review
# --------------------------------------------------------------------------
if is_staff:
    st.divider()
    st.subheader("All Stores Summary Review")
    payloads = {s_: master.get(f"{s_}_{month}_{year}", {}) for s_ in STORES}
    for s_, p in payloads.items():
        s_chk = sum(1 for v in p.get("checklist", {}).values() if v)
        s_cooh = p.get("cooh_text", "Not entered")
        st.markdown(f"**{s_}**: Checklist {s_chk}/10 Completed | Cash on Hand: `{s_cooh}`")

    with st.expander("Detailed view: missing documents and expenses"):
        rows = []
        for s_, p in payloads.items():
            miss = missing_docs(p) if p else [label for _, label in CHECKLIST_ITEMS]
            rows.append({
                "Store": s_,
                "Docs uploaded": len(p.get("files", {})) if p else 0,
                "Chargebacks": p.get("chargebacks_text") or "",
                "Expenses": total_expenses(p) if p else 0.0,
                "Missing": ", ".join(miss) if miss else "Complete",
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, **STRETCH,
                     column_config={"Expenses": st.column_config.NumberColumn(format="$%.2f")})

    if st.button("📚 Build consolidated package (all stores)", **STRETCH):
        have = {s: p for s, p in payloads.items() if p}
        if not have:
            st.info("No store has saved data for this period yet.")
        else:
            bar = st.progress(0.0, text="Starting…")
            entries = []
            for i, (s, p) in enumerate(have.items(), 1):
                bar.progress((i - 1) / len(have), text=f"Building {s}…")
                _, e = package_entries(s, period_label, p, gather_documents(p))
                entries += e
            bar.progress(1.0, text="Finishing…")
            summary = build_consolidated_workbook(period_label, have)
            entries.insert(0, (f"All Stores - {period_label} - Summary.xlsx", summary))
            st.session_state.all_pkg = {"key": f"{month}_{year}", "xlsx": summary, "zip": build_zip(entries)}
            bar.empty()
    ap = st.session_state.get("all_pkg")
    if ap and ap["key"] == f"{month}_{year}":
        a1, a2 = st.columns(2)
        a1.download_button("⬇️ All-stores summary (.xlsx)", ap["xlsx"],
                           file_name=f"All Stores - {period_label} - Summary.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           **STRETCH)
        a2.download_button("⬇️ Everything (ZIP)", ap["zip"], file_name=f"All Stores - {period_label} - Package.zip",
                           mime="application/zip", **STRETCH)
