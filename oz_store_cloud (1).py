"""The OZ Store - Cloud Accountant Package Manager (Streamlit)."""
import hashlib
import hmac
import json
import logging
import re
import secrets as pysecrets
import sqlite3
import tempfile
from contextlib import closing
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from oz_report import (
    CHECKLIST_ITEMS,
    EXPENSE_FIELDS,
    build_consolidated_workbook,
    build_store_workbook,
    build_zip,
    checklist_done,
    clean_inventory,
    inventory_workbook,
    missing_docs,
    parse_money,
    read_table,
    total_expenses,
)

try:
    import dropbox
    from dropbox.exceptions import ApiError
    from dropbox.files import WriteMode
except ImportError:  # app still works (local only) without the package
    dropbox = None
    ApiError = Exception
    WriteMode = None

log = logging.getLogger("oz_store")

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
DB_FILE = "oz_store.db"
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


def db_read_all() -> dict:
    out = {}
    with _conn() as conn:
        rows = conn.execute(
            "SELECT store_period_key, data_json, cooh_text, cooh_date, chargebacks_text, "
            "checklist_json, inventory_json, inventory_filename FROM store_data").fetchall()
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
            if st.form_submit_button("Sign In", use_container_width=True):
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
    <div style="background-color:#065f46;padding:20px;border-radius:16px;color:white;">
        <h2 style="margin:0;color:white;">🌿 The OZ Store</h2>
        <p style="margin:0;font-size:12px;color:#a7f3d0;">Executive Financial & Accountant Package Manager</p>
    </div>
    """,
    unsafe_allow_html=True,
)
h1, h2, h3 = st.columns([4, 1, 1])
with h2:
    if st.button("🔄 Reload", use_container_width=True, help="Re-read the latest data from the cloud"):
        fetch_master.clear()
        _dbx_download.clear()
        for k in [k for k in st.session_state if k.startswith(("chk_", "note_", "exp_", "cooh_", "cb_", "seen_"))]:
            del st.session_state[k]
        st.rerun()
with h3:
    if st.button("Sign Out", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

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
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("Inventory Report Formatter")
    up = st.file_uploader("Upload raw monthly inventory file (CSV or Excel)", type=["csv", "xlsx", "xls"],
                          key=f"up_{spk}_inventory_{ver}")
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
            st.success(f"Formatted {name}: {len(inventory_df)} rows (not saved yet).")
        except Exception as e:
            st.error(f"Couldn't read that file: {e}")
            pend.pop("inventory", None)
    elif stored.get("inventory_data"):
        inventory_df = pd.DataFrame(stored["inventory_data"])

    for n in inv_notes:
        st.warning(n)
    if inventory_df is not None:
        st.caption(f"Loaded file: **{inv_name}**")
        st.dataframe(inventory_df, use_container_width=True, height=380)
        st.download_button("⬇️ Download formatted inventory (.xlsx)", inv_xlsx(inventory_df, "Inventory"),
                           file_name=f"{safe_name(active_store)} - Inventory - {period_label}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No inventory file uploaded for this period yet.")

with col_right:
    st.subheader("Monthly Accountant Document Checklist")
    st.caption("Attach each document and its box is ticked automatically. Click **Save** below to keep everything.")
    chk_states, note_states = {}, {}
    for key, label in CHECKLIST_ITEMS:
        c_chk, c_up, c_note = st.columns([1.15, 1.6, 1.25])
        ck = f"chk_{spk}_{key}"
        with c_up:
            f = st.file_uploader(label, type=DOC_TYPES, key=f"up_{spk}_{key}_{ver}", label_visibility="collapsed")
            if f is not None:
                accept_upload(key, f)
            else:
                pend.pop(key, None)
        if ck not in st.session_state:
            st.session_state[ck] = bool(saved_chk.get(key)) or key in saved_files
        with c_chk:
            chk_states[key] = st.checkbox(label, key=ck)
            if key in pend:
                st.caption(f"📎 {pend[key][0]} (unsaved)")
            elif key in saved_files:
                st.caption(f"📎 {saved_files[key]['name']}")
        with c_note:
            note_states[key] = st.text_input(f"Note ({label})", value=saved_notes.get(key, ""),
                                             key=f"note_{spk}_{key}", label_visibility="collapsed",
                                             placeholder="Note")

    st.markdown("#### Cash on Hand & Chargebacks")
    cooh_text = st.text_input("Cash on Hand Amount", value=stored.get("cooh_text", ""), key=f"cooh_{spk}")
    try:
        parsed_date = datetime.strptime(stored["cooh_date"], "%Y-%m-%d").date() if stored.get("cooh_date") else None
    except ValueError:
        parsed_date = None
    cooh_date = st.date_input("Cash on Hand Date", value=parsed_date, key=f"cooh_date_{spk}")
    chargebacks_text = st.text_input("Moneris Chargebacks", value=stored.get("chargebacks_text", "Nothing to report."),
                                     key=f"cb_{spk}")

# --------------------------------------------------------------------------
# Operating expenses
# --------------------------------------------------------------------------
st.subheader("Operating Expenses")
exp_saved = stored.get("expenses", {})
exp_cols = st.columns(5)
expenses = {}
for i, (key, label) in enumerate(EXPENSE_FIELDS):
    with exp_cols[i % 5]:
        expenses[key] = st.text_input(label, value=exp_saved.get(key, "0.00"), key=f"exp_{spk}_{key}")
st.metric("Total operating expenses", f"${sum(parse_money(v) for v in expenses.values()):,.2f}")


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
if st.button("💾 Save All to Cloud & Dropbox", type="primary", use_container_width=True):
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
if st.button("⚙️ Generate report package", use_container_width=True):
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
                       use_container_width=True)
    d2.download_button("⬇️ Full package (ZIP)", pkg["zip"], file_name=f"{base_name} - Package.zip",
                       mime="application/zip", use_container_width=True)

# --------------------------------------------------------------------------
# Accountant / admin review
# --------------------------------------------------------------------------
if is_staff:
    st.divider()
    st.subheader(f"All Stores Summary — {period_label}")
    payloads = {s: master.get(f"{s}_{month}_{year}", {}) for s in STORES}
    rows = []
    for s, p in payloads.items():
        miss = missing_docs(p) if p else [label for _, label in CHECKLIST_ITEMS]
        rows.append({
            "Store": s,
            "Progress": 100 * checklist_done(p) / len(CHECKLIST_ITEMS) if p else 0,
            "Docs uploaded": len(p.get("files", {})) if p else 0,
            "Cash on hand": p.get("cooh_text") or "Not entered",
            "Chargebacks": p.get("chargebacks_text") or "",
            "Expenses": total_expenses(p) if p else 0.0,
            "Missing": ", ".join(miss) if miss else "✅ Complete",
        })
    st.dataframe(
        pd.DataFrame(rows), hide_index=True, use_container_width=True,
        column_config={
            "Progress": st.column_config.ProgressColumn("Checklist", min_value=0, max_value=100, format="%d%%"),
            "Expenses": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

    if st.button("📚 Build consolidated package (all stores)", use_container_width=True):
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
                           use_container_width=True)
        a2.download_button("⬇️ Everything (ZIP)", ap["zip"], file_name=f"All Stores - {period_label} - Package.zip",
                           mime="application/zip", use_container_width=True)
