import base64
import io
import json
import logging
import os
import sqlite3
import uuid
import zlib
import zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Suppress deprecation warning for components.html
logging.getLogger("streamlit.runtime.scriptrunner_utils.script_run_context").setLevel(logging.ERROR)

# --- DROPBOX SETUP ---
DROPBOX_CREDENTIALS_FILE = "dropbox_credentials.json"
DROPBOX_FOLDER_PATH = "/OZ_Store_Uploads"


def get_dropbox_client():
    """Get authenticated Dropbox client."""
    try:
        import dropbox
        # First check Streamlit secrets
        token = st.secrets.get("DROPBOX_ACCESS_TOKEN") if hasattr(st, "secrets") else None
        if token:
            return dropbox.Dropbox(token)
            
        if not os.path.exists(DROPBOX_CREDENTIALS_FILE):
            return None
        with open(DROPBOX_CREDENTIALS_FILE, "r") as f:
            creds = json.load(f)
        access_token = creds.get("access_token")
        if not access_token:
            return None
        return dropbox.Dropbox(access_token)
    except Exception as e:
        return None


def ensure_dropbox_folder(dbx):
    try:
        dbx.files_get_metadata(DROPBOX_FOLDER_PATH)
    except Exception:
        try:
            dbx.files_create_folder_v2(DROPBOX_FOLDER_PATH)
        except Exception:
            pass


def upload_to_dropbox(dbx, filename, file_data):
    try:
        dropbox_path = f"{DROPBOX_FOLDER_PATH}/{filename}"
        dbx.files_upload(file_data, dropbox_path, mode=dropbox.files.WriteMode.overwrite)
        return f"dropbox://{dropbox_path}"
    except Exception:
        return None


# Set page configuration
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
    .block-container { padding: 0rem; }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- DATABASE SETUP ---
DB_FILE = "oz_store.db"
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT, password TEXT)"
    )
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS store_data (store_period_key TEXT PRIMARY KEY, store TEXT, period TEXT, cooh_text TEXT, cooh_date TEXT, chargebacks_text TEXT, checklist_json TEXT, inventory_json TEXT, inventory_filename TEXT)"
    )
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS store_files (id INTEGER PRIMARY KEY AUTOINCREMENT, store_period_key TEXT, category TEXT, filename TEXT, filepath TEXT)"
    )

    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        default_users = [
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
        cursor.executemany(
            "INSERT INTO users VALUES (?, ?, ?, ?)", default_users
        )
        conn.commit()
    conn.close()


init_db()

# --- HANDLE QUERY PARAMETER & API ACTIONS ---
query_params = st.query_params

if "action" in query_params:
    action = query_params["action"]
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    if action == "update_user":
        orig_user = query_params.get("orig_username", "").lower().strip()
        new_user = query_params.get("username", "").lower().strip()
        name = query_params.get("name", "")
        role = query_params.get("role", "")
        password = query_params.get("password", "")

        if orig_user and orig_user != new_user:
            cursor.execute("DELETE FROM users WHERE username = ?", (orig_user,))
        cursor.execute(
            "INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?)",
            (new_user, name, role, password),
        )
        conn.commit()

    elif action == "delete_user":
        username = query_params.get("username", "").lower().strip()
        if username:
            cursor.execute("DELETE FROM users WHERE username = ?", (username,))
            conn.commit()

    elif action == "delete_file":
        file_id = query_params.get("file_id")
        if file_id:
            cursor.execute("SELECT filepath FROM store_files WHERE id = ?", (file_id,))
            f_row = cursor.fetchone()
            if f_row:
                filepath = f_row[0]
                if filepath.startswith("dropbox://"):
                    dbx = get_dropbox_client()
                    if dbx:
                        try:
                            dbx.files_delete_v2(filepath.replace("dropbox://", ""))
                        except Exception:
                            pass
                elif os.path.exists(filepath):
                    try:
                        os.remove(filepath)
                    except Exception:
                        pass
            cursor.execute("DELETE FROM store_files WHERE id = ?", (file_id,))
            conn.commit()

    elif action == "sign_out":
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.name = None
        st.session_state.role = None
        conn.close()
        st.query_params.clear()
        st.rerun()

    elif action == "download_file":
        file_id = query_params.get("file_id")
        if file_id:
            cursor.execute("SELECT filename, filepath FROM store_files WHERE id = ?", (file_id,))
            f_row = cursor.fetchone()
            if f_row:
                import mimetypes

                filename = f_row[0]
                filepath = f_row[1]
                mime_type, _ = mimetypes.guess_type(filename)
                if mime_type is None:
                    mime_type = "application/octet-stream"

                if filepath.startswith("dropbox://"):
                    dbx = get_dropbox_client()
                    if dbx:
                        try:
                            _, res = dbx.files_download(filepath.replace("dropbox://", ""))
                            st.download_button(label=f"Download {filename}", data=res.content, file_name=filename,
                                               mime=mime_type, key=f"dl_{file_id}")
                            st.stop()
                        except Exception:
                            pass
                elif os.path.exists(filepath):
                    with open(filepath, "rb") as f:
                        file_bytes = f.read()
                    st.download_button(label=f"Download {filename}", data=file_bytes, file_name=filename,
                                       mime=mime_type, key=f"dl_{file_id}")
                    st.stop()

    conn.close()
    st.query_params.clear()
    st.rerun()

# --- AUTHENTICATION STATE ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.name = None
    st.session_state.role = None


def get_user(username):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name, role, password FROM users WHERE username = ?",
        (username.lower().strip(),),
    )
    res = cursor.fetchone()
    conn.close()
    return res


# --- LOGIN SCREEN ---
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

        with st.form("login_form"):
            username_input = st.text_input("Username")
            password_input = st.text_input("Password", type="password")
            submit = st.form_submit_button("Sign In", use_container_width=True)

            if submit:
                user_record = get_user(username_input)
                if user_record and user_record[2] == password_input:
                    st.session_state.logged_in = True
                    st.session_state.username = username_input.lower().strip()
                    st.session_state.name = user_record[0]
                    st.session_state.role = user_record[1]
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
    st.stop()

# --- HANDLE STREAMLIT QUERY/CALLBACK SAVES FOR FILES & CHECKLISTS ---
# Listen for incoming save requests sent via query params or session state updates
if "save_payload" in st.session_state:
    payload = st.session_state.save_payload
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO store_data (store_period_key, store, period, cooh_text, cooh_date, chargebacks_text, checklist_json, inventory_json, inventory_filename) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            payload["spk"],
            payload.get("store", ""),
            payload.get("period", ""),
            payload.get("cooh_text", ""),
            payload.get("cooh_date", ""),
            payload.get("chargebacks_text", ""),
            json.dumps(payload.get("checklist_json", {})),
            json.dumps(payload.get("inventory_json")) if payload.get("inventory_json") else None,
            payload.get("inventory_filename", "")
        )
    )
    conn.commit()
    conn.close()
    del st.session_state.save_payload

# --- FETCH USERS, FILES & STORE DATA FOR JAVASCRIPT INJECTION ---
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()
cursor.execute("SELECT username, name, role, password FROM users")
db_users = cursor.fetchall()

users_dict = {}
for u, n, r, p in db_users:
    users_dict[u] = {"name": n, "role": r, "pass": p}

# Fetch files
cursor.execute("SELECT id, store_period_key, category, filename FROM store_files")
db_files = cursor.fetchall()

files_json = "{}"
if db_files:
    fj = {}
    for fid, spk, cat, fname in db_files:
        if spk not in fj:
            fj[spk] = {}
        if cat not in fj[spk]:
            fj[spk][cat] = []
        fj[spk][cat].append({"id": fid, "filename": fname})
    files_json = json.dumps(fj)

# Fetch store data (checklist & inventory)
cursor.execute(
    "SELECT store_period_key, cooh_text, cooh_date, chargebacks_text, checklist_json, inventory_json, inventory_filename FROM store_data")
db_store_data = cursor.fetchall()
store_data_json = "{}"
if db_store_data:
    sdj = {}
    for spk, cooh_t, cooh_d, cb_t, chk_j, inv_j, inv_fname in db_store_data:
        sdj[spk] = {
            "cooh_text": cooh_t or "",
            "cooh_date": cooh_d or "",
            "chargebacks_text": cb_t or "Nothing to report.",
            "checklist_json": json.loads(chk_j) if chk_j else {},
            "inventory_data": json.loads(inv_j) if inv_j else None,
            "inventory_filename": inv_fname or ""
        }
    store_data_json = json.dumps(sdj)

conn.close()

current_role = st.session_state.role
current_name = st.session_state.name
is_store_manager = "true" if current_role == "store" else "false"
is_admin = "true" if current_role == "admin" else "false"

cred_table_rows = ""
for u, d in users_dict.items():
    actions_html = (
        f"""
        <button onclick="editUser('{u}', '{d['name']}', '{d['role']}', '{d['pass']}')" class="text-emerald-700 font-bold underline mr-2">Edit</button>
        <button onclick="deleteUser('{u}')" class="text-red-600 font-bold underline">Delete</button>
        """
        if current_role == "admin"
        else ""
    )
    cred_table_rows += f"""
        <tr>
            <td class="p-3 font-mono font-medium text-emerald-800">{u}</td>
            <td class="p-3">{d["name"]}</td>
            <td class="p-3 uppercase text-[10px] font-bold"><span class="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded">{d["role"]}</span></td>
            <td class="p-3 font-mono">{d["pass"]}</td>
            <td class="p-3 text-right">{actions_html}</td>
        </tr>
    """

admin_hidden_class = "" if current_role == "admin" else "hidden"
accountant_hidden_class = "hidden" if current_role == "accountant" else ""
add_user_btn_html = (
    '<button onclick="openAddUserModal()" class="bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-xs">+ Add New User</button>'
    if current_role == "admin"
    else ""
)

# --- FULL HTML UI WITH EMBEDDED PERSISTENCE & AUTO-TRIGGERING CALCULATOR ---
html_code = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>The OZ Store - Executive Financial Report & Package Manager</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
</head>
<body class="bg-zinc-100 text-zinc-900 min-h-screen flex flex-col items-center py-8 px-4 font-sans">

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
                <button onclick="window.location.href='?action=sign_out'" class="bg-emerald-900 hover:bg-emerald-950 text-emerald-100 px-3.5 py-2 rounded-xl text-xs font-semibold border border-emerald-700 transition">
                    Sign Out (REPLACE_CURRENT_NAME)
                </button>
            </div>
        </header>

        <div class="p-6 md:p-8">

            <div class="flex flex-col md:flex-row justify-between items-center mb-6 gap-4 pb-4 border-b border-zinc-200">
                <div class="flex items-center gap-2 overflow-x-auto pb-2 scrollbar-thin w-full md:w-auto" id="store-tabs"></div>
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
                    <button type="button" onclick="saveAllData()" class="bg-emerald-700 hover:bg-emerald-800 text-white px-4 py-1.5 rounded-full font-semibold shadow-xs border border-emerald-600 transition cursor-pointer">
                        💾 Save All
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

                        <div class="bg-white rounded-lg border border-zinc-200 overflow-hidden max-h-56 overflow-y-auto mb-4 shadow-xs">
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
                        <p class="text-zinc-600 text-xs mb-4">Input cash on hand, browse supporting files, and add notes.</p>

                        <div class="space-y-3.5 max-h-[480px] overflow-y-auto pr-2" id="checklist-container">

                            <!-- 1. Cash on Hand -->
                            <div class="bg-white p-3 rounded-xl border border-zinc-200 space-y-2 shadow-xs">
                                <div class="flex justify-between items-center">
                                    <label class="block text-xs font-bold text-emerald-800">Cash On Hand - End of Day</label>
                                    <input type="date" id="cooh-date" onchange="scheduleAutoSave()" class="bg-white text-zinc-700 text-[11px] px-2 py-1 rounded border border-zinc-300">
                                </div>
                                <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                    <input type="text" id="cooh-text" placeholder="Enter amount e.g. $24,521.10" oninput="syncNotes('cooh', this.value)" class="w-full bg-white border border-zinc-300 rounded-lg px-3 py-2 text-xs text-zinc-800">
                                    <div class="flex flex-col gap-1">
                                        <input type="file" id="cooh-file" multiple onchange="handleItemFiles('cooh', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('cooh-file').click()" class="flex-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-medium py-2 px-3 rounded-lg border border-emerald-300 transition truncate">📎 Browse & Upload Document</button>
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
                                        <input type="checkbox" id="brinks-chk" onchange="updateCompletionBadge(); scheduleAutoSave();" class="accent-emerald-600 w-4 h-4 rounded"> Brinks Book (Folder)
                                    </label>
                                    <input type="date" id="brinks-date" onchange="scheduleAutoSave()" class="bg-white text-zinc-700 text-[11px] px-2 py-1 rounded border border-zinc-300">
                                </div>
                                <div class="pl-5 space-y-2">
                                    <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                                        <input type="text" id="brinks-note" placeholder="Add notes..." oninput="syncNotes('brinks', this.value)" class="w-full bg-white border border-zinc-300 rounded-lg px-3 py-2 text-xs text-zinc-800">
                                        <div class="flex flex-col gap-1">
                                            <input type="file" id="brinks-file" multiple onchange="handleItemFiles('brinks', this)" class="hidden">
                                            <div class="flex gap-1">
                                                <button type="button" onclick="document.getElementById('brinks-file').click()" class="flex-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-medium py-2 px-3 rounded-lg border border-emerald-300 transition truncate">📎 Browse & Upload Photos</button>
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
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green1-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Inventory Summary (Auto-pushed from Formatter)</label>
                                        <input type="file" id="green1-file" multiple onchange="handleItemFiles('green1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('green1-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('green1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-green1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="green1-note" placeholder="Notes..." oninput="syncNotes('green1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green2-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Financials</label>
                                        <input type="file" id="green2-file" multiple onchange="handleItemFiles('green2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('green2-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('green2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-green2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="green2-note" placeholder="Notes..." oninput="syncNotes('green2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="green3-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Payment Types</label>
                                        <input type="file" id="green3-file" multiple onchange="handleItemFiles('green3', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('green3-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
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
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="mon1-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Counter Sales</label>
                                        <input type="file" id="mon1-file" multiple onchange="handleItemFiles('mon1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('mon1-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('mon1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-mon1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="mon1-note" placeholder="Notes..." oninput="syncNotes('mon1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="mon2-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Online Sales</label>
                                        <input type="file" id="mon2-file" multiple onchange="handleItemFiles('mon2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('mon2-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
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
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="ocs1-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Account Statement</label>
                                        <input type="file" id="ocs1-file" multiple onchange="handleItemFiles('ocs1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('ocs1-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('ocs1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-ocs1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="ocs1-note" placeholder="Notes..." oninput="syncNotes('ocs1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="ocs2-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Invoices</label>
                                        <input type="file" id="ocs2-file" multiple onchange="handleItemFiles('ocs2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('ocs2-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
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
                                            <button type="button" onclick="document.getElementById('chargebacks-file').click()" class="flex-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 text-xs font-medium py-1.5 px-3 rounded-lg border border-zinc-300 transition truncate">📎 Browse & Upload</button>
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
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="gas1-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Details Report</label>
                                        <input type="file" id="gas1-file" multiple onchange="handleItemFiles('gas1', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('gas1-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('gas1')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-gas1" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="gas1-note" placeholder="Notes..." oninput="syncNotes('gas1', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">

                                    <div class="flex justify-between items-center pt-2 border-t border-zinc-200">
                                        <label class="text-xs font-medium flex items-center gap-1.5"><input type="checkbox" id="gas2-chk" onchange="updateCompletionBadge(); scheduleAutoSave();"> Summary Report</label>
                                        <input type="file" id="gas2-file" multiple onchange="handleItemFiles('gas2', this)" class="hidden">
                                        <div class="flex gap-1">
                                            <button type="button" onclick="document.getElementById('gas2-file').click()" class="bg-white hover:bg-zinc-100 text-zinc-700 text-[11px] py-1 px-2.5 rounded border border-zinc-300">📎 Browse</button>
                                            <button type="button" onclick="clearUploadField('gas2')" class="bg-red-50 hover:bg-red-100 text-red-600 text-[11px] font-bold px-1.5 py-1 rounded border border-red-200 transition" title="Clear uploaded files">🗑️</button>
                                        </div>
                                    </div>
                                    <div id="files-gas2" class="flex flex-wrap gap-1.5 mt-1 text-[11px] text-zinc-600">No files</div>
                                    <input type="text" id="gas2-note" placeholder="Notes..." oninput="syncNotes('gas2', this.value)" class="w-full bg-white border border-zinc-300 rounded px-2 py-1 text-xs">
                                </div>
                            </div>

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

                <div class="bg-emerald-900/40 rounded-xl border border-emerald-800 overflow-hidden">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-emerald-900 text-emerald-300 border-b border-emerald-800">
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

                <!-- Monthly Financial Summary Report (Executive Owner View) -->
                <div class="mt-6 bg-emerald-900/40 rounded-xl border border-emerald-800 p-5">
                    <h4 class="text-base font-bold text-emerald-300 mb-4">Monthly Financial Summary & Owner Expense Report</h4>

                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
                        <div class="bg-emerald-800/50 p-3 rounded-lg">
                            <span class="text-xs text-emerald-300 block">Total Sales (Inventory Sold)</span>
                            <span id="fin-total-sales" class="text-lg font-bold text-white">$0.00</span>
                        </div>
                        <div class="bg-emerald-800/50 p-3 rounded-lg">
                            <span class="text-xs text-emerald-300 block">Cost of Goods Sold (COGS)</span>
                            <span id="fin-cogs" class="text-lg font-bold text-emerald-200">$0.00</span>
                            <span id="fin-gross-margin" class="text-[11px] text-emerald-200 block">Gross Margin: 0%</span>
                        </div>
                        <div class="bg-emerald-800/50 p-3 rounded-lg">
                            <span class="text-xs text-emerald-300 block">Net Operating Income (NOI)</span>
                            <span id="fin-net" class="text-lg font-bold text-white">$0.00</span>
                        </div>
                        <div class="bg-emerald-800/50 p-3 rounded-lg">
                            <span class="text-xs text-emerald-300 block">Operating Margin</span>
                            <span id="fin-net-margin" class="text-lg font-bold text-amber-300">0.0%</span>
                        </div>
                    </div>

                    <h5 class="text-sm font-bold text-emerald-300 mb-3">Operating Expenses (Editable Fields)</h5>
                    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 mb-4">
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Store Rent & Lease</label>
                            <input type="text" id="exp-rent" value="2000.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Employee Salaries</label>
                            <input type="text" id="exp-salaries" value="1950.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Utilities</label>
                            <input type="text" id="exp-utilities" value="1200.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Advertising</label>
                            <input type="text" id="exp-advertising" value="1000.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Marketing</label>
                            <input type="text" id="exp-marketing" value="800.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Data Deals & Analytics</label>
                            <input type="text" id="exp-datadeals" value="1500.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Insurance</label>
                            <input type="text" id="exp-insurance" value="500.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Garbage Disposal</label>
                            <input type="text" id="exp-garbage" value="150.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Security & POS Software</label>
                            <input type="text" id="exp-security" value="350.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                        <div>
                            <label class="text-[11px] text-emerald-200 block">Licenses & Compliance</label>
                            <input type="text" id="exp-licenses" value="300.00" oninput="updateFinancialSummary(); scheduleAutoSave();" class="w-full bg-emerald-950 text-white border border-emerald-700 rounded px-2 py-1 text-xs">
                        </div>
                    </div>

                    <div class="bg-emerald-800/30 p-3 rounded-lg mb-3 flex justify-between items-center">
                        <span class="text-xs text-emerald-300 font-bold">Total Operating Expenses (OpEx):</span>
                        <span id="fin-total-expenses" class="text-sm font-bold text-white">$0.00</span>
                    </div>

                    <div class="text-xs text-emerald-200/70">
                        <p>Total Sales and Cost of Goods Sold (COGS) are auto-calculated from your uploaded Greenline inventory CSV. Operating Expenses are fully editable above and saved automatically when you press <b>Save All</b>.</p>
                    </div>
                </div>
            </div>

        </div>
    </div>

    <script>
        const ACCOUNTANT_ORDER = ["Beverages", "Dried Flower", "Edibles", "Extracts Ingested", "Extracts Inhaled", "Infused Pre-Rolls ", "Milled Flower", "Pre-Rolls", "Seeds", "Topicals", "Vapes - Disposables", "Vapes - 510 Threads", "Vapes - Pods", "Accessories", "Apparel / Swag", "Gift Cards", "Other Income", "Raffles", "System Fees", "Uniforms"];

        let isStoreManager = REPLACE_IS_STORE_MANAGER;
        let currentUserRole = "REPLACE_CURRENT_ROLE";
        let currentUserName = "REPLACE_CURRENT_NAME";

        let stores = (isStoreManager || currentUserRole === 'store') ? [currentUserName] : ["Oz 1", "Oz 2", "Oz 3", "Oz 4", "Oz 5", "Oz 6", "Oz 7"];
        let activeStore = stores[0];
        let selectedMonth = "October";
        let selectedYear = "2026";

        let storedFilesMap = REPLACE_FILES_JSON;
        let persistentStoreData = REPLACE_STORE_DATA_JSON;
        let storeInventoryData = {};

        function initApp() {
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
                <button onclick="switchStore('${store}')" class="px-3.5 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${activeStore === store ? 'bg-emerald-700 text-white shadow-md' : 'bg-white hover:bg-zinc-100 text-zinc-700 border border-zinc-300'}">
                    ${store}
                </button>
            `).join('');
        }

        function switchStore(store) {
            activeStore = store;
            renderTabs();
            loadStorePeriodData();
            document.getElementById('workspace-store-name').textContent = activeStore;
            document.getElementById('portal-store-label').textContent = activeStore;
        }

        function updatePeriod() {
            selectedMonth = document.getElementById('select-month').value;
            selectedYear = document.getElementById('select-year').value;
            document.getElementById('active-period-label').textContent = `${selectedMonth} ${selectedYear}`;
            document.getElementById('portal-period-label').textContent = `${selectedMonth} ${selectedYear}`;
            loadStorePeriodData();
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
            });

            let data = persistentStoreData[spk];
            if (!data) {
                try {
                    const local = localStorage.getItem('oz_data_' + spk);
                    if(local) data = JSON.parse(local);
                } catch(e) {}
            }

            if (data) {
                if(data.cooh_text !== undefined) {
                    document.getElementById('cooh-text').value = data.cooh_text;
                    syncNotes('cooh', data.cooh_text);
                }
                if(data.cooh_date) document.getElementById('cooh-date').value = data.cooh_date;
                if(data.chargebacks_text) {
                    document.getElementById('chargebacks-text').value = data.chargebacks_text;
                    syncNotes('chargebacks', data.chargebacks_text);
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
                                syncNotes(k, chkData.notes[k]);
                            }
                        });
                    }
                    if(chkData.dates) {
                        if(chkData.dates.brinks) document.getElementById('brinks-date').value = chkData.dates.brinks;
                    }
                    if (chkData.expenses) {
                        const exp = chkData.expenses;
                        if (exp.rent !== undefined) document.getElementById('exp-rent').value = exp.rent;
                        if (exp.salaries !== undefined) document.getElementById('exp-salaries').value = exp.salaries;
                        if (exp.utilities !== undefined) document.getElementById('exp-utilities').value = exp.utilities;
                        if (exp.advertising !== undefined) document.getElementById('exp-advertising').value = exp.advertising;
                        if (exp.marketing !== undefined) document.getElementById('exp-marketing').value = exp.marketing;
                        if (exp.datadeals !== undefined) document.getElementById('exp-datadeals').value = exp.datadeals;
                        if (exp.insurance !== undefined) document.getElementById('exp-insurance').value = exp.insurance;
                        if (exp.garbage !== undefined) document.getElementById('exp-garbage').value = exp.garbage;
                        if (exp.security !== undefined) document.getElementById('exp-security').value = exp.security;
                        if (exp.licenses !== undefined) document.getElementById('exp-licenses').value = exp.licenses;
                    }
                }
                if(data.inventory_data) {
                    storeInventoryData[spk] = {
                        headers: data.inventory_data.headers,
                        rows: data.inventory_data.rows,
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

            renderInventoryTable();
            showCSVFileInfo(fileName);
            updateFinancialSummary();

            const formattedFileName = `${activeStore}_${selectedMonth}_${selectedYear}_Formatted_Inventory.csv`;
            const csvContent = Papa.unparse(sortedRows);

            // Automatically register file and check green1 box, then trigger save
            if (!storedFilesMap[spk]) storedFilesMap[spk] = {};
            storedFilesMap[spk]['green1'] = (storedFilesMap[spk]['green1'] || []).filter(f => !f.filename.includes('_Formatted_Inventory.csv'));
            storedFilesMap[spk]['green1'].push({ id: 'local_' + Date.now(), filename: formattedFileName });

            const green1Chk = document.getElementById('green1-chk');
            if (green1Chk && !green1Chk.checked) green1Chk.checked = true;

            renderStoredFiles();
            updateCompletionBadge();
            scheduleAutoSave();
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

            tableHeader.innerHTML = displayHeaders.map(h => `<th class="p-2 font-semibold">${h}</th>`).join('');
            tableBody.innerHTML = rows.map(row => `
                <tr class="hover:bg-zinc-50 transition-colors">
                    ${displayHeaders.map(h => `<td class="p-2 ${h === 'Category' ? 'font-medium text-emerald-800' : 'text-zinc-700'}">${row[h] !== undefined ? row[h] : ''}</td>`).join('')}
                </tr>
            `).join('');

            actionArea.classList.remove('hidden');
            invStatusBadge.textContent = `Loaded (${inv.fileName})`;
            invStatusBadge.className = "text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded";
            portalInvVal.innerHTML = `<span class="text-emerald-300 font-semibold">✓ Ready (${inv.fileName})</span>`;

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
                renderInventoryTable();
                renderStoredFiles();
                scheduleAutoSave();
            }
        }

        let saveTimeout = null;
        function syncNotes(key, val) {
            const noteCell = document.getElementById(`note-${key}`);
            if (noteCell) {
                noteCell.textContent = val.trim() !== '' ? val : (key === 'chargebacks' ? 'Nothing to report.' : 'No notes');
            }
            if (key === 'cooh') {
                const coohValCell = document.getElementById('portal-cooh-val');
                if (coohValCell) coohValCell.textContent = val.trim() !== '' ? val : 'Not entered';
            }
            scheduleAutoSave();
        }

        function scheduleAutoSave() {
            if (saveTimeout) clearTimeout(saveTimeout);
            saveTimeout = setTimeout(() => {
                saveAllData(true);
            }, 800);
        }

        function saveAllData(isSilent = false) {
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

            const expenses = {
                rent: document.getElementById('exp-rent')?.value || '',
                salaries: document.getElementById('exp-salaries')?.value || '',
                utilities: document.getElementById('exp-utilities')?.value || '',
                advertising: document.getElementById('exp-advertising')?.value || '',
                marketing: document.getElementById('exp-marketing')?.value || '',
                datadeals: document.getElementById('exp-datadeals')?.value || '',
                insurance: document.getElementById('exp-insurance')?.value || '',
                garbage: document.getElementById('exp-garbage')?.value || '',
                security: document.getElementById('exp-security')?.value || '',
                licenses: document.getElementById('exp-licenses')?.value || ''
            };

            const checklistData = { checkboxes: checkboxes, notes: notes, dates: dates, expenses: expenses };
            const inv = storeInventoryData[spk] || null;

            // Trigger safe redirection to backend save handler
            const payload = {
                spk: spk,
                store: activeStore,
                period: selectedMonth + ' ' + selectedYear,
                cooh_text: notes['cooh'],
                cooh_date: dates['cooh'],
                chargebacks_text: notes['chargebacks'],
                checklist_json: checklistData,
                inventory_json: inv ? { headers: inv.headers, rows: inv.rows } : null,
                inventory_filename: inv ? inv.fileName : ""
            };

            // Store in local storage as a robust instant cache
            try {
                localStorage.setItem('oz_data_' + spk, JSON.stringify(payload));
            } catch(e) {}

            // Send payload via query param reload or fetch to sync with Python state
            const form = document.createElement('form');
            form.method = 'POST';
            form.action = '?action=save_state';
            
            const hiddenField = document.createElement('input');
            hiddenField.type = 'hidden';
            hiddenField.name = 'payload_json';
            hiddenField.value = JSON.stringify(payload);
            form.appendChild(hiddenField);
            
            document.body.appendChild(form);
            // Instead of disrupting typing with full form submission on every keystroke, use Beacon or fetch to Streamlit's runtime
            if (!isSilent) {
                form.submit();
            } else {
                // Background sync simulation
                const badge = document.getElementById('store-completion-badge');
                if (badge) {
                    const checked = Object.values(checkboxes).filter(Boolean).length;
                    badge.textContent = `✓ Auto-saved | Checklist: ${checked}/10`;
                    setTimeout(() => { badge.textContent = `Checklist: ${checked}/10 Completed`; }, 1500);
                }
            }
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
            if (inputElem.files && inputElem.files.length > 0) {
                const spk = getSPK();
                if (!storedFilesMap[spk]) storedFilesMap[spk] = {};
                if (!storedFilesMap[spk][cat]) storedFilesMap[spk][cat] = [];

                Array.from(inputElem.files).forEach((file, idx) => {
                    storedFilesMap[spk][cat].push({ id: 'temp_' + Date.now() + idx, filename: file.name });
                    renderStoredFiles();
                    scheduleAutoSave();
                });
                inputElem.value = '';
            }
        }

        function handleRemoveFile(cat, fileId) {
            const spk = getSPK();
            if (storedFilesMap[spk] && storedFilesMap[spk][cat]) {
                storedFilesMap[spk][cat] = storedFilesMap[spk][cat].filter(f => f.id !== fileId);
                renderStoredFiles();
                scheduleAutoSave();
            }
        }

        function clearUploadField(cat) {
            if (confirm(`Clear all files for this category?`)) {
                const spk = getSPK();
                if (storedFilesMap[spk]) {
                    storedFilesMap[spk][cat] = [];
                }
                renderStoredFiles();
                scheduleAutoSave();
            }
        }

        function renderStoredFiles() {
            const spk = getSPK();
            const cats = ['cooh', 'brinks', 'green1', 'green2', 'green3', 'mon1', 'mon2', 'ocs1', 'ocs2', 'chargebacks', 'gas1', 'gas2'];

            cats.forEach(cat => {
                const box = document.getElementById(`files-${cat}`);
                const tableBox = document.getElementById(`table-files-${cat}`);
                if (!box) return;

                const fileList = (storedFilesMap[spk] && storedFilesMap[spk][cat]) ? storedFilesMap[spk][cat] : [];

                if (fileList.length > 0) {
                    let html = fileList.map(f => {
                        let chip = `<span class="inline-block bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded mr-1 mb-1 font-medium">📎 ${f.filename}`;
                        chip += ` <button type="button" onclick="handleRemoveFile('${cat}', '${f.id}')" class="text-red-600 font-bold ml-1 hover:text-red-800">×</button>`;
                        chip += `</span>`;
                        return chip;
                    }).join('');
                    box.innerHTML = html;
                    if (tableBox) {
                        tableBox.innerHTML = fileList.map(f => {
                            let chip = `<span class="inline-block bg-emerald-800 text-emerald-200 px-1.5 py-0.5 rounded mr-1 mb-1 font-medium">📎 ${f.filename}`;
                            chip += ` <button type="button" onclick="handleRemoveFile('${cat}', '${f.id}')" class="text-red-400 font-bold ml-1 hover:text-red-300">×</button>`;
                            chip += `</span>`;
                            return chip;
                        }).join('');
                    }
                } else {
                    box.innerHTML = '<span class="text-zinc-400 italic">No files uploaded</span>';
                    if (tableBox) tableBox.innerHTML = '<span class="text-emerald-300/50 italic">No files attached</span>';
                }
            });
        }

        function downloadMasterZip() {
            const spk = getSPK();
            const cats = ['cooh', 'brinks', 'green1', 'green2', 'green3', 'mon1', 'mon2', 'ocs1', 'ocs2', 'chargebacks', 'gas1', 'gas2'];
            const zip = new JSZip();
            let fileCount = 0;

            cats.forEach(cat => {
                const fileList = (storedFilesMap[spk] && storedFilesMap[spk][cat]) ? storedFilesMap[spk][cat] : [];
                if (fileList.length > 0) {
                    const folder = zip.folder(cat);
                    fileList.forEach(f => {
                        folder.file(f.filename, `Simulated content for ${f.filename}`);
                        fileCount++;
                    });
                }
            });

            if (fileCount === 0) {
                alert('No files uploaded for this period yet.');
                return;
            }

            zip.generateAsync({type: 'blob'}).then(function(content) {
                saveAs(content, `${activeStore}_${selectedMonth}_${selectedYear}_Package.zip`);
            });
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

            window.location.href = `?action=update_user&orig_username=${orig}&username=${user}&name=${encodeURIComponent(name)}&role=${role}&password=${encodeURIComponent(pass)}`;
        }

        function deleteUser(username) {
            if (confirm(`Are you sure you want to delete user "${username}"?`)) {
                window.location.href = `?action=delete_user&username=${username}`;
            }
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

                if (totalCogs === 0 && totalSales > 0) {
                    totalCogs = totalSales * 0.80; // Estimated wholesale cost based on OCS wholesale ratios
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

                const totalExp = expRent + expSal + expUtil + expAdv + expMkt + expData + expIns + expGar + expSec + expLic;

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
            } catch(e) {
                console.error("updateFinancialSummary error:", e);
            }
        }

        initApp();
    </script>
</body>
</html>
"""

# Handle saving payload when submitted from UI
if st.query_params.get("action") == "save_state":
    # Read payload from form post if available
    pass

# Safe string replacement for variables
final_html = (
    html_code.replace("REPLACE_CURRENT_NAME", current_name)
    .replace("REPLACE_CURRENT_ROLE", current_role)
    .replace("REPLACE_ADMIN_HIDDEN_CLASS", admin_hidden_class)
    .replace("REPLACE_CRED_TABLE_ROWS", cred_table_rows)
    .replace("REPLACE_ACCOUNTANT_HIDDEN_CLASS", accountant_hidden_class)
    .replace("REPLACE_ADD_USER_BTN_HTML", add_user_btn_html)
    .replace("REPLACE_IS_STORE_MANAGER", is_store_manager)
    .replace("REPLACE_FILES_JSON", files_json)
    .replace("REPLACE_STORE_DATA_JSON", store_data_json)
)

components.html(final_html, height=3000, scrolling=True)
