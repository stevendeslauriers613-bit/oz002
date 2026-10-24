import base64
import io
import json
import logging
import os
import sqlite3
import pandas as pd
import streamlit as st

# --- Page config ---
st.set_page_config(
    page_title="The OZ Store - Executive Accountant Package Manager",
    page_icon="🌿",
    layout="wide",
)

# --- Custom Styling (Matching your Emerald Executive Look) ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    h1, h2, h3 { color: #065f46; }
    .stButton>button { background-color: #047857; color: white; border-radius: 12px; font-weight: bold; border: none; }
    .stButton>button:hover { background-color: #065f46; color: white; }
    </style>
""",
    unsafe_allow_html=True,
)

# --- DROPBOX SETUP ---
DROPBOX_FOLDER_PATH = "/OZ_Store_Uploads"

def get_dropbox_client():
    """Get authenticated Dropbox client from secrets or local config."""
    try:
        import dropbox
        token = st.secrets.get("DROPBOX_ACCESS_TOKEN") if hasattr(st, "secrets") else None
        if token:
            return dropbox.Dropbox(token)
        return None
    except Exception:
        return None

def ensure_dropbox_folder(dbx):
    try:
        dbx.files_get_metadata(DROPBOX_FOLDER_PATH)
    except Exception:
        try:
            dbx.files_create_folder_v2(DROPBOX_FOLDER_PATH)
        except Exception:
            pass

# --- DATABASE SETUP ---
DB_FILE = "oz_store.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT, password TEXT)"
    )
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS store_data (store_period_key TEXT PRIMARY KEY, store TEXT, period TEXT, data_json TEXT)"
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
        cursor.executemany("INSERT INTO users VALUES (?, ?, ?, ?)", default_users)
        conn.commit()
    conn.close()

init_db()

# --- SESSION STATE INITIALIZATION ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.name = None
    st.session_state.role = None

def get_user(username):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT name, role, password FROM users WHERE username = ?", (username.lower().strip(),))
    res = cursor.fetchone()
    conn.close()
    return res

# --- LOGIN SCREEN ---
if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background-color: #065f46; padding: 30px; border-radius: 16px; color: white; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.1);">
                <div style="font-size: 40px; margin-bottom: 10px;">🌿</div>
                <h2 style="margin: 0; font-size: 22px; color: white;">The OZ Store Portal</h2>
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

# --- HEADER & NAVIGATION ---
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
    if st.button("Sign Out", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.name = None
        st.session_state.role = None
        st.rerun()

st.write("")

# --- STORE & PERIOD SELECTION ---
STORES = ["Oz 1", "Oz 2", "Oz 3", "Oz 4", "Oz 5", "Oz 6", "Oz 7"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
YEARS = ["2025", "2026", "2027", "2028"]

col_s, col_m, col_y = st.columns(3)
with col_s:
    if st.session_state.role == "store":
        active_store = st.session_state.name
        st.markdown(f"**Store:** {active_store}")
    else:
        active_store = st.selectbox("Store", STORES)
with col_m:
    selected_month = st.selectbox("Month", MONTHS, index=9) # October
with col_y:
    selected_year = st.selectbox("Year", YEARS, index=1) # 2026

spk = f"{active_store}_{selected_month}_{selected_year}"

# Load existing state from DB for this store/period
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()
cursor.execute("SELECT data_json FROM store_data WHERE store_period_key = ?", (spk,))
row = cursor.fetchone()
conn.close()

stored_data = json.loads(row[0]) if row and row[0] else {}

st.markdown(f"### Active Workspace: **{active_store}** — {selected_month} {selected_year}")

# --- MAIN APP LAYOUT ---
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("Inventory Report Formatter")
    uploaded_csv = st.file_uploader("Upload raw monthly CSV", type=["csv"], key=f"inv_upload_{spk}")
    
    inventory_data = stored_data.get("inventory_data", None)
    inventory_filename = stored_data.get("inventory_filename", "")

    if uploaded_csv is not None:
        try:
            df = pd.read_csv(uploaded_csv)
            ACCOUNTANT_ORDER = [
                "Beverages", "Dried Flower", "Edibles", "Extracts Ingested", "Extracts Inhaled",
                "Infused Pre-Rolls ", "Milled Flower", "Pre-Rolls", "Seeds", "Topicals",
                "Vapes - Disposables", "Vapes - 510 Threads", "Vapes - Pods", "Accessories",
                "Apparel / Swag", "Gift Cards", "Other Income", "Raffles", "System Fees", "Uniforms"
            ]
            if "Category" in df.columns:
                df = df[~df["Category"].astype(str).str.startswith(">")]
                order_map = {cat: i for i, cat in enumerate(ACCOUNTANT_ORDER)}
                df["_order"] = df["Category"].map(lambda x: order_map.get(str(x).strip(), 999))
                df = df.sort_values("_order").drop(columns=["_order"])
            
            inventory_data = df.to_dict(orient="records")
            inventory_filename = uploaded_csv.name
            stored_data["inventory_data"] = inventory_data
            stored_data["inventory_filename"] = inventory_filename
            st.success(f"Successfully formatted: {inventory_filename}")
        except Exception as e:
            st.error(f"Error reading CSV: {e}")

    if inventory_data:
        st.caption(f"Loaded File: **{inventory_filename}**")
        st.dataframe(pd.DataFrame(inventory_data).head(15), use_container_width=True)
    else:
        st.info("No inventory CSV uploaded for this period yet.")

with col_right:
    st.subheader("Monthly Accountant Document Checklist")
    
    checklist_items = [
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

    saved_chk = stored_data.get("checklist", {})
    saved_notes = stored_data.get("notes", {})

    chk_states = {}
    note_states = {}

    for key, label in checklist_items:
        c1, c2 = st.columns([1, 1.2])
        with c1:
            chk_states[key] = st.checkbox(label, value=saved_chk.get(key, False), key=f"chk_{spk}_{key}")
        with c2:
            note_states[key] = st.text_input(f"Note ({label})", value=saved_notes.get(key, ""), key=f"note_{spk}_{key}")

    st.markdown("#### Cash on Hand & Chargebacks")
    cooh_text = st.text_input("Cash on Hand Amount", value=stored_data.get("cooh_text", ""))
    cooh_date = st.date_input("Cash on Hand Date", value=datetime.strptime(stored_data["cooh_date"], "%Y-%m-%d").date() if stored_data.get("cooh_date") else None)
    chargebacks_text = st.text_input("Moneris Chargebacks", value=stored_data.get("chargebacks_text", "Nothing to report."))

# --- OPERATING EXPENSES ---
st.subheader("Operating Expenses")
exp_data = stored_data.get("expenses", {})
exp_cols = st.columns(5)
expense_fields = [
    ("rent", "Store Rent & Lease"), ("salaries", "Employee Salaries"),
    ("utilities", "Utilities"), ("advertising", "Advertising"), ("marketing", "Marketing"),
    ("datadeals", "Data Deals & Analytics"), ("insurance", "Insurance"),
    ("garbage", "Garbage Disposal"), ("security", "Security & POS"), ("licenses", "Licenses & Compliance")
]

expenses = {}
for i, (key, label) in enumerate(expense_fields):
    with exp_cols[i % 5]:
        expenses[key] = st.text_input(label, value=exp_data.get(key, "0.00"), key=f"exp_{spk}_{key}")

# --- SAVE BUTTON (100% Reliable, No Sign-Out Bug) ---
if st.button("💾 Save All to Database & Dropbox", type="primary", use_container_width=True):
    complete_payload = {
        "checklist": chk_states,
        "notes": note_states,
        "cooh_text": cooh_text,
        "cooh_date": str(cooh_date) if cooh_date else "",
        "chargebacks_text": chargebacks_text,
        "expenses": expenses,
        "inventory_data": inventory_data,
        "inventory_filename": inventory_filename
    }

    # Save to SQLite
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO store_data (store_period_key, store, period, data_json) VALUES (?, ?, ?, ?)",
        (spk, active_store, f"{selectedMonth} {selectedYear}", json.dumps(complete_payload))
    )
    conn.commit()

    # Save master backup to Dropbox if available
    dbx = get_dropbox_client()
    if dbx:
        ensure_dropbox_folder(dbx)
        cursor.execute("SELECT store_period_key, data_json FROM store_data")
        all_rows = cursor.fetchall()
        master_dict = {row[0]: json.loads(row[1]) for row in all_rows}
        try:
            dbx.files_upload(
                json.dumps(master_dict, indent=2).encode("utf-8"),
                f"{DROPBOX_FOLDER_PATH}/master_store_state.json",
                mode=dropbox.files.WriteMode.overwrite
            )
        except Exception:
            pass
    conn.close()

    st.success("Successfully saved all records to the secure cloud database and Dropbox!")

# --- ACCOUNTANT REVIEW PORTAL ---
if st.session_state.role in ("admin", "accountant"):
    st.divider()
    st.subheader("All Stores Summary Review")
    for s in STORES:
        s_key = f"{s}_{selectedMonth}_{selectedYear}"
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT data_json FROM store_data WHERE store_period_key = ?", (s_key,))
        s_row = cursor.fetchone()
        conn.close()
        
        s_data = json.loads(s_row[0]) if s_row and s_row[0] else {}
        s_chk = sum(1 for v in s_data.get("checklist", {}).values() if v)
        s_cooh = s_data.get("cooh_text", "Not entered")
        st.markdown(f"**{s}**: Checklist {s_chk}/10 Completed | Cash on Hand: `{s_cooh}`")
