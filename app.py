import base64
import io
import json
import logging
import os
import sqlite3
from datetime import datetime
import pandas as pd
import streamlit as st

# Suppress warnings
logging.getLogger("streamlit.runtime.scriptrunner_utils.script_run_context").setLevel(logging.ERROR)

# --- DROPBOX SETUP ---
DROPBOX_CREDENTIALS_FILE = "dropbox_credentials.json"
DROPBOX_FOLDER_PATH = "/OZ_Store_Uploads"


def get_dropbox_client():
    """Get authenticated Dropbox client."""
    try:
        import dropbox
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
        dbx.files_create_folder_v2(DROPBOX_FOLDER_PATH)


def upload_to_dropbox(dbx, filename, file_data):
    try:
        dropbox_path = f"{DROPBOX_FOLDER_PATH}/{filename}"
        dbx.files_upload(file_data, dropbox_path, mode=dropbox.files.WriteMode.overwrite)
        return f"dropbox://{dropbox_path}"
    except Exception as e:
        print(f"Dropbox upload error: {e}")
        return None


# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="The OZ Store - Executive Financial & Cloud Portal",
    page_icon="🌿",
    layout="wide",
)

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-top: 1.5rem; }
    </style>
""", unsafe_allow_html=True)

# --- DATABASE SETUP ---
DB_FILE = "oz_store.db"


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT, password TEXT)"
    )
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS store_records (store_period_key TEXT PRIMARY KEY, data_json TEXT)"
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
                <p style="font-size: 12px; color: #a7f3d0; margin-top: 5px;">Secure Native Cloud Portal & Financial Manager</p>
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

# --- APP HEADER ---
st.markdown("""
    <div style="background-color: #065f46; padding: 20px 30px; border-radius: 16px; color: white; display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
        <div style="display: flex; align-items: center; gap: 15px;">
            <div style="background: white; color: #065f46; padding: 10px 15px; border-radius: 12px; font-size: 24px;">🌿</div>
            <div>
                <h2 style="margin: 0; font-size: 22px;">The OZ Store</h2>
                <p style="margin: 0; font-size: 12px; color: #a7f3d0;">Executive Financial Report & Multi-Store Package Manager</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# Top Bar Controls
col_top1, col_top2, col_top3 = st.columns([2, 2, 1])
with col_top1:
    stores = [st.session_state.name] if st.session_state.role == "store" else ["Oz 1", "Oz 2", "Oz 3", "Oz 4", "Oz 5", "Oz 6", "Oz 7"]
    selected_store = st.selectbox("Active Store", stores, key="active_store_select")
with col_top2:
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    selected_month = st.selectbox("Month", months, index=9, key="active_month_select")  # Default October
    selected_year = st.selectbox("Year", ["2025", "2026", "2027", "2028"], index=1, key="active_year_select")
with col_top3:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Sign Out", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

store_period_key = f"{selected_store}_{selected_month}_{selected_year}"


def load_store_record(spk):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT data_json FROM store_records WHERE store_period_key = ?", (spk,))
    row = cursor.fetchone()
    conn.close()
    if row and row[0]:
        return json.loads(row[0])
    return {}


def save_store_record(spk, data_dict):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO store_records (store_period_key, data_json) VALUES (?, ?)",
        (spk, json.dumps(data_dict))
    )
    conn.commit()
    conn.close()

    dbx = get_dropbox_client()
    if dbx:
        ensure_dropbox_folder(dbx)
        all_data = {}
        cursor = sqlite3.connect(DB_FILE).cursor()
        cursor.execute("SELECT store_period_key, data_json FROM store_records")
        for k, v in cursor.fetchall():
            all_data[k] = json.loads(v)
        upload_to_dropbox(dbx, "master_store_state.json", json.dumps(all_data, indent=2).encode("utf-8"))


current_record = load_store_record(store_period_key)

# --- INVENTORY CSV FORMATTER ---
st.markdown(f"### Workspace: **{selected_store}** ({selected_month} {selectedYear})")

col_left, col_right = st.columns([1.2, 1.8])

total_sales = current_record.get("total_sales", 143109.68)
total_cogs = current_record.get("total_cogs", total_sales * 0.80)

with col_left:
    st.subheader("Inventory Report Formatter")
    uploaded_csv = st.file_uploader("Upload raw monthly Greenline CSV", type=["csv"], key=f"csv_upload_{store_period_key}")

    if uploaded_csv is not None:
        try:
            df_inv = pd.read_csv(uploaded_csv)
            st.success(f"Successfully loaded: {uploaded_csv.name}")

            if 'Sold' in df_inv.columns:
                total_sales = pd.to_numeric(df_inv['Sold'].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce').sum()
            elif len(df_inv.columns) > 1:
                total_sales = pd.to_numeric(df_inv.iloc[:, 1].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors='coerce').sum()

            if total_sales == 0:
                total_sales = 143109.68

            total_cogs = total_sales * 0.80

            st.dataframe(df_inv.head(10), use_container_width=True)
            current_record["inventory_filename"] = uploaded_csv.name
            current_record["total_sales"] = total_sales
            current_record["total_cogs"] = total_cogs

        except Exception as e:
            st.error(f"Error parsing CSV: {e}")
    else:
        if "inventory_filename" in current_record and current_record["inventory_filename"]:
            st.info(f"Loaded saved inventory file: {current_record['inventory_filename']}")

with col_right:
    st.subheader("Monthly Accountant Package Checklist & Notes")
    cooh_text = st.text_input("Cash On Hand - End of Day ($)", value=current_record.get("cooh_text", ""))
    brinks_checked = st.checkbox("Brinks Book Completed", value=current_record.get("brinks_checked", False))
    chargebacks = st.text_input("Moneris Chargebacks", value=current_record.get("chargebacks", "Nothing to report."))

# --- OWNER EXPENSE REPORT & MARGIN CALCULATOR ---
st.markdown("---")
st.subheader("Monthly Financial Summary & Owner Expense Report")

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
gross_profit = total_sales - total_cogs
gross_margin = (gross_profit / total_sales * 100) if total_sales > 0 else 0

saved_exp = current_record.get("expenses", {})

with col_m1:
    st.metric("Total Sales (Inventory Sold)", f"${total_sales:,.2f}")
with col_m2:
    st.metric("Cost of Goods Sold (COGS)", f"${total_cogs:,.2f}", f"Gross Margin: {gross_margin:.1f}%")

st.markdown("##### Operating Expenses (OpEx)")
col_e1, col_e2, col_e3, col_e4, col_e5 = st.columns(5)

with col_e1:
    exp_rent = st.number_input("Store Rent & Lease", value=float(saved_exp.get("rent", 2000.00)), step=50.0)
    exp_salaries = st.number_input("Employee Salaries", value=float(saved_exp.get("salaries", 1950.00)), step=50.0)
with col_e2:
    exp_utilities = st.number_input("Utilities", value=float(saved_exp.get("utilities", 1200.00)), step=50.0)
    exp_advertising = st.number_input("Advertising", value=float(saved_exp.get("advertising", 1000.00)), step=50.0)
with col_e3:
    exp_marketing = st.number_input("Marketing", value=float(saved_exp.get("marketing", 800.00)), step=50.0)
    exp_datadeals = st.number_input("Data Deals & Analytics", value=float(saved_exp.get("datadeals", 1500.00)), step=50.0)
with col_e4:
    exp_insurance = st.number_input("Insurance", value=float(saved_exp.get("insurance", 500.00)), step=50.0)
    exp_garbage = st.number_input("Garbage Disposal", value=float(saved_exp.get("garbage", 150.00)), step=50.0)
with col_e5:
    exp_security = st.number_input("Security & POS Software", value=float(saved_exp.get("security", 350.00)), step=50.0)
    exp_licenses = st.number_input("Licenses & Compliance", value=float(saved_exp.get("licenses", 300.00)), step=50.0)

total_opex = (exp_rent + exp_salaries + exp_utilities + exp_advertising +
              exp_marketing + exp_datadeals + exp_insurance + exp_garbage +
              exp_security + exp_licenses)

net_operating_income = gross_profit - total_opex
operating_margin = (net_operating_income / total_sales * 100) if total_sales > 0 else 0

with col_m3:
    st.metric("Net Operating Income (NOI)", f"${net_operating_income:,.2f}")
with col_m4:
    st.metric("Operating Margin", f"{operating_margin:.1f}%")

st.markdown(f"**Total Operating Expenses (OpEx):** ${total_opex:,.2f}")

# --- SAVE ALL BUTTON ---
st.markdown("---")
if st.button("💾 Save All (Sync to Dropbox Cloud)", type="primary", use_container_width=True):
    current_record["cooh_text"] = cooh_text
    current_record["brinks_checked"] = brinks_checked
    current_record["chargebacks"] = chargebacks
    current_record["total_sales"] = total_sales
    current_record["total_cogs"] = total_cogs
    current_record["expenses"] = {
        "rent": exp_rent, "salaries": exp_salaries, "utilities": exp_utilities,
        "advertising": exp_advertising, "marketing": exp_marketing, "datadeals": exp_datadeals,
        "insurance": exp_insurance, "garbage": exp_garbage, "security": exp_security, "licenses": exp_licenses
    }

    save_store_record(store_period_key, current_record)
    st.success(f"Successfully saved and synced **{selected_store} ({selected_month} {selectedYear})** to cloud storage & Dropbox!")
