import json
import io
import base64
from datetime import datetime
import pandas as pd
import streamlit as st

# --- Page config ---
st.set_page_config(
    page_title="The OZ Store - Cloud Accountant Package",
    page_icon="🌿",
    layout="wide",
)

# --- Dropbox config ---
DROPBOX_FOLDER = "/OZ_Store_Uploads"
STATE_FILE = "master_store_state.json"

def get_dropbox_client():
    """Create Dropbox client from Streamlit secrets."""
    try:
        import dropbox
        token = st.secrets.get("DROPBOX_ACCESS_TOKEN")
        if not token:
            return None
        return dropbox.Dropbox(token)
    except Exception as e:
        st.error(f"Dropbox connection error: {e}")
        return None

def ensure_folder(dbx):
    try:
        dbx.files_get_metadata(DROPBOX_FOLDER)
    except Exception:
        try:
            dbx.files_create_folder_v2(DROPBOX_FOLDER)
        except Exception:
            pass

def load_state(dbx):
    """Load the shared JSON state from Dropbox."""
    if not dbx:
        return {}
    try:
        path = f"{DROPBOX_FOLDER}/{STATE_FILE}"
        _, res = dbx.files_download(path)
        return json.loads(res.content.decode("utf-8"))
    except Exception:
        return {}

def save_state(dbx, state):
    """Save the shared JSON state to Dropbox."""
    if not dbx:
        st.warning("Dropbox not connected — data will not be saved.")
        return False
    try:
        import dropbox
        path = f"{DROPBOX_FOLDER}/{STATE_FILE}"
        data = json.dumps(state, indent=2).encode("utf-8")
        dbx.files_upload(data, path, mode=dropbox.files.WriteMode.overwrite)
        return True
    except Exception as e:
        st.error(f"Failed to save state: {e}")
        return False

def upload_file_to_dropbox(dbx, filename, file_bytes):
    """Upload a binary file to Dropbox and return its path."""
    if not dbx:
        return None
    try:
        import dropbox
        path = f"{DROPBOX_FOLDER}/{filename}"
        dbx.files_upload(file_bytes, path, mode=dropbox.files.WriteMode.overwrite)
        return path
    except Exception as e:
        st.error(f"File upload failed: {e}")
        return None

def list_files_in_category(state, spk, category):
    """Return list of files for a store-period + category from state."""
    return state.get(spk, {}).get("files", {}).get(category, [])

# --- Session state init ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.name = None
    st.session_state.role = None

if "store_state" not in st.session_state:
    st.session_state.store_state = {}

# --- Simple local users ---
USERS = {
    "admin": {"name": "Master Admin", "role": "admin", "password": "admin2026"},
    "leanne": {"name": "Leanne (Accountant)", "role": "accountant", "password": "leanne2026"},
    "oz1": {"name": "Oz 1", "role": "store", "password": "oz1pass2026"},
    "oz2": {"name": "Oz 2", "role": "store", "password": "oz2pass2026"},
    "oz3": {"name": "Oz 3", "role": "store", "password": "oz3pass2026"},
    "oz4": {"name": "Oz 4", "role": "store", "password": "oz4pass2026"},
    "oz5": {"name": "Oz 5", "role": "store", "password": "oz5pass2026"},
    "oz6": {"name": "Oz 6", "role": "store", "password": "oz6pass2026"},
    "oz7": {"name": "Oz 7", "role": "store", "password": "oz7pass2026"},
}

# --- Login ---
if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown("### 🌿 The OZ Store Portal")
        st.caption("Shared Cloud Accountant Package (Dropbox-backed)")
        with st.form("login"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Sign In", use_container_width=True):
                user = USERS.get(username.lower().strip())
                if user and user["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.username = username.lower().strip()
                    st.session_state.name = user["name"]
                    st.session_state.role = user["role"]
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
    st.stop()

# --- After login: load Dropbox state ---
dbx = get_dropbox_client()
if dbx:
    ensure_folder(dbx)
    if not st.session_state.store_state:
        st.session_state.store_state = load_state(dbx)
else:
    st.warning("Dropbox token not found in secrets. Data will not persist across sessions.")

# --- Header ---
st.title("🌿 The OZ Store")
st.caption(f"Logged in as **{st.session_state.name}** ({st.session_state.role})")

if st.button("Sign Out"):
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.name = None
    st.session_state.role = None
    st.rerun()

# --- Store & Period selectors ---
STORES = ["Oz 1", "Oz 2", "Oz 3", "Oz 4", "Oz 5", "Oz 6", "Oz 7"]
MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
YEARS = ["2025", "2026", "2027", "2028"]

col_s, col_m, col_y = st.columns(3)
with col_s:
    if st.session_state.role == "store":
        active_store = st.session_state.name
        st.markdown(f"**Store:** {active_store}")
    else:
        active_store = st.selectbox("Store", STORES, index=0)
with col_m:
    selected_month = st.selectbox("Month", MONTHS, index=8)  # September
with col_y:
    selected_year = st.selectbox("Year", YEARS, index=1)  # 2026

spk = f"{active_store}_{selected_month}_{selected_year}"
st.markdown(f"### Workspace: **{active_store}** — {selected_month} {selected_year}")

# Ensure the period exists in state
if spk not in st.session_state.store_state:
    st.session_state.store_state[spk] = {
        "cooh_text": "",
        "cooh_date": "",
        "chargebacks_text": "Nothing to report.",
        "checklist": {},
        "notes": {},
        "expenses": {
            "rent": "2000.00", "salaries": "1950.00", "utilities": "1200.00",
            "advertising": "1000.00", "marketing": "800.00", "datadeals": "1500.00",
            "insurance": "500.00", "garbage": "150.00", "security": "350.00", "licenses": "300.00"
        },
        "inventory": None,
        "inventory_filename": "",
        "files": {}
    }

period_data = st.session_state.store_state[spk]

# --- Inventory CSV upload ---
st.subheader("Inventory Report Formatter")
uploaded_csv = st.file_uploader("Upload raw monthly CSV", type=["csv"], key=f"csv_{spk}")

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
        period_data["inventory"] = df.to_dict(orient="records")
        period_data["inventory_filename"] = uploaded_csv.name
        st.success(f"Loaded {uploaded_csv.name}")
        st.dataframe(df.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"Could not read CSV: {e}")

if period_data.get("inventory"):
    st.caption(f"Current inventory file: {period_data.get('inventory_filename', 'Unknown')}")

# --- Checklist & notes ---
st.subheader("Monthly Checklist & Notes")

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

if "checklist" not in period_data:
    period_data["checklist"] = {}
if "notes" not in period_data:
    period_data["notes"] = {}

cols = st.columns(2)
for i, (key, label) in enumerate(checklist_items):
    with cols[i % 2]:
        checked = st.checkbox(label, value=period_data["checklist"].get(key, False), key=f"chk_{spk}_{key}")
        period_data["checklist"][key] = checked
        note = st.text_input(f"Notes for {label}", value=period_data["notes"].get(key, ""), key=f"note_{spk}_{key}")
        period_data["notes"][key] = note

# Cash on Hand
st.markdown("**Cash On Hand**")
c1, c2 = st.columns(2)
with c1:
    period_data["cooh_text"] = st.text_input("Amount", value=period_data.get("cooh_text", ""), key=f"cooh_{spk}")
with c2:
    date_val = period_data.get("cooh_date")
    parsed_date = datetime.strptime(date_val, "%Y-%m-%d").date() if date_val else None
    new_date = st.date_input("Date", value=parsed_date, key=f"coohdate_{spk}")
    period_data["cooh_date"] = str(new_date) if new_date else ""

period_data["chargebacks_text"] = st.text_input("Moneris Chargebacks", value=period_data.get("chargebacks_text", "Nothing to report."), key=f"cb_{spk}")

# --- File uploads per category ---
st.subheader("Upload Supporting Documents")
file_categories = ["cooh", "brinks", "green1", "green2", "green3", "mon1", "mon2", "ocs1", "ocs2", "chargebacks", "gas1", "gas2"]

if "files" not in period_data:
    period_data["files"] = {}

for cat in file_categories:
    with st.expander(f"Upload files for {cat}"):
        uploaded = st.file_uploader(f"Choose files for {cat}", accept_multiple_files=True, key=f"up_{spk}_{cat}")
        if uploaded and dbx:
            for f in uploaded:
                safe_name = f"{spk}_{cat}_{f.name}"
                path = upload_file_to_dropbox(dbx, safe_name, f.getvalue())
                if path:
                    if cat not in period_data["files"]:
                        period_data["files"][cat] = []
                    if not any(x.get("filename") == f.name for x in period_data["files"][cat]):
                        period_data["files"][cat].append({"filename": f.name, "dropbox_path": path})
                    st.success(f"Uploaded {f.name}")
        existing = period_data["files"].get(cat, [])
        if existing:
            st.write("Uploaded:")
            for ef in existing:
                st.markdown(f"- {ef['filename']}")

# --- Expenses ---
st.subheader("Operating Expenses")
exp = period_data.get("expenses", {})
exp_cols = st.columns(5)
expense_fields = [
    ("rent", "Store Rent & Lease"), ("salaries", "Employee Salaries"),
    ("utilities", "Utilities"), ("advertising", "Advertising"), ("marketing", "Marketing"),
    ("datadeals", "Data Deals & Analytics"), ("insurance", "Insurance"),
    ("garbage", "Garbage Disposal"), ("security", "Security & POS"), ("licenses", "Licenses & Compliance")
]
for i, (key, label) in enumerate(expense_fields):
    with exp_cols[i % 5]:
        exp[key] = st.text_input(label, value=exp.get(key, "0.00"), key=f"exp_{spk}_{key}")
period_data["expenses"] = exp

# --- Simple financial summary ---
st.subheader("Financial Summary (from inventory)")
if period_data.get("inventory"):
    try:
        inv_df = pd.DataFrame(period_data["inventory"])
        sold_col = next((c for c in inv_df.columns if "sold" in c.lower()), None)
        total_sales = inv_df[sold_col].astype(float).sum() if sold_col else 0.0
        st.metric("Total Sold (from CSV)", f"${total_sales:,.2f}")
    except Exception:
        st.caption("Could not calculate sales from inventory.")
else:
    st.caption("Upload an inventory CSV to see sales totals.")

# --- Save button ---
if st.button("💾 Save All to Cloud", type="primary"):
    st.session_state.store_state[spk] = period_data
    if save_state(dbx, st.session_state.store_state):
        st.success("Saved to Dropbox. Other users will see this data on refresh.")
    else:
        st.error("Save failed. Check Dropbox connection.")

# --- Accountant / Admin view ---
if st.session_state.role in ("admin", "accountant"):
    st.divider()
    st.subheader("All Stores Status")
    for store in STORES:
        key = f"{store}_{selected_month}_{selected_year}"
        data = st.session_state.store_state.get(key, {})
        checked = sum(1 for v in data.get("checklist", {}).values() if v)
        st.markdown(f"**{store}**: Checklist {checked}/10 • Cash: {data.get('cooh_text', '—')}")

st.caption("Data is stored in Dropbox under /OZ_Store_Uploads. All users of this Streamlit Cloud app share the same state.")
