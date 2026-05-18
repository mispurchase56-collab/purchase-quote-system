
# ============================================================
#  Purchase Quote System – Streamlit + Google Colab
#  Run instructions at bottom of file
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import os
import io
import warnings
warnings.filterwarnings("ignore")

import gspread
from google.oauth2.service_account import Credentials

# ── Google Sheets Configuration ──────────────────────────────
CREDENTIALS_PATH = "credentials.json"
MASTER_SPREADSHEET_ID = "1SgKGzdUwjEzGNiRv-NnDFvw6ngJYvHsg0R6M-Y50ZTA"
DATABASE_SPREADSHEET_ID = "1I6osw4QSpH82SMFFEqSb4CR4SYErQby-HF_EQuQpofs"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

@st.cache_resource
def get_gspread_client():
    try:
        import json
        
        # 1. Check local file first (to prevent local Streamlit from raising an exception when secrets are empty)
        if os.path.exists(CREDENTIALS_PATH):
            with open(CREDENTIALS_PATH, 'r') as f:
                creds_info = json.load(f)
        # 2. Fall back to Streamlit Secrets (for cloud hosting like streamlit.app)
        else:
            try:
                if "gcp_service_account" in st.secrets:
                    creds_info = dict(st.secrets["gcp_service_account"])
                else:
                    st.error("❌ Missing credentials! Set st.secrets['gcp_service_account'] or provide credentials.json")
                    return None
            except Exception as e:
                st.error("❌ Missing credentials! Please provide credentials.json locally or configure Streamlit Secrets.")
                return None
        
        # AUTO-FIX: The private key often gets messed up with literal '\n' strings
        if 'private_key' in creds_info:
            pk = creds_info['private_key']
            # Replace literal "\n" strings with real newline characters
            pk = pk.replace("\\n", "\n")
            # Remove any accidental extra quotes or spaces
            pk = pk.strip().strip('"').strip("'")
            creds_info['private_key'] = pk

        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Failed to connect to Google Sheets: {e}")
        return None

gc = get_gspread_client()

@st.cache_data(ttl=600)
def load_gsheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Load data from a Google Sheet. Try API first, then Public CSV fallback."""
    # 1. Try Google Sheets API (gspread)
    if gc:
        try:
            sh = gc.open_by_key(spreadsheet_id)
            ws = sh.worksheet(sheet_name)
            data = ws.get_all_records()
            if data:
                df = pd.DataFrame(data)
                df.columns = [c.strip() for c in df.columns]
                return df
        except Exception as e:
            st.info(f"💡 Note: API access failed for '{sheet_name}'. Trying public link...")

    # 2. Fallback to Public CSV Export (Requires "Anyone with the link" access)
    # Mapping sheet names to their GIDs (you can find these in the URL gid=xxxx)
    gid_map = {
        "Item Master": "17422179",
        "Vendor Master": "359146087",
        "Location Master": "297795439",
        "Purchase Quotes": "0"
    }
    gid = gid_map.get(sheet_name, "0")
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
    
    try:
        df = pd.read_csv(url)
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception as e:
        st.warning(f"⚠️ Could not load '{sheet_name}': {e}")
        return pd.DataFrame()

def load_master_data():
    item_df = load_gsheet(MASTER_SPREADSHEET_ID, "Item Master")
    vendor_df = load_gsheet(MASTER_SPREADSHEET_ID, "Vendor Master")
    loc_df = load_gsheet(MASTER_SPREADSHEET_ID, "Location Master")
    
    # Process Location Master into a dict for the app
    loc_details = {}
    if not loc_df.empty:
        for _, r in loc_df.iterrows():
            code = str(r.get("Location Code", "")).strip()
            if code:
                loc_details[code] = {
                    "GSTIN": str(r.get("Buyer GSTIN", "")),
                    "Address": str(r.get("Delivery Address", "")),
                    "Contact": f"{r.get('Warehouse In-Charge name', '')} ({r.get('Warehouse In-Charge contact No', '')})"
                }
    return item_df, vendor_df, loc_details

# Load them globally for caching
item_df, vendor_df, LOCATION_DETAILS = load_master_data()

# ── Excel storage helpers ────────────────────────────────────
# ── Google Sheets storage helpers ────────────────────────────
REQUIRED_COLS = [
    "Quote No", "Quote Date", "Location Code", "Purchase Quote raised by", "Vendor Name", "Vendor No",
    "Cash", "Online Transfer", "CDC", "PDC", "Credit Days", 
    "Door Delivery", "Pickup", "Courier", "Expected Delivery Date", "Courier Details",
    "Vendor Address", "Vendor City", "Vendor GST", "Vendor Email",
    "S.No", "ERP Code", "Vendor Item No", "Product Description",
    "Qty", "Price Before GST", "GST %", "Price Inc. GST", "Total (Incl. GST)", "Remarks",
    "Freight Charge", "Status", "Created Date", "Modified Date",
    "Payment Method", "Mode of Delivery",
]

LOCAL_EXCEL_PATH = "Purchase_Quotes_Data.xlsx"

def load_quotes() -> pd.DataFrame:
    # Try Google Sheets first
    df = load_gsheet(DATABASE_SPREADSHEET_ID, "Purchase Quotes")
    if not df.empty:
        # Basic cleanup
        for c in REQUIRED_COLS:
            if c not in df.columns: df[c] = ""
        return df
    
    # Fallback to local Excel
    if os.path.exists(LOCAL_EXCEL_PATH):
        try:
            return pd.read_excel(LOCAL_EXCEL_PATH)
        except:
            pass
    return pd.DataFrame(columns=REQUIRED_COLS)

def save_quotes(new_df: pd.DataFrame, is_edit: bool = False, quote_no: str = None):
    # Enforce perfect column alignment to REQUIRED_COLS
    for col in REQUIRED_COLS:
        if col not in new_df.columns:
            new_df[col] = ""
    new_df = new_df[REQUIRED_COLS]

    # 1. Try Google Sheets Saving
    if gc:
        try:
            sh = gc.open_by_key(DATABASE_SPREADSHEET_ID)
            ws = sh.worksheet("Purchase Quotes")
            
            # Ensure Google Sheet headers match REQUIRED_COLS exactly
            headers = ws.row_values(1)
            if not headers:
                ws.append_row(REQUIRED_COLS)
            elif len(headers) < len(REQUIRED_COLS):
                ws.update('A1', [REQUIRED_COLS])
                
            if is_edit and quote_no:
                all_data = ws.get_all_records()
                full_df = pd.DataFrame(all_data)
                
                # Align existing data columns
                for col in REQUIRED_COLS:
                    if col not in full_df.columns:
                        full_df[col] = ""
                full_df = full_df[REQUIRED_COLS]
                
                full_df = full_df[full_df["Quote No"] != quote_no]
                full_df = pd.concat([full_df, new_df], ignore_index=True)
                ws.clear()
                ws.update([full_df.columns.values.tolist()] + full_df.values.tolist())
            else:
                ws.append_rows(new_df.values.tolist())
            return True
        except Exception as e:
            st.error(f"⚠️ Google Sheets Save Failed: {e}")

    # 2. Fallback to Local Excel
    try:
        full_df = pd.DataFrame(columns=REQUIRED_COLS)
        if os.path.exists(LOCAL_EXCEL_PATH):
            full_df = pd.read_excel(LOCAL_EXCEL_PATH)
        
        if is_edit and quote_no:
            full_df = full_df[full_df["Quote No"] != quote_no]
        
        full_df = pd.concat([full_df, new_df], ignore_index=True)
        full_df.to_excel(LOCAL_EXCEL_PATH, index=False)
        st.info("💾 Saved locally to 'Purchase_Quotes_Data.xlsx' as backup.")
        return True
    except Exception as e:
        st.error(f"❌ Error saving even to local Excel: {e}")
        return False

def next_quote_number(df: pd.DataFrame) -> str:
    if df.empty or "Quote No" not in df.columns:
        return "Q0001"
    existing = df["Quote No"].dropna().tolist()
    nums = []
    for q in existing:
        try:
            nums.append(int(str(q).replace("Q", "")))
        except Exception:
            pass
    nxt = max(nums) + 1 if nums else 1
    return f"Q{nxt:04d}"

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Purchase Quote System",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(160deg, #0f172a 0%, #1e293b 100%);
    color: #e2e8f0;
}
section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] .stRadio label { font-size: 15px; }

/* Metric cards */
div[data-testid="metric-container"] {
    background: linear-gradient(135deg, #1e293b, #0f172a);
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 16px;
    color: #f1f5f9;
}
div[data-testid="metric-container"] label { color: #94a3b8 !important; }
div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
    color: #38bdf8 !important; font-size: 2rem !important; font-weight: 700 !important;
}

/* Section headers */
.section-header {
    background: linear-gradient(90deg, #0ea5e9, #6366f1);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 1.4rem;
    font-weight: 700;
    margin-bottom: 12px;
}

/* ERP form card */
.erp-card {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 18px;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #0ea5e9, #6366f1);
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.5rem 1.8rem;
    font-weight: 600;
    font-size: 15px;
    transition: opacity .2s;
}
.stButton > button:hover { opacity: 0.88; }

/* Table */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* Success / Warning */
div[data-baseweb="notification"] { border-radius: 10px !important; }

/* Dark background for main area */
/* Ensure disabled text is clearly visible in light/dark mode */
div[data-testid="stTextInput"] input:disabled, 
div[data-testid="stNumberInput"] input:disabled {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    background-color: #f1f5f9 !important;
    font-weight: 600 !important;
    border: 1px solid #cbd5e1 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Session state defaults ───────────────────────────────────
if "role" not in st.session_state:
    st.session_state.role = "Salesperson"
if "line_items" not in st.session_state:
    st.session_state.line_items = []
if "edit_quote_no" not in st.session_state:
    st.session_state.edit_quote_no = None
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

# (LOCATION_DETAILS is now loaded dynamically from Google Sheets above)

# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    # Display Logo
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass
    
    st.markdown("## 📋 Purchase Quote System")
    st.markdown("---")
    st.session_state.role = st.radio(
        "Login As",
        ["Product manager / product team", "Admin / Purchase Team"],
        index=0 if st.session_state.role == "Product manager / product team" else 1,
    )
    st.markdown("---")
    # Both roles now use "Create / Edit Quote" as the primary workspace
    pages = ["Dashboard", "Create / Edit Quote", "View Quotes"]
    
    st.session_state.page = st.radio("Navigation", pages, index=pages.index(st.session_state.page) if st.session_state.page in pages else 0)
    st.markdown("---")
    st.caption("v1.0 | Purchase Quote ERP")

# (item_df and vendor_df are already loaded globally above)
quotes_df = load_quotes()

# ── Helper: safe column access ───────────────────────────────
def col(df, *names, default=""):
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series([default] * len(df))

# ── PO Excel Export ────────────────────────────────────────────
def export_po_excel(quote_no: str, df: pd.DataFrame) -> io.BytesIO:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"PO_{quote_no}"
    
    dark_blue = PatternFill(start_color="001f497d", end_color="001f497d", fill_type="solid")
    light_blue = PatternFill(start_color="00b4c6e7", end_color="00b4c6e7", fill_type="solid")
    dark_red = PatternFill(start_color="00c00000", end_color="00c00000", fill_type="solid")
    light_orange = PatternFill(start_color="fce4d6", end_color="fce4d6", fill_type="solid")
    
    w_bold = Font(color="FFFFFF", bold=True)
    b_bold = Font(color="000000", bold=True)
    c_align = Alignment(horizontal="center", vertical="center")
    r_align = Alignment(horizontal="right", vertical="center")
    
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    q_df = df[df["Quote No"] == quote_no]
    if q_df.empty: return None
    row1 = q_df.iloc[0]
    
    vendor_name = str(row1.get("Vendor Name", ""))
    v_addr = str(row1.get("Vendor Address", ""))
    v_city = str(row1.get("Vendor City", ""))
    v_gst = str(row1.get("Vendor GST", ""))
    v_email = str(row1.get("Vendor Email", ""))
    
    if (not v_addr or v_addr == "nan") and vendor_name:
        vendor_df = load_master_sheet("Vendor Master")
        if not vendor_df.empty:
            vrow = vendor_df[vendor_df["Name"] == vendor_name] if "Name" in vendor_df.columns else vendor_df[vendor_df["name"] == vendor_name]
            if not vrow.empty:
                vr = vrow.iloc[0]
                v_addr = str(vr.get("address", vr.get("Address", "")))
                v_city = str(vr.get("city", vr.get("City", "")))
                v_gst = str(vr.get("vatRegistrationNo", vr.get("GST", vr.get("GSTIN", ""))))
                v_email = str(vr.get("email", vr.get("Email", vr.get("E-Mail", ""))))
    
    v_addr = v_addr if v_addr != "nan" else ""
    v_city = v_city if v_city != "nan" else ""
    v_gst = v_gst if v_gst != "nan" else ""
    v_email = v_email if v_email != "nan" else ""
    
    # ── PO Header Details ──
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 15
    ws.row_dimensions[3].height = 15

    ws.merge_cells('A1:I1')
    ws['A1'] = "SUPREME COMPUTERS INDIA PVT. LTD."
    ws['A1'].font = Font(name='Calibri', size=16, bold=True)
    ws['A1'].alignment = c_align
    
    ws.merge_cells('A2:I2')
    ws['A2'] = "18/18, Majestic Plaza, Narasingapuram Street, Mount Road, Chennai - 600002"
    ws['A2'].font = Font(name='Calibri', size=11)
    ws['A2'].alignment = c_align
    
    loc_code = str(row1.get("Location Code", ""))
    loc_info = LOCATION_DETAILS.get(loc_code, {})
    gstin_val = loc_info.get("GSTIN", "33AAOCS1408H1ZR")
    ws.merge_cells('A3:I3')
    ws['A3'] = f"GSTIN: {gstin_val}  |  Tel: 044-42082020"
    ws['A3'].font = Font(name='Calibri', size=11)
    ws['A3'].alignment = c_align

    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        from openpyxl.drawing.image import Image
        img = Image(logo_path)
        img.width = 120
        img.height = 40
        img.anchor = 'A1'
        ws.add_image(img)

    ws.merge_cells('A5:I5')
    ws['A5'] = "PURCHASE ORDER"
    ws['A5'].fill = dark_blue
    ws['A5'].font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
    ws['A5'].alignment = c_align

    def set_cell(r, c1, c2, val, is_label=False):
        cell = ws[f'{c1}{r}']
        cell.value = val
        cell.border = tb
        if is_label:
            cell.font = w_bold
            cell.fill = dark_blue
        else:
            cell.fill = light_orange
            cell.alignment = Alignment(wrap_text=True)
        if c1 != c2:
            ws.merge_cells(f'{c1}{r}:{c2}{r}')
            # Apply border to all merged cells
            for col in range(openpyxl.utils.column_index_from_string(c1), openpyxl.utils.column_index_from_string(c2) + 1):
                ws.cell(row=r, column=col).border = tb

    # Row 7: PO Number & Date
    set_cell(7, 'A', 'B', "PO Number:", True)
    set_cell(7, 'C', 'E', quote_no)
    set_cell(7, 'F', 'G', "PO Date:", True)
    set_cell(7, 'H', 'I', row1.get("Quote Date", ""))

    # Row 8: Vendor Name
    set_cell(8, 'A', 'B', "Vendor Name:", True)
    set_cell(8, 'C', 'I', vendor_name)

    # Row 9: Vendor Address
    set_cell(9, 'A', 'B', "Vendor Address:", True)
    full_addr = v_addr
    if v_city: full_addr += f", {v_city}"
    set_cell(9, 'C', 'I', full_addr)

    # Row 10: Vendor GSTIN & Email
    set_cell(10, 'A', 'B', "Vendor GSTIN:", True)
    set_cell(10, 'C', 'E', v_gst)
    set_cell(10, 'F', 'G', "Vendor Email:", True)
    set_cell(10, 'H', 'I', v_email)

    # Row 11: Vendor No & Payment Terms
    set_cell(11, 'A', 'B', "Vendor No:", True)
    set_cell(11, 'C', 'E', str(row1.get("Vendor No", "")))
    set_cell(11, 'F', 'G', "Payment Terms:", True)
    set_cell(11, 'H', 'I', str(row1.get("Credit Days", "")))

    # Row 12: Location & Payment Method
    set_cell(12, 'A', 'B', "Location:", True)
    set_cell(12, 'C', 'E', loc_code)
    set_cell(12, 'F', 'G', "Payment Method:", True)
    set_cell(12, 'H', 'I', str(row1.get("Payment Method", "")))

    # Row 13: Mode of Delivery & Delivery Date
    md = str(row1.get("Mode of Delivery", ""))
    c_details = str(row1.get("Courier Details", ""))
    if c_details: md += f" ({c_details})"
    set_cell(13, 'A', 'B', "Mode of Delivery:", True)
    set_cell(13, 'C', 'E', md)
    set_cell(13, 'F', 'G', "Delivery Date:", True)
    set_cell(13, 'H', 'I', str(row1.get("Expected Delivery Date", "")))

    # ── Delivery Address Box ──
    ws.merge_cells('A15:I15'); ws['A15'] = "Delivery Address"; ws['A15'].fill = dark_blue; ws['A15'].font = w_bold
    for col in range(1, 10): ws.cell(row=15, column=col).border = tb
    
    ws.merge_cells('A16:I16'); ws['A16'] = loc_info.get("Address", "")
    for col in range(1, 10): ws.cell(row=16, column=col).border = tb
    
    ws.merge_cells('A17:I17'); ws['A17'] = f"Contact: {loc_info.get('Contact', '')}"
    for col in range(1, 10): ws.cell(row=17, column=col).border = tb
        
    headers = ["S.No", "ERP Code", "Vendor Item No.", "Product Description", "Qty", "Price (Excl.)", "GST %", "Total (Incl. GST)", "Remarks"]
    for i, h in enumerate(headers, 1): c = ws.cell(row=19, column=i, value=h); c.fill = dark_red; c.font = w_bold; c.alignment = c_align; c.border = tb
        
    r = 20; sub = 0.0
    for _, l in q_df.iterrows():
        ws.cell(row=r, column=1, value=l.get("S.No", "")).border = tb; ws.cell(row=r, column=2, value=l.get("ERP Code", "")).border = tb
        ws.cell(row=r, column=3, value=l.get("Vendor Item No", "")).border = tb; ws.cell(row=r, column=4, value=l.get("Product Description", "")).border = tb
        qty = float(l.get("Qty", 0) or 0); pb = float(l.get("Price Before GST", 0) or 0)
        pi = float(l.get("Price Inc. GST", 0) or 0); gstp = float(l.get("GST %", 0) or 0)
        line_total = qty * pi
        sub += (qty * pb)
        ws.cell(row=r, column=5, value=qty).border = tb; ws.cell(row=r, column=6, value=pb).border = tb
        ws.cell(row=r, column=7, value=gstp).border = tb; ws.cell(row=r, column=8, value=line_total).border = tb; ws.cell(row=r, column=9, value=l.get("Remarks", "")).border = tb
        for c in range(1, 10): ws.cell(row=r, column=c).fill = light_orange
        r += 1
        
    while r <= 30:
        for c in range(1, 10): ws.cell(row=r, column=c).border = tb
        r += 1
        
    fr = float(row1.get("Freight Charge", 0) or 0)
    ti = sum([float(l.get("Price Inc. GST", 0) or 0) * float(l.get("Qty", 0) or 0) for _, l in q_df.iterrows()])
    ws.cell(row=r+1, column=7, value="Freight Charge").font = b_bold; ws.cell(row=r+1, column=8, value=fr).border = tb
    ws.cell(row=r+2, column=7, value="GST Total").font = b_bold; ws.cell(row=r+2, column=8, value=ti - sub).border = tb
    ws.cell(row=r+3, column=7, value="Total Order Value").font = b_bold; ws.cell(row=r+3, column=8, value=ti + fr).border = tb
    
    ws.column_dimensions['A'].width = 15; ws.column_dimensions['B'].width = 15; ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 40; ws.column_dimensions['F'].width = 15; ws.column_dimensions['H'].width = 15; ws.column_dimensions['G'].width = 15
    
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── PO PDF Export ───────────────────────────────────────────
def export_po_pdf(quote_no: str, df: pd.DataFrame) -> io.BytesIO:
    try:
        from fpdf import FPDF
    except ImportError:
        return None

    q_df = df[df["Quote No"] == quote_no]
    if q_df.empty:
        return None
    row1 = q_df.iloc[0]
    loc_code = str(row1.get("Location Code", ""))
    loc_info = LOCATION_DETAILS.get(loc_code, {})

    class PDF(FPDF):
        def header(self):
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if os.path.exists(logo_path):
                self.image(logo_path, 10, 8, 45)
            self.set_font("helvetica", "B", 14)
            self.set_xy(60, 10)
            self.cell(0, 8, "SUPREME COMPUTERS INDIA PVT. LTD.", 0, 1, "C")
            self.set_font("helvetica", "", 9)
            self.set_x(60)
            self.cell(0, 5, "18/18, Majestic Plaza, Narasingapuram Street, Mount Road, Chennai - 600002", 0, 1, "C")
            gstin_val = loc_info.get("GSTIN", "33AAOCS1408H1ZR")
            self.set_x(60)
            self.cell(0, 5, f"GSTIN: {gstin_val}  |  Tel: 044-42082020", 0, 1, "C")
            self.ln(4)
            self.set_fill_color(31, 73, 125)
            self.set_text_color(255, 255, 255)
            self.set_font("helvetica", "B", 12)
            self.cell(0, 8, "PURCHASE ORDER", 0, 1, "C", True)
            self.set_text_color(0, 0, 0)
            self.ln(3)

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("helvetica", "", 9)

    # PO & Vendor Details
    vendor_name = str(row1.get("Vendor Name", ""))
    v_addr = str(row1.get("Vendor Address", ""))
    v_city = str(row1.get("Vendor City", ""))
    v_gst = str(row1.get("Vendor GST", ""))
    v_email = str(row1.get("Vendor Email", ""))
    
    if (not v_addr or v_addr == "nan") and vendor_name:
        vendor_df = load_master_sheet("Vendor Master")
        if not vendor_df.empty:
            vrow = vendor_df[vendor_df["Name"] == vendor_name] if "Name" in vendor_df.columns else vendor_df[vendor_df["name"] == vendor_name]
            if not vrow.empty:
                vr = vrow.iloc[0]
                v_addr = str(vr.get("address", vr.get("Address", "")))
                v_city = str(vr.get("city", vr.get("City", "")))
                v_gst = str(vr.get("vatRegistrationNo", vr.get("GST", vr.get("GSTIN", ""))))
                v_email = str(vr.get("email", vr.get("Email", vr.get("E-Mail", ""))))
                
    v_addr = v_addr if v_addr != "nan" else ""
    v_city = v_city if v_city != "nan" else ""
    v_gst = v_gst if v_gst != "nan" else ""
    v_email = v_email if v_email != "nan" else ""
    
    pdf.set_fill_color(189, 215, 238)
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "PO Number:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(quote_no), 1, 0, "L")
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "PO Date:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Quote Date", "")), 1, 1, "L")

    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Vendor Name:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(155, 7, vendor_name, 1, 1, "L")

    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Vendor Address:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    full_addr = v_addr
    if v_city: full_addr += f", {v_city}"
    pdf.cell(155, 7, full_addr, 1, 1, "L")

    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Vendor GSTIN:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, v_gst, 1, 0, "L")
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Vendor Email:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, v_email, 1, 1, "L")

    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Vendor No:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Vendor No", "")), 1, 0, "L")
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Payment Terms:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Credit Days", "")), 1, 1, "L")

    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Location:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, loc_code, 1, 0, "L")
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Payment Method:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Payment Method", "")), 1, 1, "L")
    
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Mode of Delivery:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Mode of Delivery", "")), 1, 0, "L")
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 7, "Delivery Date:", 1, 0, "L", True)
    pdf.set_font("helvetica", "", 9)
    pdf.cell(60, 7, str(row1.get("Expected Delivery Date", "")), 1, 1, "L")
    
    c_details = str(row1.get("Courier Details", ""))
    if c_details:
        pdf.set_font("helvetica", "B", 9)
        pdf.cell(35, 7, "Courier/Transport:", 1, 0, "L", True)
        pdf.set_font("helvetica", "", 9)
        pdf.cell(155, 7, c_details, 1, 1, "L")
        
    pdf.ln(3)

    # Delivery Address
    pdf.set_fill_color(31, 73, 125)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(0, 7, "Delivery Address", 1, 1, "L", True)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 9)
    addr_text = loc_info.get("Address", "")
    contact_text = loc_info.get("Contact", "")
    pdf.multi_cell(0, 6, f"{addr_text}\nContact: {contact_text}", 1)
    pdf.ln(3)

    # Line Items Table Header
    pdf.set_fill_color(192, 0, 0)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 8)
    col_w = [10, 25, 25, 60, 10, 22, 10, 28] # Adjusted widths
    headers = ["S.No", "ERP Code", "Vendor Item", "Description", "Qty", "Price(Excl.)", "GST%", "Total(Incl.)"]
    for i, h in enumerate(headers):
        pdf.cell(col_w[i], 7, h, 1, 0, "C", True)
    pdf.ln()

    # Rows
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 8)
    grand_total = 0.0
    sub_total = 0.0
    for _, l in q_df.iterrows():
        qty = float(l.get("Qty", 0) or 0)
        pb = float(l.get("Price Before GST", 0) or 0)
        pi = float(l.get("Price Inc. GST", 0) or 0)
        row_total = qty * pi
        grand_total += row_total
        sub_total += qty * pb
        desc = str(l.get("Product Description", ""))[:38]
        erp = str(l.get("ERP Code", ""))[:14]
        vin = str(l.get("Vendor Item No", ""))[:14]
        pdf.cell(col_w[0], 6, str(l.get("S.No", "")), 1, 0, "C")
        pdf.cell(col_w[1], 6, erp, 1, 0, "L")
        pdf.cell(col_w[2], 6, vin, 1, 0, "L")
        pdf.cell(col_w[3], 6, desc, 1, 0, "L")
        pdf.cell(col_w[4], 6, str(int(qty)), 1, 0, "C")
        pdf.cell(col_w[5], 6, f"{pb:,.2f}", 1, 0, "R")
        pdf.cell(col_w[6], 6, str(l.get("GST %", "")), 1, 0, "C")
        pdf.cell(col_w[7], 6, f"{row_total:,.2f}", 1, 1, "R")

    # Totals
    fr = float(row1.get("Freight Charge", 0) or 0)
    gst_total = grand_total - sub_total
    total_cols_w = sum(col_w[:-1])
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(total_cols_w, 7, "Sub Total (Excl. GST):", 0, 0, "R")
    pdf.cell(col_w[-1], 7, f"{sub_total:,.2f}", 1, 1, "R")
    pdf.cell(total_cols_w, 7, "Total GST:", 0, 0, "R")
    pdf.cell(col_w[-1], 7, f"{gst_total:,.2f}", 1, 1, "R")
    pdf.cell(total_cols_w, 7, "Freight Charge:", 0, 0, "R")
    pdf.cell(col_w[-1], 7, f"{fr:,.2f}", 1, 1, "R")

    pdf.set_fill_color(31, 73, 125)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(total_cols_w, 8, "GRAND TOTAL (Incl. GST):  ", 0, 0, "R", True)
    pdf.cell(col_w[-1], 8, f"{grand_total + fr:,.2f}", 1, 1, "R", True)
    pdf.set_text_color(0, 0, 0)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf

# ── DASHBOARD ───────────────────────────────────────────────
def page_dashboard():
    # Full View Logo at the top
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass
    
    st.markdown('<div class="section-header">📊 Executive Dashboard</div>', unsafe_allow_html=True)
    df = load_quotes()

    total   = len(df["Quote No"].dropna().unique()) if not df.empty else 0
    pending = len(df[df["Status"] == "Pending Approval"]) if not df.empty else 0
    approved= len(df[df["Status"] == "Approved"]) if not df.empty else 0
    modified= len(df[df["Status"] == "Modified"]) if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Quotes", total)
    c2.metric("Pending Approval", pending)
    c3.metric("Approved", approved)
    c4.metric("Modified", modified)

    st.markdown("---")

    if not df.empty:
        st.subheader("Recent Quotes")
        
        # Calculate correct Total Order Value
        if "Freight Charge" not in df.columns:
            df["Freight Charge"] = 0.0
            
        # Create temporary columns for proper Line-Item multiplication (Qty * Unit Price Inc GST)
        df_temp = df.copy()
        df_temp["Qty_num"] = pd.to_numeric(df_temp["Qty"], errors="coerce").fillna(0.0)
        df_temp["Price_Inc_num"] = pd.to_numeric(df_temp["Price Inc. GST"], errors="coerce").fillna(0.0)
        df_temp["Line_Total_Inc"] = df_temp["Qty_num"] * df_temp["Price_Inc_num"]
        
        quote_totals = df_temp.groupby("Quote No").agg(
            total_items_inc_gst=("Line_Total_Inc", "sum"),
            freight=("Freight Charge", lambda x: pd.to_numeric(x.iloc[0], errors="coerce") if not x.empty else 0.0)
        ).reset_index()
        quote_totals["freight"] = quote_totals["freight"].fillna(0.0)
        quote_totals["Total Order Value"] = quote_totals["total_items_inc_gst"] + quote_totals["freight"]
        
        recent = df.drop_duplicates("Quote No", keep="last").sort_values("Created Date", ascending=False).head(10)
        recent = recent.merge(quote_totals[["Quote No", "Total Order Value"]], on="Quote No", how="left")
        
        disp = recent[["Quote No", "Purchase Quote raised by", "Vendor Name", "Status", "Total Order Value", "Created Date"]].copy()
        
        # Formatting
        disp["Total Order Value"] = disp["Total Order Value"].apply(lambda x: f"₹{x:,.2f}" if pd.notnull(x) else "₹0.00")
        disp = disp.astype({c: str for c in ["Quote No", "Purchase Quote raised by", "Vendor Name", "Status", "Created Date"]})
        st.dataframe(disp, width='stretch')

        st.markdown("---")
        st.subheader("Quotes by Status")
        status_counts = df.drop_duplicates("Quote No", keep="last")["Status"].value_counts()
        st.bar_chart(status_counts)
    else:
        st.info("No quotes found. Create your first quote!")

# ── CREATE / EDIT QUOTE ──────────────────────────────────────────
def page_create_quote():
    st.markdown('<div class="section-header">➕ Create / Edit Purchase Quote</div>', unsafe_allow_html=True)

    quotes_df = load_quotes()
    
    # ── Edit Mode Selection ──
    edit_mode = False
    sel_quote_to_edit = "-- New Quote --"
    if not quotes_df.empty:
        existing_quotes = sorted(quotes_df["Quote No"].dropna().unique().tolist(), reverse=True)
        sel_quote_to_edit = st.selectbox("📝 Create New or Edit Existing Quote", ["-- New Quote --"] + existing_quotes)
        if sel_quote_to_edit != "-- New Quote --":
            edit_mode = True

    # ── Logic to load existing data if in edit mode ──
    if edit_mode and ("last_loaded_quote" not in st.session_state or st.session_state.last_loaded_quote != sel_quote_to_edit):
        q_data = quotes_df[quotes_df["Quote No"] == sel_quote_to_edit]
        if not q_data.empty:
            first_row = q_data.iloc[0]
            st.session_state.edit_salesperson = str(first_row.get("Purchase Quote raised by", ""))
            st.session_state.edit_vendor = str(first_row.get("Vendor Name", ""))
            st.session_state.edit_location = str(first_row.get("Location Code", "EGMR"))
            st.session_state.edit_status = str(first_row.get("Status", "Pending Approval"))
            st.session_state.edit_payment_method = str(first_row.get("Payment Method", ""))
            st.session_state.edit_credit_days = str(first_row.get("Credit Days", ""))
            st.session_state.edit_mode_of_delivery = str(first_row.get("Mode of Delivery", ""))
            st.session_state.edit_courier_details = str(first_row.get("Courier Details", ""))
            
            # Load Expected Delivery Date
            ed_val = first_row.get("Expected Delivery Date", None)
            try:
                if pd.notnull(ed_val) and str(ed_val).strip():
                    st.session_state.edit_expected_delivery = pd.to_datetime(ed_val).date()
                else:
                    st.session_state.edit_expected_delivery = None
            except:
                st.session_state.edit_expected_delivery = None
            
            # Load line items
            new_lines = []
            for _, row in q_data.iterrows():
                new_lines.append({
                    "erp_code": str(row.get("ERP Code", "")),
                    "vendor_item_no": str(row.get("Vendor Item No", "")),
                    "description": str(row.get("Product Description", "")),
                    "qty": int(row.get("Qty", 1)),
                    "price_before_gst": float(row.get("Price Before GST", 0.0)),
                    "gst_percent": float(row.get("GST %", 18.0)),
                    "remarks": str(row.get("Remarks", ""))
                })
            st.session_state.line_items = new_lines
            st.session_state.last_loaded_quote = sel_quote_to_edit
    elif not edit_mode:
        if "last_loaded_quote" in st.session_state:
            del st.session_state.last_loaded_quote
            st.session_state.line_items = []
            # Reset other edit fields if needed

    quote_no = sel_quote_to_edit if edit_mode else next_quote_number(quotes_df)

    # ── Header section ───────────────────────────────────────
    st.markdown('<div class="erp-card">', unsafe_allow_html=True)
    st.subheader("Quote Header")
    h1, h2, h3, h4 = st.columns(4)
    with h1:
        st.text_input("Quote Number", value=quote_no, disabled=True)
    with h2:
        quote_date = st.date_input("Quote Date", value=date.today())
    with h3:
        location_opts = sorted(list(LOCATION_DETAILS.keys()))
        def_loc = st.session_state.get("edit_location", "EGMR")
        loc_idx = location_opts.index(def_loc) if def_loc in location_opts else 0
        location_code = st.selectbox("Location Code", location_opts, index=loc_idx)
        
        # Display selected location details briefly
        loc_info = LOCATION_DETAILS.get(location_code, {})
        if loc_info:
            st.caption(f"📍 {loc_info['Address'][:50]}...")
    with h4:
        salesperson = st.text_input("Purchase Quote raised by", value=st.session_state.get("edit_salesperson", ""), placeholder="Enter your name")

    # Vendor selection
    st.markdown("#### Vendor")
    vendor_names = []
    if not vendor_df.empty:
        for nc in ["name", "Name", "searchName", "SearchName"]:
            if nc in vendor_df.columns:
                vendor_names = vendor_df[nc].dropna().unique().tolist()
                break

    v1, v2, v3, v4 = st.columns(4)
    with v1:
        def_v = st.session_state.get("edit_vendor", "-- Select Vendor --")
        v_idx = vendor_names.index(def_v) + 1 if def_v in vendor_names else 0
        selected_vendor = st.selectbox("Vendor Name", ["-- Select Vendor --"] + vendor_names, index=v_idx)
    
    vendor_no = ""
    vendor_addr = ""
    vendor_city = ""
    vendor_gst = ""
    vendor_email = ""
    v_pay_meth = ""
    v_cred_days = ""
    
    if selected_vendor != "-- Select Vendor --" and not vendor_df.empty:
        vrow = vendor_df[vendor_df["Name"] == selected_vendor] if "Name" in vendor_df.columns else vendor_df[vendor_df["name"] == selected_vendor]
        if not vrow.empty:
            vr = vrow.iloc[0]
            vendor_no = str(vr.get("number", vr.get("vendorNo", "")))
            vendor_addr = str(vr.get("address", vr.get("Address", "")))
            vendor_city = str(vr.get("city", vr.get("City", "")))
            vendor_gst = str(vr.get("vatRegistrationNo", vr.get("GST", vr.get("GSTIN", ""))))
            vendor_email = str(vr.get("email", vr.get("Email", vr.get("E-Mail", ""))))
            v_pay_meth = str(vr.get("paymentMethodCode", ""))
            v_cred_days = str(vr.get("paymentTermsCode", ""))

    with v2:
        st.text_input("Vendor No", value=vendor_no, disabled=True)
    with v3:
        st.text_input("Vendor Address", value=vendor_addr, disabled=True)
    with v4:
        st.text_input("City", value=vendor_city, disabled=True)
    
    if vendor_gst:
        st.caption(f"🛡️ Vendor GST: {vendor_gst}")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Payment & Delivery ───────────────────────────────────
    st.markdown('<div class="erp-card">', unsafe_allow_html=True)
    st.subheader("Payment & Delivery Options")
    
    p1, p2, p3 = st.columns(3)
    pm_opts = ["", "Cash", "Online Transfer", "CDC", "PDC"]
    
    # Use existing data if editing, otherwise vendor default
    if edit_mode:
        curr_pm = st.session_state.get("edit_payment_method", "")
        pm_idx = pm_opts.index(curr_pm) if curr_pm in pm_opts else 0
    else:
        pm_idx = pm_opts.index(v_pay_meth) if v_pay_meth in pm_opts else 0
        
    with p1: 
        payment_method = st.selectbox("Payment Method", pm_opts, index=pm_idx)
    with p2: 
        credit_days = st.text_input("Credit Days", value=st.session_state.get("edit_credit_days", v_cred_days))
    
    md_opts = ["", "Door Delivery", "Pickup", "Courier"]
    def_md = st.session_state.get("edit_mode_of_delivery", "")
    md_idx = md_opts.index(def_md) if def_md in md_opts else 0
    with p3: 
        mode_of_delivery = st.selectbox("Mode of Delivery", md_opts, index=md_idx)
        
    d1, d2 = st.columns(2)
    with d1: 
        def_ed = st.session_state.get("edit_expected_delivery", None)
        expected_delivery = st.date_input("Expected Delivery Date", value=def_ed)
    with d2:
        courier_details = st.text_input("Courier / Transport Details", value=st.session_state.get("edit_courier_details", ""))
        
    cash = "Yes" if payment_method == "Cash" else ""
    online_transfer = "Yes" if payment_method == "Online Transfer" else ""
    cdc = "Yes" if payment_method == "CDC" else ""
    pdc = "Yes" if payment_method == "PDC" else ""
    
    door_delivery = "Yes" if mode_of_delivery == "Door Delivery" else ""
    pickup = "Yes" if mode_of_delivery == "Pickup" else ""
    courier = "Yes" if mode_of_delivery == "Courier" else ""
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Line Items ───────────────────────────────────────────
    st.markdown('<div class="erp-card">', unsafe_allow_html=True)
    st.subheader("Line Items")

    erp_codes = []
    if not item_df.empty:
        for nc in ["No", "Item_No", "ERP_Code"]:
            if nc in item_df.columns:
                erp_codes = item_df[nc].dropna().unique().tolist()
                break

    if "line_items" not in st.session_state:
        st.session_state.line_items = []

    col_add, col_clear = st.columns([1, 5])
    with col_add:
        if st.button("➕ Add Line"):
            st.session_state.line_items.append({
                "erp_code": "", 
                "vendor_item_no": "", 
                "description": "", 
                "qty": None, 
                "price_before_gst": None,
                "gst_percent": 18.0,
                "remarks": ""
            })
    with col_clear:
        if st.button("🗑️ Clear All Lines"):
            st.session_state.line_items = []

    line_data_for_saving = []
    items_to_remove = []

    # UI Table Header
    h_col0, h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7, h_col8, h_col9 = st.columns([0.5, 1.5, 1.5, 2.5, 1, 1.2, 1, 1.2, 1.5, 0.5])
    h_col0.write("**S.No**")
    h_col1.write("**ERP Code**")
    h_col2.write("**Vendor Item No**")
    h_col3.write("**Product Description**")
    h_col4.write("**Qty**")
    h_col5.write("**Price Before GST**")
    h_col6.write("**GST %**")
    h_col7.write("**Total (Incl. GST)**")
    h_col8.write("**Remarks**")
    h_col9.write("")

    for i, line in enumerate(st.session_state.line_items):
        c0, c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns([0.5, 1.5, 1.5, 2.5, 1, 1.2, 1, 1.2, 1.5, 0.5])
        
        with c0:
            st.write(f"{i+1}")
            
        with c1:
            # Using ERP Code as the main selector
            sel_erp = st.selectbox(f"Select ERP #{i+1}", ["-- Select --"] + erp_codes, 
                                    index=erp_codes.index(line["erp_code"]) + 1 if line["erp_code"] in erp_codes else 0,
                                    key=f"erp_sel_{i}", label_visibility="collapsed")
        
        description = line["description"]
        vendor_item_no = line["vendor_item_no"]
        item_price = line["price_before_gst"]
        
        if sel_erp != "-- Select --" and not item_df.empty:
            for nc in ["No", "Item_No", "ERP_Code"]:
                if nc in item_df.columns:
                    irow = item_df[item_df[nc] == sel_erp]
                    if not irow.empty:
                        # Auto-fill logic
                        if sel_erp != line["erp_code"]:
                            for dc in ["Description", "Search_Description", "Product Description", "Product_Description", "Item Description", "Name"]:
                                if dc in irow.columns: description = str(irow.iloc[0][dc]); break
                            for vc in ["Vendor_Item_No", "VendorItemNo", "Vendor Item No", "Vendor_No"]:
                                if vc in irow.columns: vendor_item_no = str(irow.iloc[0][vc]); break
                            for pc in ["Quoting_Price_WIN", "Unit_Price", "Last_Direct_Cost"]:
                                if pc in irow.columns:
                                    try: item_price = float(irow.iloc[0][pc])
                                    except: pass
                                    break
                            # Update session state correctly for widgets
                            st.session_state.line_items[i]["erp_code"] = sel_erp
                            st.session_state.line_items[i]["vendor_item_no"] = vendor_item_no
                            # We keep price as None if the user wants to type it, or auto-fill if preferred.
                            # The user said "I will type", so we'll leave it blank even after selection 
                            # unless they specifically wanted auto-fill back. 
                            # For now, let's keep it manual as requested.
                            st.session_state.line_items[i]["price_before_gst"] = None
                            st.session_state.line_items[i]["description"] = description
                            st.rerun()
                        break
        
        with c2:
            st.text_input(f"VItem {i}", value=vendor_item_no, disabled=True, label_visibility="collapsed")
        with c3:
            st.text_input(f"Desc {i}", value=description, disabled=True, label_visibility="collapsed")
        with c4:
            qty_val = line["qty"]
            qty = st.number_input(f"Qty {i}", min_value=0, value=qty_val, step=1, key=f"qty_{i}", label_visibility="collapsed")
        with c5:
            price_val = line["price_before_gst"]
            price = st.number_input(f"Price {i}", min_value=0.0, value=price_val, step=0.01, key=f"price_{i}", label_visibility="collapsed")
        with c6:
            gst_p = st.number_input(f"GST% {i}", min_value=0.0, max_value=100.0, value=float(line["gst_percent"]), step=0.1, key=f"gstp_{i}", label_visibility="collapsed")
        
        # Handle blank (None) values for calculations
        safe_price = price if price is not None else 0.0
        safe_qty = qty if qty is not None else 0
        
        price_inc_gst = safe_price * (1 + gst_p/100)
        line_total_before_gst = safe_qty * safe_price
        line_total_inc_gst = safe_qty * price_inc_gst
        
        with c7:
            st.text_input(f"TotalIncGST {i}", value=f"{line_total_inc_gst:,.2f}", disabled=True, label_visibility="collapsed")
        with c8:
            remarks = st.text_input(f"Rem {i}", value=line["remarks"], key=f"rem_{i}", label_visibility="collapsed")
        with c9:
            if st.button("❌", key=f"remove_{i}"):
                items_to_remove.append(i)

        # Keep session state in sync
        st.session_state.line_items[i] = {
            "erp_code": sel_erp,
            "vendor_item_no": vendor_item_no,
            "description": description,
            "qty": qty,
            "price_before_gst": price,
            "gst_percent": gst_p,
            "remarks": remarks
        }
        
        if sel_erp != "-- Select --":
            line_data_for_saving.append({
                "S.No": i + 1,
                "ERP Code": sel_erp,
                "Vendor Item No": vendor_item_no,
                "Product Description": description,
                "Qty": safe_qty,
                "Price Before GST": safe_price,
                "GST %": gst_p,
                "Price Inc. GST": price_inc_gst,
                "Remarks": remarks,
                "line_total_before_gst": line_total_before_gst,
                "line_total_inc_gst": safe_qty * price_inc_gst
            })

    for idx in reversed(items_to_remove):
        st.session_state.line_items.pop(idx)
        st.rerun()

    # ── Totals ───────────────────────────────────────────────
    subtotal_before_gst = sum(l["line_total_before_gst"] for l in line_data_for_saving)
    total_inc_gst = sum(l["line_total_inc_gst"] for l in line_data_for_saving)
    gst_amt = total_inc_gst - subtotal_before_gst

    st.markdown("---")
    f_col1, f_col2 = st.columns([3, 1])
    with f_col2:
        freight_charge = st.number_input("Freight Charge", min_value=0.0, value=0.0, step=0.01)
    
    total_order_value = total_inc_gst + freight_charge

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Sub Total (Excl. GST)", f"₹{subtotal_before_gst:,.2f}")
    t2.metric("Total GST", f"₹{gst_amt:,.2f}")
    t3.metric("Freight Charge", f"₹{freight_charge:,.2f}")
    t4.metric("Total Order Value", f"₹{total_order_value:,.2f}")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Submit ───────────────────────────────────────────────
    st.markdown("---")
    sb1, sb2 = st.columns([1, 5])
    with sb1:
        submit = st.button("💾 Submit Quote", use_container_width=True)

    if submit:
        errors = []
        if not salesperson.strip():
            errors.append("Salesperson Name is required.")
        if selected_vendor == "-- Select Vendor --":
            errors.append("Please select a Vendor.")
        if not line_data_for_saving:
            errors.append("Add at least one valid line item.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            final_rows = []
            for l in line_data_for_saving:
                final_rows.append({
                    "Quote No": quote_no,
                    "Quote Date": str(quote_date),
                    "Location Code": location_code,
                    "Purchase Quote raised by": salesperson.strip(),
                    "Vendor Name": selected_vendor,
                    "Vendor No": vendor_no,
                    "Vendor Address": vendor_addr,
                    "Vendor City": vendor_city,
                    "Vendor GST": vendor_gst,
                    "Vendor Email": vendor_email,
                    "Payment Method": payment_method,
                    "Cash": cash,
                    "Online Transfer": online_transfer,
                    "CDC": cdc,
                    "PDC": pdc,
                    "Credit Days": credit_days,
                    "Mode of Delivery": mode_of_delivery,
                    "Door Delivery": door_delivery,
                    "Pickup": pickup,
                    "Courier": courier,
                    "Expected Delivery Date": str(expected_delivery) if expected_delivery else "",
                    "Courier Details": courier_details,
                    "S.No": l["S.No"],
                    "ERP Code": l["ERP Code"],
                    "Vendor Item No": l["Vendor Item No"],
                    "Product Description": l["Product Description"],
                    "Qty": l["Qty"],
                    "Price Before GST": l["Price Before GST"],
                    "GST %": l["GST %"],
                    "Price Inc. GST": l["Price Inc. GST"],
                    "Total (Incl. GST)": l["line_total_inc_gst"],
                    "Remarks": l["Remarks"],
                    "Freight Charge": freight_charge,
                    "Status": "Pending Approval",
                    "Created Date": now,
                    "Modified Date": "",
                })
            
            if final_rows:
                new_df = pd.DataFrame(final_rows)
                # Align columns to REQUIRED_COLS to prevent misaligned database inserts
                for col in REQUIRED_COLS:
                    if col not in new_df.columns:
                        new_df[col] = ""
                new_df = new_df[REQUIRED_COLS]
                
                if save_quotes(new_df, is_edit=edit_mode, quote_no=quote_no):
                    st.success(f"✅ Quote **{quote_no}** {'updated' if edit_mode else 'submitted'} successfully!")
                    # Clear session state if it was a new quote, keep it if it was an edit so buttons are visible?
                    # Actually, better to keep it so they can see the Email/Download buttons that appear below.
                    st.cache_data.clear()
                    # st.rerun() # Don't rerun immediately so they can see the success message and use buttons

        # ── Post-Submit / Edit Actions (Email & Download) ────────
    if edit_mode:
        st.markdown("---")
        st.subheader(f"Actions for Quote: {quote_no}")
        
        # We need the current rows to generate the email/PO
        q_rows_for_actions = quotes_df[quotes_df["Quote No"] == quote_no]
        
        c_dl, c_mail, c_status = st.columns([1, 1.5, 1.5])
        
        with c_dl:
            # 1. Download Actions
            st.write("**📥 Export PO**")
            # Excel
            buf_xl = export_po_excel(quote_no, quotes_df)
            if buf_xl:
                st.download_button("Excel PO", buf_xl, file_name=f"PO_{quote_no}.xlsx", 
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 use_container_width=True)
            # PDF
            buf_pdf = export_po_pdf(quote_no, quotes_df)
            if buf_pdf:
                st.download_button("PDF PO", buf_pdf, file_name=f"PO_{quote_no}.pdf", 
                                 mime="application/pdf",
                                 use_container_width=True)
        
        with c_mail:
            # 2. Email Configuration
            st.write("**✉️ Email Submission**")
            total_qty_em = q_rows_for_actions["Qty"].fillna(0).sum()
            raw_fr_em = q_rows_for_actions["Freight Charge"].iloc[0] if "Freight Charge" in q_rows_for_actions.columns else 0
            fr_em = float(raw_fr_em) if pd.notnull(raw_fr_em) else 0.0
            line_totals = q_rows_for_actions["Qty"].fillna(0) * q_rows_for_actions["Price Inc. GST"].fillna(0)
            total_amt_em = line_totals.sum() + fr_em
            
            # Construct HTML table for email
            html_table = f"""
            <table style="border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 13px; border: 1px solid black;">
                <tr style="background-color: #BDD7EE; color: black; font-weight: bold; text-align: center;">
                    <td style="border: 1px solid black; padding: 5px;">S.No</td>
                    <td style="border: 1px solid black; padding: 5px;">Vendor Item No</td>
                    <td style="border: 1px solid black; padding: 5px;">Product Description</td>
                    <td style="border: 1px solid black; padding: 5px;">Qty</td>
                    <td style="border: 1px solid black; padding: 5px;">Price (Excl.)</td>
                    <td style="border: 1px solid black; padding: 5px;">GST %</td>
                    <td style="border: 1px solid black; padding: 5px;">Total (Incl. GST)</td>
                </tr>
            """
            for _, r in q_rows_for_actions.iterrows():
                qty = float(r['Qty']) if pd.notnull(r['Qty']) else 0.0
                p_b = float(r['Price Before GST']) if pd.notnull(r['Price Before GST']) else 0.0
                g_p = float(r['GST %']) if pd.notnull(r['GST %']) else 0.0
                p_i = float(r['Price Inc. GST']) if pd.notnull(r['Price Inc. GST']) else 0.0
                row_total = qty * p_i
                html_table += f"<tr><td style='border: 1px solid black; padding: 5px; text-align: center;'>{r['S.No']}</td><td style='border: 1px solid black; padding: 5px;'>{r['Vendor Item No']}</td><td style='border: 1px solid black; padding: 5px;'>{r['Product Description']}</td><td style='border: 1px solid black; padding: 5px; text-align: center;'>{int(qty)}</td><td style='border: 1px solid black; padding: 5px; text-align: right;'>{p_b:,.2f}</td><td style='border: 1px solid black; padding: 5px; text-align: center;'>{g_p}</td><td style='border: 1px solid black; padding: 5px; text-align: right;'>{row_total:,.2f}</td></tr>"

            html_table += f"<tr style='font-weight: bold;'><td style='border: 1px solid black; padding: 5px;' colspan='3'></td><td style='border: 1px solid black; padding: 5px; text-align: center; background-color: #f2f2f2;'>{total_qty_em}</td><td style='border: 1px solid black; padding: 5px;'></td><td style='border: 1px solid black; padding: 5px; text-align: center; background-color: #E2EFDA;'>Total</td><td style='border: 1px solid black; padding: 5px; text-align: right; background-color: #f2f2f2;'>{total_amt_em:,.2f}</td></tr></table>"

            # Construct plain-text table for fallback
            text_table = "S.No\tVendor Item No\tProduct Description\tQty\tPrice (Excl.)\tGST %\tTotal (Incl.)\n"
            for _, r in q_rows_for_actions.iterrows():
                qty = float(r['Qty']) if pd.notnull(r['Qty']) else 0.0
                p_b = float(r['Price Before GST']) if pd.notnull(r['Price Before GST']) else 0.0
                g_p = float(r['GST %']) if pd.notnull(r['GST %']) else 0.0
                p_i = float(r['Price Inc. GST']) if pd.notnull(r['Price Inc. GST']) else 0.0
                row_total = qty * p_i
                text_table += f"{r['S.No']}\t{r['Vendor Item No']}\t{r['Product Description']}\t{int(qty)}\t{p_b:,.2f}\t{g_p}\t{row_total:,.2f}\n"

            def_greeting = f"Dear Sir/Madam,\n\nGreetings from Supreme Computers India Pvt. Ltd.\n\nPlease find below the Purchase Quote details for your reference:"
            def_closing = f"Total Qty: {total_qty_em}\nTotal Amount (Incl. GST): {total_amt_em:,.2f}\n\nKindly review and confirm. Please feel free to contact us for any clarification."
            
            email_greeting = st.text_area("Email Greeting & Intro", value=def_greeting, height=120)
            email_closing = st.text_area("Email Closing & Totals", value=def_closing, height=120)
            
            from email.message import EmailMessage
            subject_em = f"Purchase Quote Submission - {quote_no}"
            cc_emails = "purchase@supremeindia.com, mis3@supremeindia.com"
            
            msg = EmailMessage()
            msg['Subject'] = subject_em
            msg['Cc'] = cc_emails
            
            # Text Fallback
            full_text_body = f"{email_greeting}\n\n{text_table}\n{email_closing}"
            msg.set_content(full_text_body)
            
            # HTML Main Version
            html_body = f"""
            <html>
            <body style='font-family: Calibri, Arial, sans-serif;'>
                <p>{email_greeting.replace(chr(10), '<br>')}</p>
                {html_table}
                <p>{email_closing.replace(chr(10), '<br>')}</p>
            </body>
            </html>
            """
            msg.add_alternative(html_body, subtype='html')
            
            if buf_pdf:
                msg.add_attachment(buf_pdf.getvalue(), maintype='application', subtype='pdf', filename=f"PO_{quote_no}.pdf")
            
            st.download_button("🚀 Open Outlook Draft (with PO PDF)", msg.as_bytes(), file_name=f"Draft_{quote_no}.eml", mime="message/rfc822", use_container_width=True, type="primary")
            st.success("✅ **Draft Ready!** Click the blue button above to download the Outlook file. Once opened, your **PO PDF will be automatically attached** and the table will be perfectly formatted.")

        with c_status:
            # 3. Status Update
            st.write("**🚦 Progress Status**")
            current_st = q_rows_for_actions.iloc[0]["Status"] if not q_rows_for_actions.empty else "Pending Approval"
            st_opts = ["Pending Approval", "Modified", "Approved", "Rejected", "Converted to PO"]
            new_st = st.selectbox("Update Status", st_opts, index=st_opts.index(current_st) if current_st in st_opts else 0)
            
            if new_st != current_st:
                if st.button("Apply Status Update"):
                    quotes_df.loc[quotes_df["Quote No"] == quote_no, "Status"] = new_st
                    quotes_df.loc[quotes_df["Quote No"] == quote_no, "Modified Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_quotes(quotes_df)
                    st.success(f"Status updated to {new_st}")
                    st.cache_data.clear()
                    st.rerun()

        # ── Show Table for Preview ──
        st.markdown(html_table, unsafe_allow_html=True)

# ── VIEW QUOTES ──────────────────────────────────────────────
def page_view_quotes():
    st.markdown('<div class="section-header">🔍 View Quotes</div>', unsafe_allow_html=True)
    df = load_quotes()
    if df.empty:
        st.info("No quotes available yet.")
        return

    # Filters
    f1, f2, f3 = st.columns(3)
    with f1:
        sp_list = ["All"] + df["Purchase Quote raised by"].dropna().unique().tolist()
        sp_filter = st.selectbox("Filter by Purchaser", sp_list)
    with f2:
        st_list = ["All"] + df["Status"].dropna().unique().tolist()
        st_filter = st.selectbox("Filter by Status", st_list)
    with f3:
        date_filter = st.date_input("Filter by Date (Created)", value=None)

    filtered = df.copy()
    if sp_filter != "All":
        filtered = filtered[filtered["Purchase Quote raised by"] == sp_filter]
    if st_filter != "All":
        filtered = filtered[filtered["Status"] == st_filter]
    if date_filter:
        filtered["_date"] = pd.to_datetime(filtered["Created Date"], errors="coerce").dt.date
        filtered = filtered[filtered["_date"] == date_filter]
        filtered = filtered.drop(columns=["_date"], errors="ignore")

    st.markdown(f"**Showing {len(filtered)} records**")
    disp_f = filtered.copy()
    for c in disp_f.select_dtypes(include='object').columns:
        disp_f[c] = disp_f[c].astype(str)
    st.dataframe(disp_f, width='stretch')

    # Download
    buf = io.BytesIO()
    filtered.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button("⬇️ Download as Excel", buf, file_name="quotes_export.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=False)

# ── Routing ──────────────────────────────────────────────────
page = st.session_state.page
if page == "Dashboard":
    page_dashboard()
elif page == "Create / Edit Quote":
    page_create_quote()
elif page == "View Quotes":
    page_view_quotes()
elif page == "Manage Quotes":
    st.warning("The Manage menu has been moved inside the 'Create / Edit Quote' page. Please use the dropdown at the top of that page to select a quote to manage.")
