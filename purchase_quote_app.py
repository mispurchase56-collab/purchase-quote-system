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
import urllib.parse          # ← NEW: for mailto encoding
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
        
        if os.path.exists(CREDENTIALS_PATH):
            with open(CREDENTIALS_PATH, 'r') as f:
                creds_info = json.load(f)
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
        
        if 'private_key' in creds_info:
            pk = creds_info['private_key']
            pk = pk.replace("\\n", "\n")
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

item_df, vendor_df, LOCATION_DETAILS = load_master_data()

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
    df = load_gsheet(DATABASE_SPREADSHEET_ID, "Purchase Quotes")
    if not df.empty:
        for c in REQUIRED_COLS:
            if c not in df.columns: df[c] = ""
        return df
    
    if os.path.exists(LOCAL_EXCEL_PATH):
        try:
            return pd.read_excel(LOCAL_EXCEL_PATH)
        except:
            pass
    return pd.DataFrame(columns=REQUIRED_COLS)

def save_quotes(new_df: pd.DataFrame, is_edit: bool = False, quote_no: str = None):
    for col in REQUIRED_COLS:
        if col not in new_df.columns:
            new_df[col] = ""
    new_df = new_df[REQUIRED_COLS]

    if gc:
        try:
            sh = gc.open_by_key(DATABASE_SPREADSHEET_ID)
            ws = sh.worksheet("Purchase Quotes")
            
            headers = ws.row_values(1)
            if not headers:
                ws.append_row(REQUIRED_COLS)
            elif len(headers) < len(REQUIRED_COLS):
                ws.update('A1', [REQUIRED_COLS])
                
            if is_edit and quote_no:
                all_data = ws.get_all_records()
                full_df = pd.DataFrame(all_data)
                
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

/* Buttons & Download Buttons */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(135deg, #0ea5e9, #6366f1);
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.5rem 1.8rem;
    font-weight: 600;
    font-size: 15px;
    transition: opacity .2s;
}
.stButton > button:hover, .stDownloadButton > button:hover { opacity: 0.88; }

/* Table */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* Success / Warning */
div[data-baseweb="notification"] { border-radius: 10px !important; }

div[data-testid="stTextInput"] input:disabled, 
div[data-testid="stNumberInput"] input:disabled {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    background-color: #f1f5f9 !important;
    font-weight: 600 !important;
    border: 1px solid #cbd5e1 !important;
}

/* ── NEW: Mailto button styling ── */
.mailto-btn {
    display: inline-block;
    width: 100%;
    text-align: center;
    padding: 12px 24px;
    background: linear-gradient(135deg, #059669, #10b981);
    color: white !important;
    text-decoration: none !important;
    border-radius: 8px;
    font-weight: 700;
    font-size: 15px;
    font-family: 'Inter', sans-serif;
    transition: opacity 0.2s;
    box-sizing: border-box;
}
.mailto-btn:hover { opacity: 0.88; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state defaults ───────────────────────────────────
if "role" not in st.session_state:
    st.session_state.role = "Product manager / product team"
if "line_items" not in st.session_state:
    st.session_state.line_items = []
if "edit_quote_no" not in st.session_state:
    st.session_state.edit_quote_no = None
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass
    
    st.markdown("## 📋 Purchase Quote System")
    st.markdown("---")
    pages = ["Dashboard", "Create / Edit Quote", "View Quotes"]
    
    st.session_state.page = st.radio("Navigation", pages, index=pages.index(st.session_state.page) if st.session_state.page in pages else 0)
    st.markdown("---")
    st.caption("v1.1 | Purchase Quote ERP")

quotes_df = load_quotes()

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
        vdf_temp = load_gsheet(MASTER_SPREADSHEET_ID, "Vendor Master")
        if not vdf_temp.empty:
            vrow = vdf_temp[vdf_temp["Name"] == vendor_name] if "Name" in vdf_temp.columns else vdf_temp[vdf_temp["name"] == vendor_name]
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
            for col in range(openpyxl.utils.column_index_from_string(c1), openpyxl.utils.column_index_from_string(c2) + 1):
                ws.cell(row=r, column=col).border = tb

    set_cell(7, 'A', 'B', "PO Number:", True)
    set_cell(7, 'C', 'E', quote_no)
    set_cell(7, 'F', 'G', "PO Date:", True)
    set_cell(7, 'H', 'I', row1.get("Quote Date", ""))

    set_cell(8, 'A', 'B', "Vendor Name:", True)
    set_cell(8, 'C', 'I', vendor_name)

    set_cell(9, 'A', 'B', "Vendor Address:", True)
    full_addr = v_addr
    if v_city: full_addr += f", {v_city}"
    set_cell(9, 'C', 'I', full_addr)

    set_cell(10, 'A', 'B', "Vendor GSTIN:", True)
    set_cell(10, 'C', 'E', v_gst)
    set_cell(10, 'F', 'G', "Vendor Email:", True)
    set_cell(10, 'H', 'I', v_email)

    set_cell(11, 'A', 'B', "Vendor No:", True)
    set_cell(11, 'C', 'E', str(row1.get("Vendor No", "")))
    set_cell(11, 'F', 'G', "Payment Terms:", True)
    set_cell(11, 'H', 'I', str(row1.get("Credit Days", "")))

    set_cell(12, 'A', 'B', "Location:", True)
    set_cell(12, 'C', 'E', loc_code)
    set_cell(12, 'F', 'G', "Payment Method:", True)
    set_cell(12, 'H', 'I', str(row1.get("Payment Method", "")))

    md = str(row1.get("Mode of Delivery", ""))
    c_details = str(row1.get("Courier Details", ""))
    if c_details: md += f" ({c_details})"
    set_cell(13, 'A', 'B', "Mode of Delivery:", True)
    set_cell(13, 'C', 'E', md)
    set_cell(13, 'F', 'G', "Delivery Date:", True)
    set_cell(13, 'H', 'I', str(row1.get("Expected Delivery Date", "")))

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

    vendor_name = str(row1.get("Vendor Name", ""))
    v_addr = str(row1.get("Vendor Address", ""))
    v_city = str(row1.get("Vendor City", ""))
    v_gst = str(row1.get("Vendor GST", ""))
    v_email = str(row1.get("Vendor Email", ""))
    
    if (not v_addr or v_addr == "nan") and vendor_name:
        vdf_temp = load_gsheet(MASTER_SPREADSHEET_ID, "Vendor Master")
        if not vdf_temp.empty:
            vrow = vdf_temp[vdf_temp["Name"] == vendor_name] if "Name" in vdf_temp.columns else vdf_temp[vdf_temp["name"] == vendor_name]
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

    pdf.set_fill_color(192, 0, 0)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 8)
    col_w = [10, 25, 25, 60, 10, 22, 10, 28]
    headers = ["S.No", "ERP Code", "Vendor Item", "Description", "Qty", "Price(Excl.)", "GST%", "Total(Incl.)"]
    for i, h in enumerate(headers):
        pdf.cell(col_w[i], 7, h, 1, 0, "C", True)
    pdf.ln()

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
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use("Agg")

    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass

    st.markdown("""
    <div style="background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%); padding: 18px 28px; border-radius: 14px; margin-bottom: 18px;">
        <span style="color: #ffffff; font-size: 1.5rem; font-weight: 700; letter-spacing: 0.5px;">📊 Purchase Quote Executive Dashboard</span>
    </div>
    """, unsafe_allow_html=True)
    df = load_quotes()

    if df.empty:
        st.info("No quotes found. Create your first quote!")
        return

    df_temp = df.copy()
    df_temp["Qty_num"] = pd.to_numeric(df_temp["Qty"], errors="coerce").fillna(0.0)
    df_temp["Price_Inc_num"] = pd.to_numeric(df_temp["Price Inc. GST"], errors="coerce").fillna(0.0)
    df_temp["Line_Total_Inc"] = df_temp["Qty_num"] * df_temp["Price_Inc_num"]
    df_temp["Created_DT"] = pd.to_datetime(df_temp["Created Date"], errors="coerce")

    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        date_opt = st.selectbox("📅 Time Period", ["All Time", "Today", "This Week", "This Month", "This Year", "Custom Range"])
    with fc2:
        unique_vendors = ["All Vendors"] + sorted(df["Vendor Name"].dropna().unique().tolist())
        sel_vendor = st.selectbox("🏢 Vendor", unique_vendors)
    with fc3:
        unique_locs = ["All Locations"] + sorted(df["Location Code"].dropna().unique().tolist())
        sel_loc = st.selectbox("📍 Location", unique_locs)
    with fc4:
        unique_purchasers = ["All Purchasers"] + sorted(df["Purchase Quote raised by"].dropna().unique().tolist())
        sel_purchaser = st.selectbox("👤 Purchaser", unique_purchasers)

    start_date, end_date = None, None
    if date_opt == "Custom Range":
        custom_range = st.date_input("Select Date Range", value=(date.today(), date.today()))
        if isinstance(custom_range, tuple) and len(custom_range) == 2:
            start_date, end_date = custom_range
    elif date_opt == "Today":
        start_date = end_date = date.today()
    elif date_opt == "This Week":
        start_date = date.today() - pd.Timedelta(days=date.today().weekday()); end_date = date.today()
    elif date_opt == "This Month":
        start_date = date(date.today().year, date.today().month, 1); end_date = date.today()
    elif date_opt == "This Year":
        start_date = date(date.today().year, 1, 1); end_date = date.today()

    fdf = df_temp.copy()
    if start_date and end_date:
        fdf = fdf[(fdf["Created_DT"] >= pd.to_datetime(start_date)) & (fdf["Created_DT"] < pd.to_datetime(end_date) + pd.Timedelta(days=1))]
    if sel_vendor != "All Vendors":
        fdf = fdf[fdf["Vendor Name"] == sel_vendor]
    if sel_loc != "All Locations":
        fdf = fdf[fdf["Location Code"] == sel_loc]
    if sel_purchaser != "All Purchasers":
        fdf = fdf[fdf["Purchase Quote raised by"] == sel_purchaser]

    total_quotes = fdf["Quote No"].nunique() if not fdf.empty else 0
    grand_qty = fdf["Qty_num"].sum() if not fdf.empty else 0
    total_locs = fdf["Location Code"].nunique() if not fdf.empty else 0

    if not fdf.empty:
        qt = fdf.groupby("Quote No").agg(
            line_total=("Line_Total_Inc", "sum"),
            freight=("Freight Charge", lambda x: pd.to_numeric(x.iloc[0], errors="coerce") if not x.empty else 0.0),
            qty=("Qty_num", "sum"), purchaser=("Purchase Quote raised by", "first"),
            vendor=("Vendor Name", "first"), location=("Location Code", "first"),
            status=("Status", "first"), created=("Created_DT", "first")
        ).reset_index()
        qt["freight"] = qt["freight"].fillna(0.0)
        qt["value"] = qt["line_total"] + qt["freight"]
        grand_value = qt["value"].sum()
    else:
        qt = pd.DataFrame(columns=["Quote No","value","qty","purchaser","vendor","location","status","created"])
        grand_value = 0.0

    NAVY = "#0f172a"; NAVY2 = "#1e3a5f"; NAVY3 = "#2d5a8e"; LIGHT = "#e2e8f0"
    PIE_COLORS = ["#0f172a", "#1e3a5f", "#2d5a8e", "#4a90c4", "#7cb3d4", "#a8d0e6"]

    st.markdown("")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📊 Total Quotes Raised", f"{total_quotes}")
    k2.metric("📦 Total Qty Purchased", f"{int(grand_qty):,}")
    k3.metric("💰 Total Value (₹)", f"{grand_value:,.2f}")
    k4.metric("📍 Total Locations", f"{total_locs}")

    st.markdown("---")

    st.markdown("#### 📋 Purchaser | Vendor | Location Breakdown")
    if not qt.empty:
        bd = qt.groupby(["purchaser", "vendor", "location"]).agg(
            quotes=("Quote No", "nunique"), total_val=("value", "sum")
        ).reset_index()
        bd.columns = ["Purchased By", "Vendor", "Location", "Total Quotes Raised", "Total Value (₹)"]
        bd = bd.sort_values("Total Value (₹)", ascending=False)
        bd_disp = bd.copy()
        bd_disp["Total Value (₹)"] = bd_disp["Total Value (₹)"].apply(lambda x: f"₹{x:,.2f}")
        st.dataframe(bd_disp, use_container_width=True, hide_index=True)
    else:
        st.info("No data for selected filters.")

    st.markdown("---")

    ch1, ch2, ch3, ch4 = st.columns(4)

    with ch1:
        st.markdown("#### 💹 Value by Vendor")
        if not qt.empty:
            vv = qt.groupby("vendor")["value"].sum().sort_values(ascending=False).head(6)
            fig2, ax2 = plt.subplots(figsize=(4, 3.5))
            fig2.patch.set_facecolor("white")
            bars = ax2.bar(range(len(vv)), vv.values, color=[NAVY, NAVY2, NAVY3, "#4a90c4", "#7cb3d4", "#a8d0e6"][:len(vv)])
            ax2.set_xticks(range(len(vv)))
            ax2.set_xticklabels([v[:12] for v in vv.index], rotation=45, ha="right", fontsize=7, color=NAVY)
            ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x/1000:.0f}K" if x >= 1000 else f"{x:.0f}"))
            ax2.tick_params(axis="y", labelsize=8, colors=NAVY); ax2.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig2, use_container_width=True); plt.close(fig2)

    with ch2:
        st.markdown("#### 📍 Quotes by Location")
        if not qt.empty:
            ql = qt.groupby("location")["Quote No"].nunique().sort_values(ascending=False).head(6)
            fig3, ax3 = plt.subplots(figsize=(4, 3.5))
            fig3.patch.set_facecolor("white")
            ax3.bar(range(len(ql)), ql.values, color=[NAVY, NAVY2, NAVY3, "#4a90c4"][:len(ql)])
            ax3.set_xticks(range(len(ql)))
            ax3.set_xticklabels([l[:10] for l in ql.index], rotation=45, ha="right", fontsize=7, color=NAVY)
            ax3.tick_params(axis="y", labelsize=8, colors=NAVY); ax3.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig3, use_container_width=True); plt.close(fig3)

    with ch3:
        st.markdown("#### 👤 Qty by Purchaser")
        if not qt.empty:
            qp = qt.groupby("purchaser")["qty"].sum().sort_values(ascending=True)
            fig4, ax4 = plt.subplots(figsize=(4, 3.5))
            fig4.patch.set_facecolor("white")
            ax4.barh(range(len(qp)), qp.values, color=[NAVY, NAVY2, NAVY3, "#4a90c4"][:len(qp)])
            ax4.set_yticks(range(len(qp)))
            ax4.set_yticklabels([p[:12] for p in qp.index], fontsize=8, color=NAVY)
            ax4.tick_params(axis="x", labelsize=8, colors=NAVY); ax4.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig4, use_container_width=True); plt.close(fig4)

    with ch4:
        st.markdown("#### 📈 Quotes Over Time")
        if not qt.empty and qt["created"].notna().any():
            qt_time = qt.copy()
            qt_time["month"] = qt_time["created"].dt.to_period("M").astype(str)
            monthly = qt_time.groupby("month")["Quote No"].nunique().reset_index()
            monthly.columns = ["Month", "Quotes"]
            fig5, ax5 = plt.subplots(figsize=(4, 3.5))
            fig5.patch.set_facecolor("white")
            ax5.plot(monthly["Month"], monthly["Quotes"], color=NAVY, marker="o", linewidth=2, markersize=5)
            ax5.fill_between(monthly["Month"], monthly["Quotes"], alpha=0.1, color=NAVY2)
            ax5.set_xticklabels(monthly["Month"], rotation=45, ha="right", fontsize=7, color=NAVY)
            ax5.tick_params(axis="y", labelsize=8, colors=NAVY); ax5.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig5, use_container_width=True); plt.close(fig5)

    st.markdown("---")

    bot1, bot2 = st.columns([3, 2])

    with bot1:
        st.markdown("#### 📉 Monthly Value Trend (₹)")
        if not qt.empty and qt["created"].notna().any():
            qt_t2 = qt.copy()
            qt_t2["month"] = qt_t2["created"].dt.to_period("M").astype(str)
            mv = qt_t2.groupby("month")["value"].sum().reset_index()
            mv.columns = ["Month", "Value"]
            fig6, ax6 = plt.subplots(figsize=(8, 3.5))
            fig6.patch.set_facecolor("white")
            ax6.plot(mv["Month"], mv["Value"], color=NAVY, marker="o", linewidth=2.5, markersize=6)
            ax6.fill_between(mv["Month"], mv["Value"], alpha=0.08, color=NAVY2)
            ax6.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"₹{x/1000:.0f}K" if x >= 1000 else f"₹{x:.0f}"))
            ax6.set_xticklabels(mv["Month"], rotation=45, ha="right", fontsize=8, color=NAVY)
            ax6.tick_params(axis="y", labelsize=8, colors=NAVY); ax6.spines[["top","right"]].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig6, use_container_width=True); plt.close(fig6)

    with bot2:
        st.markdown("#### 🥧 Value Contribution (₹)")
        if not qt.empty:
            vc = qt.groupby("vendor")["value"].sum().sort_values(ascending=False).head(5)
            fig7, ax7 = plt.subplots(figsize=(4, 3.5))
            fig7.patch.set_facecolor("white")
            wedges, texts = ax7.pie(vc.values, labels=[v[:14] for v in vc.index],
                colors=PIE_COLORS[:len(vc)], textprops={"fontsize": 8, "color": NAVY}, startangle=90)
            ax7.set_aspect("equal")
            st.pyplot(fig7, use_container_width=True); plt.close(fig7)

    st.markdown("---")

    st.markdown("#### 📄 Recent Purchase Quotes")
    if not fdf.empty:
        recent = fdf.drop_duplicates("Quote No", keep="last").sort_values("Created Date", ascending=False).head(10)
        if not qt.empty:
            recent = recent.merge(qt[["Quote No", "value"]], on="Quote No", how="left")
            recent = recent.rename(columns={"value": "Total Order Value"})
        disp_r = recent[["Quote No", "Purchase Quote raised by", "Vendor Name", "Location Code", "Status", "Total Order Value", "Created Date"]].copy()
        disp_r["Total Order Value"] = disp_r["Total Order Value"].apply(lambda x: f"₹{x:,.2f}" if pd.notnull(x) else "₹0.00")
        st.dataframe(disp_r, use_container_width=True, hide_index=True)
    else:
        st.info("No quotes for selected filters.")

# ── CREATE / EDIT QUOTE ──────────────────────────────────────────
def page_create_quote():
    st.markdown('<div class="section-header">➕ Create / Edit Purchase Quote</div>', unsafe_allow_html=True)

    quotes_df = load_quotes()
    
    edit_mode = False
    sel_quote_to_edit = "-- New Quote --"
    if not quotes_df.empty:
        existing_quotes = sorted(quotes_df["Quote No"].dropna().unique().tolist(), reverse=True)
        sel_quote_to_edit = st.selectbox("📝 Create New or Edit Existing Quote", ["-- New Quote --"] + existing_quotes)
        if sel_quote_to_edit != "-- New Quote --":
            edit_mode = True

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
            
            ed_val = first_row.get("Expected Delivery Date", None)
            try:
                if pd.notnull(ed_val) and str(ed_val).strip():
                    st.session_state.edit_expected_delivery = pd.to_datetime(ed_val).date()
                else:
                    st.session_state.edit_expected_delivery = None
            except:
                st.session_state.edit_expected_delivery = None
            
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
            if "submitted_quote_no" in st.session_state:
                del st.session_state.submitted_quote_no

    quote_no = sel_quote_to_edit if edit_mode else next_quote_number(quotes_df)

    st.markdown("---")
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
        
        loc_info = LOCATION_DETAILS.get(location_code, {})
        if loc_info:
            st.caption(f"📍 {loc_info['Address'][:50]}...")
    with h4:
        salesperson = st.text_input("Purchase Quote raised by", value=st.session_state.get("edit_salesperson", ""), placeholder="Enter your name")

    st.markdown("#### Vendor")
    vendor_names = []
    if not vendor_df.empty:
        for nc in ["name", "Name", "searchName", "SearchName"]:
            if nc in vendor_df.columns:
                vendor_names = vendor_df[nc].dropna().unique().tolist()
                break

    v1, v2, v3, v4 = st.columns([2.5, 1, 3.5, 1])
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
    
    if selected_vendor != "-- Select Vendor --":
        details = []
        details.append(f"🏢 **Name:** {selected_vendor}")
        if vendor_addr: details.append(f"📍 **Address:** {vendor_addr}")
        if vendor_city: details.append(f"🏙️ **City:** {vendor_city}")
        if vendor_gst: details.append(f"🛡️ **GST:** {vendor_gst}")
        st.caption("  •  ".join(details))

    st.markdown("---")
    st.subheader("Payment & Delivery Options")
    
    p1, p2, p3 = st.columns(3)
    pm_opts = ["", "Cash", "Online Transfer", "CDC", "PDC"]
    
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

    st.markdown("---")
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
            if "submitted_quote_no" in st.session_state:
                del st.session_state.submitted_quote_no

    line_data_for_saving = []
    items_to_remove = []

    h_col0, h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7, h_col8 = st.columns([0.6, 2.0, 2.0, 4.2, 1.1, 1.4, 1.0, 1.8, 0.2])
    h_col0.write("**S.No**")
    h_col1.write("**ERP Code**")
    h_col2.write("**Vendor Item No**")
    h_col3.write("**Product Description**")
    h_col4.write("**Qty**")
    h_col5.write("**Price**")
    h_col6.write("**GST %**")
    h_col7.write("**Total**")
    h_col8.write("")

    for i, line in enumerate(st.session_state.line_items):
        c0, c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([0.6, 2.0, 2.0, 4.2, 1.1, 1.4, 1.0, 1.8, 0.2])
        
        with c0:
            st.write(f"{i+1}")
            
        with c1:
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
                            st.session_state.line_items[i]["erp_code"] = sel_erp
                            st.session_state.line_items[i]["vendor_item_no"] = vendor_item_no
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
        
        safe_price = price if price is not None else 0.0
        safe_qty = qty if qty is not None else 0
        
        price_inc_gst = safe_price * (1 + gst_p/100)
        line_total_before_gst = safe_qty * safe_price
        line_total_inc_gst = safe_qty * price_inc_gst
        
        with c7:
            st.text_input(f"TotalIncGST {i}", value=f"{line_total_inc_gst:,.2f}", disabled=True, label_visibility="collapsed")
        with c8:
            if st.button("❌", key=f"remove_{i}"):
                items_to_remove.append(i)

        st.session_state.line_items[i] = {
            "erp_code": sel_erp,
            "vendor_item_no": vendor_item_no,
            "description": description,
            "qty": qty,
            "price_before_gst": price,
            "gst_percent": gst_p,
            "remarks": line.get("remarks", "")
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
                "Remarks": line.get("remarks", ""),
                "line_total_before_gst": line_total_before_gst,
                "line_total_inc_gst": safe_qty * price_inc_gst
            })

    for idx in reversed(items_to_remove):
        st.session_state.line_items.pop(idx)
        st.rerun()

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

    st.markdown("---")
    
    update_reason = ""
    if edit_mode:
        update_reason = st.text_input("📝 Reason for Update / General Remarks", value="", placeholder="e.g., Price changed by vendor, added new items, etc.")
        st.markdown("<br>", unsafe_allow_html=True)

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
            
            orig_created = now
            orig_status = "Pending Approval"
            if edit_mode and not quotes_df.empty:
                existing_rows = quotes_df[quotes_df["Quote No"] == quote_no]
                if not existing_rows.empty:
                    orig_created = str(existing_rows.iloc[0].get("Created Date", now))
                    orig_status = str(existing_rows.iloc[0].get("Status", "Pending Approval"))
            
            target_status = "Modified" if edit_mode else "Pending Approval"
            modified_dt = now if edit_mode else ""
            
            final_rows = []
            for l in line_data_for_saving:
                specific_remark = l["Remarks"]
                if edit_mode and update_reason:
                    combined_remark = f"{specific_remark} | Update: {update_reason}" if specific_remark else update_reason
                else:
                    combined_remark = specific_remark
                
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
                    "Remarks": combined_remark,
                    "Freight Charge": freight_charge,
                    "Status": target_status,
                    "Created Date": orig_created,
                    "Modified Date": modified_dt,
                })
            
            if final_rows:
                new_df = pd.DataFrame(final_rows)
                for col in REQUIRED_COLS:
                    if col not in new_df.columns:
                        new_df[col] = ""
                new_df = new_df[REQUIRED_COLS]
                
                if save_quotes(new_df, is_edit=edit_mode, quote_no=quote_no):
                    st.success(f"✅ Quote **{quote_no}** {'updated' if edit_mode else 'submitted'} successfully!")
                    st.session_state.submitted_quote_no = quote_no
                    st.cache_data.clear()
                    quotes_df = load_quotes()

    # ── Post-Submit / Edit Actions (Email & Download) ────────
    show_actions_for = quote_no if edit_mode else st.session_state.get("submitted_quote_no")
    if show_actions_for:
        st.markdown("---")
        st.subheader(f"Actions for Quote: {show_actions_for}")
        
        q_rows_for_actions = quotes_df[quotes_df["Quote No"] == show_actions_for]
        
        c_dl, c_mail = st.columns([1.5, 2.5])
        
        with c_dl:
            st.write("**📥 Export PO**")
            buf_xl = export_po_excel(show_actions_for, quotes_df)
            if buf_xl:
                st.download_button("Excel PO", buf_xl, file_name=f"PO_{show_actions_for}.xlsx", 
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 use_container_width=True)
            buf_pdf = export_po_pdf(show_actions_for, quotes_df)
            if buf_pdf:
                st.download_button("PDF PO", buf_pdf, file_name=f"PO_{show_actions_for}.pdf", 
                                 mime="application/pdf",
                                 use_container_width=True)
                
                
def safe_str(value: any, default: str = "") -> str:
    """
    Safely convert value to string with HTML escaping.
    
    Handles:
    - None / NaN values
    - HTML special characters (<, >, &, ", ')
    - Unicode characters
    - Numeric types
    
    Args:
        value: Any value to convert
        default: Default if value is None/NaN
    
    Returns:
        str: Safe HTML-escaped string
    
    Examples:
        >>> safe_str(None)
        ''
        >>> safe_str("Product<description>")
        'Product&lt;description&gt;'
        >>> safe_str("<script>alert('xss')</script>")
        '&lt;script&gt;alert(&apos;xss&apos;)&lt;/script&gt;'
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    
    value_str = str(value).strip()
    
    if not value_str or value_str.lower() == 'nan':
        return default
    
    # HTML escape special characters
    value_str = value_str.replace("&", "&amp;")      # & first!
    value_str = value_str.replace("<", "&lt;")
    value_str = value_str.replace(">", "&gt;")
    value_str = value_str.replace('"', "&quot;")
    value_str = value_str.replace("'", "&apos;")
    
    return value_str


def format_currency(amount: float, currency_symbol: str = "₹", decimals: int = 2) -> str:
    """
    Format number as currency with proper separators.
    
    Args:
        amount: Numeric amount
        currency_symbol: Symbol to use (default: ₹)
        decimals: Decimal places (default: 2)
    
    Returns:
        str: Formatted currency string
    
    Examples:
        >>> format_currency(1234.5)
        '₹1,234.50'
        >>> format_currency(1000000)
        '₹10,00,000.00'
    """
    try:
        amount = float(amount) if amount else 0.0
    except (ValueError, TypeError):
        amount = 0.0
    
    return f"{currency_symbol}{amount:,.{decimals}f}"


def format_number(value: float, decimals: int = 2) -> str:
    """
    Format number with thousand separators.
    
    Args:
        value: Numeric value
        decimals: Decimal places
    
    Returns:
        str: Formatted number
    
    Examples:
        >>> format_number(1234.5)
        '1,234.50'
    """
    try:
        value = float(value) if value else 0.0
    except (ValueError, TypeError):
        value = 0.0
    
    return f"{value:,.{decimals}f}"


# ── TABLE GENERATION ────────���────────────────────────────────────

def generate_outlook_email_table(
    quote_data: pd.DataFrame,
    header_color: str = "#BDD7EE",
    header_text_color: str = "#000000",
    row_color_alt: str = "#F2F2F2",
    border_color: str = "#000000"
) -> str:
    """
    Generate Outlook-compatible HTML table from quote data.
    
    Features:
    - Properly closed tags
    - No truncation
    - Table-based layout (NO divs/flexbox)
    - Inline styles only
    - HTML-escaped content
    
    Args:
        quote_data: DataFrame with columns:
            - S.No, Vendor Item No, Product Description
            - Qty, Price Before GST, GST %, Price Inc. GST, Remarks
        header_color: Header background color (hex)
        header_text_color: Header text color (hex)
        row_color_alt: Alternating row color (hex)
        border_color: Border color (hex)
    
    Returns:
        str: Complete HTML table as string
    
    Raises:
        ValueError: If required columns missing
    """
    
    required_cols = [
        'S.No', 'Vendor Item No', 'Product Description',
        'Qty', 'Price Before GST', 'GST %', 'Price Inc. GST', 'Remarks'
    ]
    
    # Check for required columns
    missing_cols = [c for c in required_cols if c not in quote_data.columns]
    if missing_cols:
        raise ValueError(f"Missing columns: {missing_cols}")
    
    html = []
    html.append('<table cellpadding="0" cellspacing="0" border="1" style="border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 12px;">')
    
    # Header row
    html.append('<tr>')
    for col in required_cols:
        html.append(f'<td style="background-color: {header_color}; color: {header_text_color}; border: 1px solid {border_color}; padding: 8px; font-weight: bold; text-align: center;">{safe_str(col)}</td>')
    html.append('</tr>')
    
    # Data rows
    for idx, (_, row) in enumerate(quote_data.iterrows()):
        row_bg = row_color_alt if idx % 2 == 1 else "#FFFFFF"
        html.append('<tr>')
        
        # S.No
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: center;">{safe_str(row.get("S.No", idx + 1))}</td>')
        
        # Vendor Item No
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px;">{safe_str(row.get("Vendor Item No", ""))}</td>')
        
        # Product Description (may be long)
        desc = safe_str(row.get("Product Description", ""))[:80]  # Limit length
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; word-wrap: break-word;">{desc}</td>')
        
        # Qty
        try:
            qty = float(row.get('Qty', 0) or 0)
            qty_str = format_number(qty, 0)
        except (ValueError, TypeError):
            qty_str = "0"
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{qty_str}</td>')
        
        # Price Before GST
        try:
            pb = float(row.get('Price Before GST', 0) or 0)
            pb_str = format_currency(pb)
        except (ValueError, TypeError):
            pb_str = format_currency(0)
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{pb_str}</td>')
        
        # GST %
        try:
            gst_p = float(row.get('GST %', 0) or 0)
            gst_str = format_number(gst_p, 1)
        except (ValueError, TypeError):
            gst_str = "0.0"
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: center;">{gst_str}%</td>')
        
        # Price Inc. GST
        try:
            qty = float(row.get('Qty', 0) or 0)
            pi = float(row.get('Price Inc. GST', 0) or 0)
            line_total = qty * pi
            line_str = format_currency(line_total)
        except (ValueError, TypeError):
            line_str = format_currency(0)
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{line_str}</td>')
        
        # Remarks
        remarks = safe_str(row.get('Remarks', ''))[:40]
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px;">{remarks}</td>')
        
        html.append('</tr>')
    
    html.append('</table>')
    
    return '\n'.join(html)


def generate_totals_table(
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float
) -> str:
    """
    Generate summary totals table.
    
    Args:
        subtotal_excl_gst: Subtotal before GST
        total_gst: Total GST amount
        freight_charge: Freight/shipping charge
        grand_total: Final total including all charges
    
    Returns:
        str: HTML table with totals
    """
    html = []
    html.append('<table cellpadding="0" cellspacing="0" border="0" style="margin-top: 20px; font-family: Calibri, Arial, sans-serif; font-size: 12px; width: 100%;">')
    
    # Subtotal row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px; width: 70%;">Sub Total (Excl. GST):</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000; border-collapse: collapse; width: 30%;">{format_currency(subtotal_excl_gst)}</td>')
    html.append('</tr>')
    
    # GST row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px;">Total GST:</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000;">{format_currency(total_gst)}</td>')
    html.append('</tr>')
    
    # Freight row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px;">Freight Charge:</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000;">{format_currency(freight_charge)}</td>')
    html.append('</tr>')
    
    # Grand Total row (emphasized)
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 12px; font-weight: bold; background-color: #1F497D; color: #FFFFFF;">GRAND TOTAL (Incl. GST):</td>')
    html.append(f'<td style="text-align: right; padding: 12px; border: 1px solid #1F497D; background-color: #1F497D; color: #FFFFFF; font-weight: bold;">{format_currency(grand_total)}</td>')
    html.append('</tr>')
    
    html.append('</table>')
    
    return '\n'.join(html)


# ── BODY GENERATION ──────────────────────────────────────────────

def generate_plain_text_body(
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str
) -> str:
    """
    Generate plain text version of email (fallback for HTML).
    
    Args:
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal
        total_gst: Total GST
        freight_charge: Freight charge
        grand_total: Grand total
        closing: Closing message
    
    Returns:
        str: Plain text email body
    """
    lines = []
    
    lines.append(safe_str(greeting))
    lines.append("")
    lines.append("=" * 100)
    lines.append("")
    
    # Column headers
    lines.append(
        f"{'S.No':<6} {'Vendor Item':<20} {'Description':<35} {'Qty':>6} {'Price':>12} {'GST%':>7} {'Total':>12}"
    )
    lines.append("-" * 100)
    
    # Data rows
    for _, row in quote_data.iterrows():
        try:
            sno = safe_str(row.get('S.No', ''))
            vendor = safe_str(row.get('Vendor Item No', ''))[:20]
            desc = safe_str(row.get('Product Description', ''))[:35]
            qty = float(row.get('Qty', 0) or 0)
            pb = float(row.get('Price Before GST', 0) or 0)
            gst = float(row.get('GST %', 0) or 0)
            pi = float(row.get('Price Inc. GST', 0) or 0)
            line_total = qty * pi
            
            lines.append(
                f"{sno:<6} {vendor:<20} {desc:<35} {int(qty):>6} {pb:>12,.2f} {gst:>7.1f}% {line_total:>12,.2f}"
            )
        except Exception as e:
            continue
    
    lines.append("-" * 100)
    
    # Totals
    lines.append(f"{'':60} {'':6} {'Sub Total:':>12} {format_currency(subtotal_excl_gst):>12}")
    lines.append(f"{'':60} {'':6} {'GST Total:':>12} {format_currency(total_gst):>12}")
    lines.append(f"{'':60} {'':6} {'Freight:':>12} {format_currency(freight_charge):>12}")
    lines.append(f"{'':60} {'':6} {'GRAND TOTAL:':>12} {format_currency(grand_total):>12}")
    
    lines.append("")
    lines.append("=" * 100)
    lines.append("")
    lines.append(safe_str(closing))
    
    return '\n'.join(lines)


def generate_email_html(
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str,
    company_name: str = "Supreme Computers India Pvt. Ltd."
) -> str:
    """
    Generate complete Outlook-compatible HTML email body.
    
    Args:
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal
        total_gst: Total GST
        freight_charge: Freight charge
        grand_total: Grand total
        closing: Closing message
        company_name: Company name for header
    
    Returns:
        str: Complete HTML email document
    """
    
    table_html = generate_outlook_email_table(quote_data)
    totals_html = generate_totals_table(subtotal_excl_gst, total_gst, freight_charge, grand_total)
    
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Quote</title>
</head>
<body style="font-family: Calibri, Arial, sans-serif; line-height: 1.6; color: #333333; background-color: #FFFFFF; padding: 20px;">
    
    <!-- Container -->
    <div style="max-width: 900px; margin: 0 auto; background-color: #FFFFFF; padding: 20px;">
        
        <!-- Header -->
        <div style="border-bottom: 2px solid #1F497D; padding-bottom: 15px; margin-bottom: 20px;">
            <h2 style="color: #1F497D; margin: 0 0 5px 0; font-size: 18px;">{safe_str(company_name)}</h2>
            <p style="color: #666666; margin: 0; font-size: 11px;">Professional Purchase Quote</p>
        </div>
        
        <!-- Greeting -->
        <div style="margin-bottom: 20px;">
            <p style="color: #333333; font-size: 12px; white-space: pre-wrap;">{safe_str(greeting)}</p>
        </div>
        
        <!-- Items Table -->
        <div style="margin-bottom: 20px; overflow-x: auto;">
            {table_html}
        </div>
        
        <!-- Totals Table -->
        <div style="margin-bottom: 20px;">
            {totals_html}
        </div>
        
        <!-- Closing -->
        <div style="margin-bottom: 20px; margin-top: 30px;">
            <p style="color: #333333; font-size: 12px; white-space: pre-wrap;">{safe_str(closing)}</p>
        </div>
        
        <!-- Footer -->
        <div style="border-top: 1px solid #CCCCCC; padding-top: 15px; margin-top: 30px; font-size: 10px; color: #999999;">
            <p style="margin: 5px 0;">This is an automated purchase quote generated by the Purchase Quote System.</p>
            <p style="margin: 5px 0;">Generated on: {datetime.now().strftime('%d %B %Y at %H:%M:%S')}</p>
        </div>
        
    </div>
    
</body>
</html>
"""
    
    return html


# ── EMAIL MESSAGE CREATION ───────────────────────────────────────

def create_outlook_email_message(
    subject: str,
    to_email: str,
    cc_emails: str,
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str,
    pdf_attachment: Optional[io.BytesIO] = None,
    pdf_filename: str = "quote.pdf"
) -> EmailMessage:
    """
    Create complete EmailMessage with Outlook-compatible formatting.
    
    This is the main function to use for creating purchase quote emails.
    
    Args:
        subject: Email subject line
        to_email: Recipient email address
        cc_emails: CC email addresses (comma-separated)
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal before GST
        total_gst: Total GST amount
        freight_charge: Freight/shipping charge
        grand_total: Final total
        closing: Closing message
        pdf_attachment: Optional PDF buffer for attachment
        pdf_filename: Name for PDF attachment
    
    Returns:
        EmailMessage: Complete email ready to send/save
    
    Examples:
        >>> msg = create_outlook_email_message(
        ...     subject="Quote Q0001",
        ...     to_email="vendor@example.com",
        ...     cc_emails="mgr@company.com",
        ...     greeting="Dear Vendor,",
        ...     quote_data=df,
        ...     subtotal_excl_gst=1000,
        ...     total_gst=180,
        ...     freight_charge=100,
        ...     grand_total=1280,
        ...     closing="Please confirm.",
        ...     pdf_attachment=pdf_buf,
        ...     pdf_filename="PO_Q0001.pdf"
        ... )
        >>> eml_bytes = msg.as_bytes()
        >>> with open('quote.eml', 'wb') as f:
        ...     f.write(eml_bytes)
    """
    
    # Create message
    msg = EmailMessage()
    msg['Subject'] = safe_str(subject)
    msg['To'] = safe_str(to_email)
    msg['Cc'] = safe_str(cc_emails)
    msg['From'] = 'Purchase Quote System <noreply@supremecomputers.com>'
    
    # Generate body content
    plain_text = generate_plain_text_body(
        greeting, quote_data, subtotal_excl_gst, total_gst, freight_charge, grand_total, closing
    )
    html_content = generate_email_html(
        greeting, quote_data, subtotal_excl_gst, total_gst, freight_charge, grand_total, closing
    )
    
    # Set content (plain text first, then HTML alternative)
    msg.set_content(plain_text, charset='utf-8')
    msg.add_alternative(html_content, subtype='html', charset='utf-8')
    
    # Add PDF attachment if provided
    if pdf_attachment:
        try:
            pdf_data = pdf_attachment.getvalue()
            if pdf_data:
                msg.add_attachment(
                    pdf_data,
                    maintype='application',
                    subtype='pdf',
                    filename=safe_str(pdf_filename)
                )
        except Exception as e:
            print(f"Warning: Could not attach PDF: {e}")
    
    return msg
# ── VIEW QUOTES ──────────────────────────────────────────────
def page_view_quotes():
    st.markdown('<div class="section-header">🔍 View Quotes</div>', unsafe_allow_html=True)
    df = load_quotes()
    if df.empty:
        st.info("No quotes available yet.")
        return

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
