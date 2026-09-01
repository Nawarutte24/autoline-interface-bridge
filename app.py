import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import os
import sys
import datetime
from io import BytesIO

# =============================================================================
# 1. DEFAULT SETTINGS & FALLBACK CONFIGURATION
# =============================================================================
DEFAULT_SETTINGS = {
    "journal_desc": "After Sale",
    "doc_code": "ARI",
    "currency": "THB",
    "tax_group": "U",
    "terms": 30,
    "src_branch": "0001",
    "subaccount_default": "A0011",
    "ar_gl_code": "11311001",
    "ar_department": "0000",
    "parts": {
        "doc_seq": "PINVOICE",
        "gl_code": "41211004",
        "department": "4002",
        "tax_code": "S",
        "aftstype": "R",
        "partfran": "J",
        "partprod": "A"
    },
    "labor": {
        "doc_seq": "SINVOICE",
        "gl_code": "41111001",
        "department": "4001",
        "tax_code": "S",
        "aftstype": "R",
        "servprod": "A",
        "servicefranc": "JEEP"
    },
    "sublet": {
        "doc_seq": "SINVOICE",
        "gl_code": "41311001",
        "department": "4003",
        "tax_code": "S",
        "aftstype": "R",
        "servprod": "A",
        "servicefranc": "JEEP"
    }
}

# =============================================================================
# 2. TEMPLATE BUILDERS (AUTOLINE & MASTER DATA)
# =============================================================================
def get_base_autoline_workbook(template_path=None):
    """
    Loads output_template.xlsx if exists, or builds exact Autoline schema from code.
    """
    if template_path and os.path.exists(template_path):
        try:
            return openpyxl.load_workbook(template_path)
        except Exception:
            pass
            
    # Dynamic generation of exact Output Template (NEW)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Row 1
    ws.cell(1, 1).value = "This template is for AR/AP documents using debit and credit value columns"
    ws.cell(1, 11).value = "Rows 1 to 6 must be copied to any new sheets. Data for import must start on line 7 of the sheet."
    
    # Row 3 (Header Schema)
    h_headers = [
        "JNLDESC", "JNLCODE", "JNLPAGE", "JNLBATCH", "JNLITEM", "JNLDOCCODE",
        "CURRENCY", "DOCSEQ", "TOTALTAX", "TOTALVALUE", "DOCDATE", "DOCNUM",
        "ITEMREFERENCE", "MISCREFERENCE", "DELIVEREDDATE", "SUBACCOUNT", "CONTACT",
        "SUBACNAME", "SUBACPOSTCODE", "SUBACTAXNO", "TAXGROUP", "SETTTERMS",
        "TERMS", "SRCBRANCH", "CRCODE"
    ]
    for idx, h in enumerate(h_headers, 1):
        ws.cell(3, idx).value = h
        
    # Row 5 (Item Schema)
    i_headers = [
        "", "JNLCODE", "JNLPAGE", "JNLBATCH", "JNLITEM", "JNLDOCCODE", "CURRENCY",
        "JNLLINE", "BRANCH", "DEPARTMENT", "GLCODE", "SUFFIX", "DEBIT", "CREDIT",
        "NARRATIVE", "QUANTITY", "TAXCODE", "DISCOUNTLINE", "ADVICENOTE", "ORDERNUMBER",
        "MANUFACT", "MODEL", "SALETYPE", "AFTSTYPE", "PRODGROUP", "PARTFRAN",
        "SERVPROD", "PARTPROD", "SERVICEFRANC", "LA10"
    ]
    for idx, h in enumerate(i_headers, 1):
        if h:
            ws.cell(5, idx).value = h
            
    return wb

def get_master_template_bytes():
    """Generates Master Data Template Excel in memory."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
    
    # Sheet 1: Customer_Subaccount
    ws1 = wb.create_sheet(title="Customer_Subaccount")
    ws1.append(["customer_code", "customer_name", "tax_ID", "autoline_subaccount", "terms", "tax_group", "remark"])
    ws1.append(["1624", "บริษัท มาสเตอร์ ไดรฟเวอร์ แอนด์ เซอร์วิสเซส ( ประเทศไทย ) จำกัด", "0105550078073", "A0011", 30, "U", "ลูกหนี้บริษัท"])
    ws1.append(["2357", "คุณ ภูมิธเนษฐ์ นพสุวรรณวงศ์", "", "A0011", 30, "U", "ลูกค้าทั่วไป"])
    ws1.append(["DEFAULT", "ค่าเริ่มต้น", "", "A0011", 30, "U", "รหัสกลาง"])
    
    # Sheet 2: Branch_Mapping
    ws2 = wb.create_sheet(title="Branch_Mapping")
    ws2.append(["dealer_prefix", "branch_name", "dealer_code", "autoline_branch_code", "remark"])
    ws2.append(["SKV", "Sukhumvit", 5, "0001", "สาขาสุขุมวิท"])
    ws2.append(["RM4", "Rama 4", 6, "0002", "สาขาพระราม 4"])
    ws2.append(["DEFAULT", "", "", "0001", "สาขาเริ่มต้น"])
    
    # Sheet 3: GL_Mapping
    ws3 = wb.create_sheet(title="GL_Mapping")
    ws3.append(["category", "category_name", "gl_code", "department", "tax_code", "aftstype", "partfran", "servprod", "partprod", "servicefranc", "doc_seq"])
    ws3.append(["P", "Parts (อะไหล่)", "41211004", "4002", "S", "R", "J", "", "A", "", "PINVOICE"])
    ws3.append(["L", "Labor (ค่าแรง)", "41111001", "4001", "S", "R", "", "A", "", "JEEP", "SINVOICE"])
    ws3.append(["S", "Sublet (บริการภายนอก)", "41311001", "4003", "S", "R", "", "A", "", "JEEP", "SINVOICE"])
    ws3.append(["AR", "ลูกหนี้การค้า (AR Control)", "11311001", "0000", "", "", "", "", "", "", ""])
    
    for ws in wb.worksheets:
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=1, column=col_idx)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center")
            
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# =============================================================================
# 3. PARSER & TRANSFORMER LOGIC
# =============================================================================
def clean_advisor_first_name(advisor_name):
    if not advisor_name or pd.isna(advisor_name):
        return ""
    s = str(advisor_name).strip()
    parts = s.split()
    return parts[0] if parts else s

def load_and_validate_inputs(header_file, detail_file):
    df_header = pd.read_excel(header_file)
    df_detail = pd.read_excel(detail_file)
        
    df_header.columns = [str(c).strip() for c in df_header.columns]
    df_detail.columns = [str(c).strip() for c in df_detail.columns]
    
    if "invoice_number" not in df_header.columns:
        raise ValueError("ไฟล์ Header ไม่มีคอลัมน์ 'invoice_number'")
    if "invoice_number" not in df_detail.columns:
        raise ValueError("ไฟล์ Detail ไม่มีคอลัมน์ 'invoice_number'")
        
    df_header["invoice_number"] = df_header["invoice_number"].astype(str).str.strip()
    df_detail["invoice_number"] = df_detail["invoice_number"].astype(str).str.strip()
    
    df_header = df_header[df_header["invoice_number"] != ""].copy()
    df_detail = df_detail[df_detail["invoice_number"] != ""].copy()
    
    df_detail["category"] = df_detail["category"].fillna("").astype(str).str.strip().str.upper()
    
    for col in ["sales", "sales_tax", "nett_price", "discount_amount"]:
        if col in df_detail.columns:
            df_detail[col] = pd.to_numeric(df_detail[col], errors="coerce").fillna(0.0)
            
    for col in ["sales", "sales_tax", "total", "labor_amount", "parts_amount", "other_amount", "nett_price"]:
        if col in df_header.columns:
            df_header[col] = pd.to_numeric(df_header[col], errors="coerce").fillna(0.0)
            
    if "invoice_create_date" in df_header.columns:
        df_header["parsed_date"] = pd.to_datetime(df_header["invoice_create_date"], errors="coerce")
    else:
        df_header["parsed_date"] = pd.Timestamp.now()
        
    return df_header, df_detail

def parse_master_data_file(master_file):
    xl = pd.ExcelFile(master_file)
    master_dict = {"customers": {}, "branches": {}, "gl_rules": {}}
    
    for s in xl.sheet_names:
        s_low = s.lower()
        if "customer" in s_low or "subaccount" in s_low:
            df_c = pd.read_excel(master_file, sheet_name=s)
            for _, r in df_c.iterrows():
                code = str(r.get("customer_code", "")).strip()
                tax_id = str(r.get("tax_ID", "")).strip()
                sub = str(r.get("autoline_subaccount", "")).strip()
                terms = r.get("terms", 30)
                if code and sub:
                    master_dict["customers"][code] = {"subaccount": sub, "terms": terms}
                if tax_id and sub:
                    master_dict["customers"][tax_id] = {"subaccount": sub, "terms": terms}
                    
        elif "branch" in s_low:
            df_b = pd.read_excel(master_file, sheet_name=s)
            for _, r in df_b.iterrows():
                prefix = str(r.get("dealer_prefix", "")).strip().upper()
                b_name = str(r.get("branch_name", "")).strip().lower()
                code = str(r.get("autoline_branch_code", "")).strip()
                if prefix and code:
                    master_dict["branches"][prefix] = code
                if b_name and code:
                    master_dict["branches"][b_name] = code
                    
        elif "gl" in s_low:
            df_gl = pd.read_excel(master_file, sheet_name=s)
            for _, r in df_gl.iterrows():
                cat = str(r.get("category", "")).strip().upper()
                if cat:
                    master_dict["gl_rules"][cat] = {
                        "gl_code": str(r.get("gl_code", "")).strip(),
                        "department": str(r.get("department", "")).strip(),
                        "tax_code": str(r.get("tax_code", "S")).strip(),
                        "aftstype": str(r.get("aftstype", "R")).strip(),
                        "partfran": str(r.get("partfran", "")).strip() if pd.notna(r.get("partfran")) else None,
                        "partprod": str(r.get("partprod", "")).strip() if pd.notna(r.get("partprod")) else None,
                        "servprod": str(r.get("servprod", "")).strip() if pd.notna(r.get("servprod")) else None,
                        "servicefranc": str(r.get("servicefranc", "")).strip() if pd.notna(r.get("servicefranc")) else None,
                        "doc_seq": str(r.get("doc_seq", "")).strip() if pd.notna(r.get("doc_seq")) else None,
                    }
    return master_dict

def transform_to_autoline_data(df_header, df_detail, config, master_dict=None):
    rows_to_write = []
    preview_records = []
    
    batch_idx = 1
    total_debit_sum = 0.0
    total_credit_sum = 0.0
    total_tax_sum = 0.0
    
    if master_dict is None:
        master_dict = {}
    cust_master = master_dict.get("customers", {})
    branch_master = master_dict.get("branches", {})
    gl_master = master_dict.get("gl_rules", {})
    
    detail_by_inv = {inv: grp for inv, grp in df_detail.groupby("invoice_number")}
        
    for _, h_row in df_header.iterrows():
        inv_no = str(h_row["invoice_number"]).strip()
        if not inv_no:
            continue
            
        doc_date = h_row.get("parsed_date")
        if pd.isna(doc_date) or not isinstance(doc_date, (datetime.datetime, pd.Timestamp)):
            doc_date = datetime.datetime.now()
        else:
            doc_date = doc_date.to_pydatetime() if isinstance(doc_date, pd.Timestamp) else doc_date
            
        advisor_full = h_row.get("service_advisor_name", "")
        advisor_first = clean_advisor_first_name(advisor_full)
        narrative = f"{advisor_first}_{inv_no}" if advisor_first else inv_no
        
        h_nett = float(h_row.get("nett_price", 0.0))
        h_tax = float(h_row.get("sales_tax", 0.0))
        h_total = float(h_row.get("total", 0.0))
        
        cust_code = str(h_row.get("customer_code", "")).strip()
        tax_id = str(h_row.get("tax_ID", "")).strip()
        subaccount = config.get("subaccount_default", "A0011")
        terms_val = int(config.get("terms", 30))
        
        if cust_code in cust_master:
            subaccount = cust_master[cust_code].get("subaccount", subaccount)
            terms_val = int(cust_master[cust_code].get("terms", terms_val))
        elif tax_id in cust_master:
            subaccount = cust_master[tax_id].get("subaccount", subaccount)
            terms_val = int(cust_master[tax_id].get("terms", terms_val))
            
        dealer_pfx = str(h_row.get("dealer_prefix", "")).strip().upper()
        branch_name_val = str(h_row.get("branch_name", "")).strip().lower()
        src_branch = config.get("src_branch", "0001")
        if dealer_pfx in branch_master:
            src_branch = branch_master[dealer_pfx]
        elif branch_name_val in branch_master:
            src_branch = branch_master[branch_name_val]
        
        d_group = detail_by_inv.get(inv_no, pd.DataFrame())
        
        cat_sales = {}
        has_parts = False
        has_service = False
        
        if not d_group.empty:
            cat_sums = d_group.groupby("category")["sales"].sum().to_dict()
            cat_tax_sums = d_group.groupby("category")["sales_tax"].sum().to_dict()
            
            valid_cats = {k: v for k, v in cat_sums.items() if v > 0 and k in ["P", "L", "S"]}
            valid_tax_cats = {k: v for k, v in cat_tax_sums.items() if v > 0 and k in ["P", "L", "S"]}
            
            if len(valid_cats) == 1:
                cat_name = list(valid_cats.keys())[0]
                cat_sales[cat_name] = round(h_nett, 2)
            elif len(valid_cats) > 1:
                total_tax_valid = sum(valid_tax_cats.values())
                if total_tax_valid > 0:
                    allocated = 0.0
                    cats_list = list(valid_tax_cats.keys())
                    for c in cats_list[:-1]:
                        c_amt = round(h_nett * (valid_tax_cats[c] / total_tax_valid), 2)
                        cat_sales[c] = c_amt
                        allocated += c_amt
                    cat_sales[cats_list[-1]] = round(h_nett - allocated, 2)
                else:
                    total_s_valid = sum(valid_cats.values())
                    allocated = 0.0
                    cats_list = list(valid_cats.keys())
                    for c in cats_list[:-1]:
                        c_amt = round(h_nett * (valid_cats[c] / total_s_valid), 2)
                        cat_sales[c] = c_amt
                        allocated += c_amt
                    cat_sales[cats_list[-1]] = round(h_nett - allocated, 2)
            else:
                cat_sales["P"] = round(h_nett, 2)
        else:
            cat_sales["P"] = round(h_nett, 2)
            
        if "P" in cat_sales and cat_sales["P"] > 0:
            has_parts = True
        if any(cat_sales.get(c, 0) > 0 for c in ["L", "S"]):
            has_service = True
                
        if has_parts and not has_service:
            doc_seq = gl_master.get("P", {}).get("doc_seq") or config["parts"].get("doc_seq", "PINVOICE")
        else:
            doc_seq = gl_master.get("L", {}).get("doc_seq") or config["labor"].get("doc_seq", "SINVOICE")
            
        # 1. Header Row
        header_record = {
            "type": "HEADER",
            "A": config.get("journal_desc", "After Sale"),
            "B": inv_no,
            "C": 1,
            "D": batch_idx,
            "E": None,
            "F": config.get("doc_code", "ARI"),
            "G": config.get("currency", "THB"),
            "H": doc_seq,
            "I": h_tax,
            "J": h_total,
            "K": doc_date,
            "L": None,
            "M": inv_no,
            "N": narrative,
            "O": doc_date,
            "P": subaccount,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "U": config.get("tax_group", "U"),
            "V": None,
            "W": terms_val,
            "X": src_branch,
            "Y": None
        }
        rows_to_write.append(header_record)
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": "HEADER",
            "Batch": batch_idx,
            "Line": "",
            "Doc Code": config.get("doc_code", "ARI"),
            "Doc Seq": doc_seq,
            "Date": doc_date.strftime("%Y-%m-%d"),
            "GL Code": "",
            "Department": "",
            "Debit": "",
            "Credit": "",
            "Tax": h_tax,
            "Total Value": h_total,
            "Narrative": narrative,
            "Subaccount": subaccount
        })
        
        # 2. Blank Row (Row 8 style)
        rows_to_write.append({"type": "BLANK"})
        
        # 3. Line 1: DEBIT Line (AR Control Account) -> UP FIRST
        ar_gl = gl_master.get("AR", {}).get("gl_code") or str(config.get("ar_gl_code", "11311001"))
        ar_dept = gl_master.get("AR", {}).get("department") or str(config.get("ar_department", "0000"))
        
        debit_record = {
            "type": "ITEM",
            "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
            "G": config.get("currency", "THB"), "H": 1,
            "I": src_branch, "J": ar_dept, "K": int(ar_gl), "L": None,
            "M": h_total, "N": None, "O": narrative, "P": None, "Q": None,
            "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
            "X": None, "Y": None, "Z": None, "AA": None, "AB": None, "AC": None, "AD": None
        }
        rows_to_write.append(debit_record)
        total_debit_sum += h_total
        
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": "DEBIT (AR)",
            "Batch": batch_idx,
            "Line": 1,
            "Doc Code": "",
            "Doc Seq": "",
            "Date": doc_date.strftime("%Y-%m-%d"),
            "GL Code": ar_gl,
            "Department": ar_dept,
            "Debit": h_total,
            "Credit": "",
            "Tax": "",
            "Total Value": "",
            "Narrative": narrative,
            "Subaccount": subaccount
        })
        
        # 4. Line 2+: CREDIT Lines
        line_num = 2
        for cat, amount in cat_sales.items():
            if amount <= 0:
                continue
                
            m_rule = gl_master.get(cat, {})
            
            if cat == "P":
                cfg = config["parts"]
                gl_code_val = int(m_rule.get("gl_code") or cfg.get("gl_code", "41211004"))
                dept_val = str(m_rule.get("department") or cfg.get("department", "4002"))
                tax_code_val = m_rule.get("tax_code") or cfg.get("tax_code", "S")
                aftstype_val = m_rule.get("aftstype") or cfg.get("aftstype", "R")
                partfran = m_rule.get("partfran") or cfg.get("partfran", "J")
                partprod = m_rule.get("partprod") or cfg.get("partprod", "A")
                servprod = None
                servicefranc = None
            elif cat == "L":
                cfg = config["labor"]
                gl_code_val = int(m_rule.get("gl_code") or cfg.get("gl_code", "41111001"))
                dept_val = str(m_rule.get("department") or cfg.get("department", "4001"))
                tax_code_val = m_rule.get("tax_code") or cfg.get("tax_code", "S")
                aftstype_val = m_rule.get("aftstype") or cfg.get("aftstype", "R")
                partfran = None
                partprod = None
                servprod = m_rule.get("servprod") or cfg.get("servprod", "A")
                servicefranc = m_rule.get("servicefranc") or cfg.get("servicefranc", "JEEP")
            else: # S
                cfg = config["sublet"]
                gl_code_val = int(m_rule.get("gl_code") or cfg.get("gl_code", "41311001"))
                dept_val = str(m_rule.get("department") or cfg.get("department", "4003"))
                tax_code_val = m_rule.get("tax_code") or cfg.get("tax_code", "S")
                aftstype_val = m_rule.get("aftstype") or cfg.get("aftstype", "R")
                partfran = None
                partprod = None
                servprod = m_rule.get("servprod") or cfg.get("servprod", "A")
                servicefranc = m_rule.get("servicefranc") or cfg.get("servicefranc", "JEEP")
            
            credit_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": config.get("currency", "THB"), "H": line_num,
                "I": src_branch, "J": dept_val, "K": gl_code_val, "L": None,
                "M": None, "N": amount, "O": narrative, "P": None, "Q": tax_code_val,
                "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
                "X": aftstype_val, "Y": None, "Z": partfran, "AA": servprod, "AB": partprod, "AC": servicefranc, "AD": None
            }
            rows_to_write.append(credit_record)
            total_credit_sum += amount
            
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": f"CREDIT ({cat})",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d"),
                "GL Code": str(gl_code_val),
                "Department": dept_val,
                "Debit": "",
                "Credit": amount,
                "Tax": tax_code_val,
                "Total Value": "",
                "Narrative": narrative,
                "Subaccount": ""
            })
            line_num += 1
            
        total_tax_sum += h_tax
        batch_idx += 1
        
    summary_stats = {
        "total_invoices": len(df_header),
        "total_debit": total_debit_sum,
        "total_credit": total_credit_sum,
        "total_tax": total_tax_sum,
        "is_balanced": abs(total_debit_sum - (total_credit_sum + total_tax_sum)) < 0.05,
        "difference": round(total_debit_sum - (total_credit_sum + total_tax_sum), 2)
    }
    
    preview_df = pd.DataFrame(preview_records)
    return rows_to_write, summary_stats, preview_df

def generate_output_excel(template_path, rows_to_write):
    wb = get_base_autoline_workbook(template_path)
    ws = wb.active
    
    if ws.max_row >= 7:
        ws.delete_rows(7, ws.max_row - 6 + 10)
        
    current_row = 7
    for rec in rows_to_write:
        rec_type = rec.get("type")
        if rec_type == "BLANK":
            current_row += 1
            continue
            
        for col_letter, val in rec.items():
            if col_letter == "type":
                continue
            if val is not None:
                col_idx = openpyxl.utils.column_index_from_string(col_letter)
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = val
                
                if hasattr(val, "strftime"):
                    cell.number_format = "yyyy-mm-dd"
                elif isinstance(val, (int, float)) and col_letter in ["I", "J", "M", "N"]:
                    cell.number_format = "#,##0.00"
                    
        current_row += 1
        
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

# =============================================================================
# 4. STREAMLIT USER INTERFACE
# =============================================================================
st.set_page_config(
    page_title="Autoline Excel Interface Bridge",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #546E7A;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Base directories
base_dir = os.path.dirname(__file__)
config_path = os.path.join(base_dir, "config", "default_settings.json")
template_path = os.path.join(base_dir, "data", "output_template.xlsx")

def load_settings():
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()

if "settings" not in st.session_state:
    st.session_state.settings = load_settings()

# Sidebar: GL Configuration
with st.sidebar:
    st.header("⚙️ การตั้งค่าผังบัญชี (GL Mapping)")
    st.caption("ปรับเปลี่ยนรหัสบัญชีและค่าคงที่สำหรับส่งออก Autoline")
    
    with st.expander("📌 ข้อมูลทั่วไป (General)", expanded=True):
        st.session_state.settings["journal_desc"] = st.text_input(
            "คำอธิบายสมุดรายวัน (JNLDESC)",
            value=st.session_state.settings.get("journal_desc", "After Sale")
        )
        st.session_state.settings["src_branch"] = st.text_input(
            "รหัสสาขาเริ่มต้น (SRCBRANCH)",
            value=st.session_state.settings.get("src_branch", "0001")
        )
        st.session_state.settings["subaccount_default"] = st.text_input(
            "รหัสลูกหนี้เริ่มต้น (SUBACCOUNT)",
            value=st.session_state.settings.get("subaccount_default", "A0011")
        )
        
    with st.expander("🏦 ลูกหนี้การค้า (AR Debit Line)", expanded=False):
        st.session_state.settings["ar_gl_code"] = st.text_input(
            "ผังบัญชีลูกหนี้ (GLCODE)",
            value=str(st.session_state.settings.get("ar_gl_code", "11311001"))
        )
        st.session_state.settings["ar_department"] = st.text_input(
            "แผนกลูกหนี้ (DEPARTMENT)",
            value=str(st.session_state.settings.get("ar_department", "0000"))
        )
        
    with st.expander("🔩 หมวดอะไหล่ (Parts - P)", expanded=False):
        p_cfg = st.session_state.settings.get("parts", {})
        p_cfg["gl_code"] = st.text_input("GL Code อะไหล่", value=str(p_cfg.get("gl_code", "41211004")))
        p_cfg["department"] = st.text_input("แผนกอะไหล่", value=str(p_cfg.get("department", "4002")))
        p_cfg["partfran"] = st.text_input("PARTFRAN", value=str(p_cfg.get("partfran", "J")))
        p_cfg["partprod"] = st.text_input("PARTPROD", value=str(p_cfg.get("partprod", "A")))
        st.session_state.settings["parts"] = p_cfg
        
    with st.expander("🔧 หมวดค่าแรง (Labor - L)", expanded=False):
        l_cfg = st.session_state.settings.get("labor", {})
        l_cfg["gl_code"] = st.text_input("GL Code ค่าแรง", value=str(l_cfg.get("gl_code", "41111001")))
        l_cfg["department"] = st.text_input("แผนกค่าแรง", value=str(l_cfg.get("department", "4001")))
        l_cfg["servprod"] = st.text_input("SERVPROD", value=str(l_cfg.get("servprod", "A")))
        l_cfg["servicefranc"] = st.text_input("SERVICEFRANC", value=str(l_cfg.get("servicefranc", "JEEP")))
        st.session_state.settings["labor"] = l_cfg
        
    with st.expander("🛠️ หมวดบริการภายนอก (Sublet - S)", expanded=False):
        s_cfg = st.session_state.settings.get("sublet", {})
        s_cfg["gl_code"] = st.text_input("GL Code Sublet", value=str(s_cfg.get("gl_code", "41311001")))
        s_cfg["department"] = st.text_input("แผนก Sublet", value=str(s_cfg.get("department", "4003")))
        s_cfg["servprod"] = st.text_input("SERVPROD (Sublet)", value=str(s_cfg.get("servprod", "A")))
        s_cfg["servicefranc"] = st.text_input("SERVICEFRANC (Sublet)", value=str(s_cfg.get("servicefranc", "JEEP")))
        st.session_state.settings["sublet"] = s_cfg

    if st.button("💾 บันทึกการตั้งค่า (Save Settings)", use_container_width=True):
        os.makedirs(os.path.dirname(config_path), exist_ok=True)
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(st.session_state.settings, f, ensure_ascii=False, indent=2)
            st.success("บันทึกการตั้งค่าเรียบร้อยแล้ว!")
        except Exception:
            st.success("บันทึกใน Session ปัจจุบันเรียบร้อย!")

# Main Header
st.markdown('<div class="main-header">🚗 Autoline Excel Interface Bridge</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบแปลงข้อมูลบิลบริการและอะไหล่เข้าสู่รูปแบบ Autoline (CDK) AR/AP Journal Template</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 แปลงข้อมูล & ดาวน์โหลด (Transform & Export)", "📑 จัดการ Master Data (Master Mapping)"])

with tab1:
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.subheader("1. อัปโหลดไฟล์ HEADER")
        st.caption("ไฟล์ที่มีเลขที่บิล, วันที่, รหัสลูกค้า, และยอดรวม (เช่น Input Template (HEADER).xlsx)")
        header_file = st.file_uploader("เลือกไฟล์ HEADER (.xlsx)", type=["xlsx", "xls"], key="header_upload")
        
    with col_up2:
        st.subheader("2. อัปโหลดไฟล์ DETAIL")
        st.caption("ไฟล์ที่มีรายการแยกตามหมวด P, L, S (เช่น Input Template (DETAIL).xlsx)")
        detail_file = st.file_uploader("เลือกไฟล์ DETAIL (.xlsx)", type=["xlsx", "xls"], key="detail_upload")
        
    master_file = st.file_uploader(
        "📎 (ทางเลือก) อัปโหลดไฟล์ Master Data เพื่อ Map รหัสลูกหนี้/สาขาอัตโนมัติ (.xlsx)",
        type=["xlsx", "xls"],
        key="master_upload"
    )
    master_dict = None
    if master_file:
        try:
            master_dict = parse_master_data_file(master_file)
            st.success(f"✅ โหลด Master Data สำเร็จ! (พบข้อมูลลูกค้า {len(master_dict.get('customers', {}))} ราย, สาขา {len(master_dict.get('branches', {}))} สาขา)")
        except Exception as e:
            st.warning(f"ไม่สามารถอ่านไฟล์ Master Data ได้: {str(e)}")

    if header_file and detail_file:
        try:
            with st.spinner("กำลังอ่านและตรวจสอบความถูกต้องของข้อมูล..."):
                df_header, df_detail = load_and_validate_inputs(header_file, detail_file)
                rows_to_write, summary_stats, preview_df = transform_to_autoline_data(
                    df_header, df_detail, st.session_state.settings, master_dict=master_dict
                )
                
            st.success("✅ ประมวลผลและจับคู่ข้อมูลสำเร็จ!")
            
            # KPI Dashboard
            st.markdown("### 📊 สรุปภาพรวมข้อมูลและตรวจความสมดุลทางบัญชี (Balance Check)")
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric("จำนวนใบแจ้งหนี้", f"{summary_stats['total_invoices']:,} ใบ")
            kpi2.metric("ยอดขายสุทธิ (Credit)", f"{summary_stats['total_credit']:,.2f} ฿")
            kpi3.metric("ภาษี 7% (VAT)", f"{summary_stats['total_tax']:,.2f} ฿")
            kpi4.metric("ยอดลูกหนี้รวม (Debit)", f"{summary_stats['total_debit']:,.2f} ฿")
            
            if summary_stats["is_balanced"]:
                kpi5.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta="Dr = Cr + Tax")
            else:
                kpi5.metric("สถานะดุลบัญชี", "ยอดไม่ดุล ⚠️", delta=f"Diff: {summary_stats['difference']}")

            # Export Button
            st.markdown("---")
            col_dl1, col_dl2 = st.columns([2, 1])
            with col_dl1:
                st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline")
                st.caption("ไฟล์ Excel ในโครงสร้าง Autoline AR/AP พร้อมนำเข้าทันที (DEBIT แถว 9, CREDIT แถว 10, เว้นแถว 8)")
            with col_dl2:
                excel_output = generate_output_excel(template_path, rows_to_write)
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="⬇️ ดาวน์โหลด Autoline_Import.xlsx",
                    data=excel_output,
                    file_name=f"Autoline_Import_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            # Table Preview
            st.markdown("### 🔍 ตัวอย่างข้อมูลที่แปลงแล้ว (Data Preview)")
            search_inv = st.text_input("ค้นหาเลขที่บิล (Filter by Invoice Number)", placeholder="พิมพ์เลขที่บิล เช่น 01S26080001")
            
            display_df = preview_df
            if search_inv.strip():
                display_df = preview_df[preview_df["Invoice"].astype(str).str.contains(search_inv.strip(), case=False, na=False)]
                
            st.dataframe(display_df, use_container_width=True, height=450)
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ HEADER และ DETAIL ด้านบนเพื่อเริ่มต้นใช้งาน")

with tab2:
    st.header("📑 เทมเพลตไฟล์ Master Data")
    st.markdown("""
    คุณสามารถเตรียมไฟล์ Excel Master Data สำหรับจับคู่รหัสลูกหนี้ สาขา และผังบัญชีได้ โดยไฟล์มี 3 แผ่นงาน (Sheets):
    1. **Customer_Subaccount:** จับคู่ `customer_code` หรือ `tax_ID` กับ `autoline_subaccount`
    2. **Branch_Mapping:** จับคู่ `dealer_prefix` หรือ `branch_name` กับ `autoline_branch_code`
    3. **GL_Mapping:** จับคู่ `category` (P, L, S, AR) กับ `gl_code`, `department`, `partfran`, `servprod`
    """)
    
    master_bytes = get_master_template_bytes()
    st.download_button(
        label="⬇️ ดาวน์โหลดเทมเพลต Master Data (Excel)",
        data=master_bytes,
        file_name="Master_Data_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
