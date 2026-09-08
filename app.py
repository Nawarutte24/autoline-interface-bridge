import streamlit as st
import pandas as pd
import openpyxl
import json
import os
import sys
import datetime
from io import BytesIO

# =============================================================================
# 1. FIXED RULES & CONFIGURATION (DEALERPRO -> AUTOLINE)
# =============================================================================
AUTOMATED_CONFIG = {
    "doc_code": "ARI",
    "currency": "THB",
    "tax_group": "U",
    "terms": 30,
    "subaccount_default": "A0011",
    "ar_gl_code": "11311001",
    "ar_department": "0000",
    "parts": {
        "gl_code": "41211001",  # Part -> 41211001
        "department": "4002",   # งานอะไหล่: 4002
        "tax_code": "S",
        "aftstype": "R",
        "partfran": "J",
        "partprod": "A"
    },
    "labor": {
        "gl_code": "42111001",  # Labor -> 42111001
        "department": "5002",   # งานบริการ: 5002
        "tax_code": "S",
        "aftstype": "R",
        "servprod": "S",        # SERVPROD ปรับเป็น S
        "servicefranc": "JEEP"
    },
    "sublet": {
        "gl_code": "42111001",  # Sublet -> 42111001
        "department": "5002",   # งานบริการ: 5002
        "tax_code": "S",
        "aftstype": "R",
        "servprod": "S",        # SERVPROD ปรับเป็น S
        "servicefranc": "JEEP"
    }
}

def resolve_branch_by_invoice(invoice_number):
    """
    BRANCH Mapping Rule:
    ดูจากตัวเลข 2 หลักแรกของ invoice_number:
    - ถ้าเป็น 01, 03, 04, 05, 06, 07, 08, 09 -> mapping เป็น '0001' ใน output
    - ถ้าเป็น 02 -> mapping เป็น '0002' ใน output
    """
    if not invoice_number or pd.isna(invoice_number):
        return "0001"
    inv_str = str(invoice_number).strip()
    prefix2 = inv_str[:2]
    if prefix2 == "02":
        return "0002"
    elif prefix2 in ["01", "03", "04", "05", "06", "07", "08", "09"]:
        return "0001"
    else:
        return "0001"

# =============================================================================
# 2. TEMPLATE GENERATOR
# =============================================================================
def get_base_autoline_workbook(template_path=None):
    if template_path and os.path.exists(template_path):
        try:
            return openpyxl.load_workbook(template_path)
        except Exception:
            pass
            
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    ws.cell(1, 1).value = "This template is for AR/AP documents using debit and credit value columns"
    ws.cell(1, 11).value = "Rows 1 to 6 must be copied to any new sheets. Data for import must start on line 7 of the sheet."
    
    h_headers = [
        "JNLDESC", "JNLCODE", "JNLPAGE", "JNLBATCH", "JNLITEM", "JNLDOCCODE",
        "CURRENCY", "DOCSEQ", "TOTALTAX", "TOTALVALUE", "DOCDATE", "DOCNUM",
        "ITEMREFERENCE", "MISCREFERENCE", "DELIVEREDDATE", "SUBACCOUNT", "CONTACT",
        "SUBACNAME", "SUBACPOSTCODE", "SUBACTAXNO", "TAXGROUP", "SETTTERMS",
        "TERMS", "SRCBRANCH", "CRCODE"
    ]
    for idx, h in enumerate(h_headers, 1):
        ws.cell(3, idx).value = h
        
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

# =============================================================================
# 3. PROCESSING & TRANSFORMATION ENGINE
# =============================================================================
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

def transform_to_autoline_data(df_header, df_detail):
    rows_to_write = []
    preview_records = []
    
    batch_idx = 1
    total_debit_sum = 0.0
    total_credit_sum = 0.0
    total_tax_sum = 0.0
    
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
            
        cust_name = str(h_row.get("customer_name", "")).strip() if pd.notna(h_row.get("customer_name")) else ""
        ro_no = str(h_row.get("repair_order_number", "")).strip() if pd.notna(h_row.get("repair_order_number")) else ""
        
        # Col N: (invoice_number)_(customer_name)_(repair_order_number)
        if ro_no and cust_name:
            misc_ref = f"{inv_no}_{cust_name}_{ro_no}"
        elif cust_name:
            misc_ref = f"{inv_no}_{cust_name}"
        elif ro_no:
            misc_ref = f"{inv_no}_{ro_no}"
        else:
            misc_ref = inv_no
            
        narrative = misc_ref
        
        h_nett = float(h_row.get("nett_price", 0.0))
        h_tax = float(h_row.get("sales_tax", 0.0))
        h_total = float(h_row.get("total", 0.0))
        
        subaccount = AUTOMATED_CONFIG["subaccount_default"]
        terms_val = AUTOMATED_CONFIG["terms"]
        src_branch = resolve_branch_by_invoice(inv_no)
        
        d_group = detail_by_inv.get(inv_no, pd.DataFrame())
        
        cat_sales = {}
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
            
        # Col H: DOCSEQ ยึดตาม Col F JNLDOCCODE (ARI -> SINVOICV, ARC -> SCREDITV)
        doc_code = AUTOMATED_CONFIG.get("doc_code", "ARI")
        if doc_code == "ARC":
            doc_seq = "SCREDITV"
        else:
            doc_seq = "SINVOICV"
            
        # 1. Header Row
        # Col A JNLDESC ใช้เลขที่ใบแจ้งหนี้ (inv_no)
        header_record = {
            "type": "HEADER",
            "A": inv_no,
            "B": inv_no,
            "C": 1,
            "D": batch_idx,
            "E": None,
            "F": doc_code,
            "G": AUTOMATED_CONFIG["currency"],
            "H": doc_seq,
            "I": h_tax,
            "J": h_total,
            "K": doc_date,
            "L": None,
            "M": inv_no,
            "N": misc_ref,
            "O": doc_date,
            "P": subaccount,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "U": AUTOMATED_CONFIG["tax_group"],
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
            "Doc Code": doc_code,
            "Doc Seq": doc_seq,
            "Date": doc_date.strftime("%Y-%m-%d"),
            "GL Code": "",
            "Department": "",
            "Debit": "",
            "Credit": "",
            "Tax": h_tax,
            "Total Value": h_total,
            "Narrative": misc_ref,
            "Branch": src_branch,
            "Subaccount": subaccount
        })
        
        # 2. Blank Row
        rows_to_write.append({"type": "BLANK"})
        
        # 3. Line 1: DEBIT Line (AR Control Account) -> UP FIRST
        ar_gl = AUTOMATED_CONFIG["ar_gl_code"]
        ar_dept = AUTOMATED_CONFIG["ar_department"]
        
        debit_record = {
            "type": "ITEM",
            "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
            "G": AUTOMATED_CONFIG["currency"], "H": 1,
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
            "Branch": src_branch,
            "Subaccount": subaccount
        })
        
        # 4. Line 2+: CREDIT Lines (Revenue by Category)
        line_num = 2
        for cat, amount in cat_sales.items():
            if amount <= 0:
                continue
                
            if cat == "P":
                cfg = AUTOMATED_CONFIG["parts"]
                gl_code_val = int(cfg["gl_code"])      # 41211001
                dept_val = str(cfg["department"])      # 4002 (งานอะไหล่)
                tax_code_val = cfg["tax_code"]
                aftstype_val = cfg["aftstype"]
                partfran = cfg["partfran"]
                partprod = cfg["partprod"]
                servprod = None
                servicefranc = None
            elif cat == "L":
                cfg = AUTOMATED_CONFIG["labor"]
                gl_code_val = int(cfg["gl_code"])      # 42111001
                dept_val = str(cfg["department"])      # 5002 (งานบริการ)
                tax_code_val = cfg["tax_code"]
                aftstype_val = cfg["aftstype"]
                partfran = None
                partprod = None
                servprod = cfg["servprod"]             # S
                servicefranc = cfg["servicefranc"]
            else: # S (Sublet)
                cfg = AUTOMATED_CONFIG["sublet"]
                gl_code_val = int(cfg["gl_code"])      # 42111001
                dept_val = str(cfg["department"])      # 5002 (งานบริการ)
                tax_code_val = cfg["tax_code"]
                aftstype_val = cfg["aftstype"]
                partfran = None
                partprod = None
                servprod = cfg["servprod"]             # S
                servicefranc = cfg["servicefranc"]
            
            credit_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": AUTOMATED_CONFIG["currency"], "H": line_num,
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
                "Branch": src_branch,
                "Subaccount": ""
            })
            line_num += 1
            
        total_tax_sum += h_tax
        batch_idx += 1
        # 5. เว้น 1 บรรทัดว่างหลังจบแต่ละบิล ก่อนขึ้นบิลถัดไป
        rows_to_write.append({"type": "BLANK"})
        
    # ลบแถวว่างส่วนเกินท้ายสุดหลังบิลสุดท้าย (ถ้ามี)
    if rows_to_write and rows_to_write[-1].get("type") == "BLANK":
        rows_to_write.pop()
        
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
# 4. STREAMLIT USER INTERFACE (MINIMALIST & ULTRA CLEAN)
# =============================================================================
st.set_page_config(
    page_title="Autoline Interface Bridge",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.3rem;
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

base_dir = os.path.dirname(__file__)
template_path = os.path.join(base_dir, "data", "output_template.xlsx")

# Main Header
st.markdown('<div class="main-header">🚗 Autoline Interface Bridge</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">ระบบแปลงข้อมูลบิลบริการและอะไหล่อัตโนมัติ เข้าสู่แบบฟอร์ม Autoline AR/AP Journal Import</div>', unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)
with col_up1:
    st.subheader("1. อัปโหลดไฟล์ HEADER")
    st.caption("ไฟล์ที่มีเลขที่บิล, วันที่, ข้อมูลลูกค้า, และยอดรวม")
    header_file = st.file_uploader("เลือกไฟล์ HEADER (.xlsx)", type=["xlsx", "xls"], key="header_upload")
    
with col_up2:
    st.subheader("2. อัปโหลดไฟล์ DETAIL")
    st.caption("ไฟล์ที่มีรายการแยกตามหมวด อะไหล่ P, ค่าแรง L, บริการ S")
    detail_file = st.file_uploader("เลือกไฟล์ DETAIL (.xlsx)", type=["xlsx", "xls"], key="detail_upload")

if header_file and detail_file:
    try:
        with st.spinner("กำลังแปลงข้อมูลและคำนวณดุลบัญชีอัตโนมัติ..."):
            df_header, df_detail = load_and_validate_inputs(header_file, detail_file)
            rows_to_write, summary_stats, preview_df = transform_to_autoline_data(
                df_header, df_detail
            )
            
        st.success("✅ แปลงข้อมูลสำเร็จเรียบร้อย!")
        
        # KPI Dashboard
        st.markdown("### 📊 สรุปตัวเลขและความสมดุลทางบัญชี (Balance Verification)")
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("จำนวนใบแจ้งหนี้", f"{summary_stats['total_invoices']:,} ใบ")
        kpi2.metric("ยอดขายสุทธิ (Credit)", f"{summary_stats['total_credit']:,.2f} ฿")
        kpi3.metric("ภาษี 7% (VAT)", f"{summary_stats['total_tax']:,.2f} ฿")
        kpi4.metric("ยอดลูกหนี้รวม (Debit)", f"{summary_stats['total_debit']:,.2f} ฿")
        
        if summary_stats["is_balanced"]:
            kpi5.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta="Dr = Cr + Tax")
        else:
            kpi5.metric("สถานะดุลบัญชี", "ยอดไม่ดุล ⚠️", delta=f"Diff: {summary_stats['difference']}")

        # Download Button
        st.markdown("---")
        col_dl1, col_dl2 = st.columns([2, 1])
        with col_dl1:
            st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline")
            st.caption("ไฟล์ Excel ในโครงสร้าง Autoline AR/AP (DEBIT แถว 9, CREDIT แถว 10, เว้นแถว 8) พร้อมนำเข้าทันที")
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
        st.markdown("### 🔍 ตรวจสอบข้อมูลก่อนดาวน์โหลด (Data Preview)")
        search_inv = st.text_input("ค้นหาเลขที่บิล", placeholder="พิมพ์เลขที่บิล เช่น 01S26080001")
        
        display_df = preview_df
        if search_inv.strip():
            display_df = preview_df[preview_df["Invoice"].astype(str).str.contains(search_inv.strip(), case=False, na=False)]
            
        st.dataframe(display_df, use_container_width=True, height=450)
        
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
else:
    st.info("💡 กรุณาอัปโหลดไฟล์ HEADER และ DETAIL ด้านบนเพื่อเริ่มต้นแปลงข้อมูล")
