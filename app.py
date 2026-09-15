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
    "tax_group_non_vat": "OS",
    "tax_code_non_vat": "O",
    "terms": 30,
    "subaccount_default": "A0011",
    "credit_reason_code": "00000004",  # Incorrect (Credit Reason for Autoline ARC)
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

def resolve_branch(h_row=None, invoice_number=None):
    """
    BRANCH Mapping Rule:
    - ถ้าเป็น dealer_prefix == 'KSN' หรือ branch_name == 'KasetNawamin' -> '0002'
    - หรือถ้า invoice_number ขึ้นต้นด้วย '02' -> '0002'
    - นอกเหนือจากนี้ -> '0001'
    """
    if h_row is not None:
        pfx = str(h_row.get("dealer_prefix", "")).strip().upper()
        b_name = str(h_row.get("branch_name", "")).strip().lower()
        if pfx == "KSN" or "kasetnawamin" in b_name:
            return "0002"
            
    inv_str = ""
    if invoice_number:
        inv_str = str(invoice_number).strip()
    elif h_row is not None:
        inv_str = str(h_row.get("invoice_number", "")).strip()
        
    if inv_str.startswith("02"):
        return "0002"
        
    return "0001"

def resolve_branch_by_invoice(invoice_number):
    return resolve_branch(invoice_number=invoice_number)

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
    
    if "Preview_Document" in df_header.columns:
        df_header["Preview_Document"] = df_header["Preview_Document"].fillna("").astype(str).str.strip()
        df_header.loc[df_header["Preview_Document"].str.lower() == "nan", "Preview_Document"] = ""
    else:
        df_header["Preview_Document"] = ""
        
    if "transaction_type" in df_header.columns:
        df_header["transaction_type"] = df_header["transaction_type"].fillna("").astype(str).str.strip().str.upper()
    else:
        df_header["transaction_type"] = ""
        
    for c in ["customer_name", "repair_order_number"]:
        if c in df_header.columns:
            df_header[c] = df_header[c].fillna("").astype(str).str.strip()
            df_header.loc[df_header[c].str.lower() == "nan", c] = ""
            
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
    
    inv_revenue = 0.0
    inv_tax = 0.0
    inv_ar = 0.0
    
    cn_revenue = 0.0
    cn_tax = 0.0
    cn_ar = 0.0
    
    count_inv = 0
    count_cn = 0
    unbalanced_docs = []
    
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
            
        trans_type = str(h_row.get("transaction_type", "")).strip().upper()
        h_nett = float(h_row.get("nett_price", 0.0))
        h_tax = float(h_row.get("sales_tax", 0.0))
        h_total = float(h_row.get("total", 0.0))
        prev_doc = str(h_row.get("Preview_Document", "")).strip() if pd.notna(h_row.get("Preview_Document")) else ""
        if prev_doc.lower() == "nan":
            prev_doc = ""
            
        # ตรวจสอบว่าเป็นใบลดหนี้ (Credit Note - CN) หรือไม่
        is_cn = (trans_type == "C") or (h_total < 0)
        
        doc_code = "ARC" if is_cn else "ARI"
        doc_seq = "SCREDITV" if is_cn else "SINVOICV"
        
        doc_nett = round(abs(h_nett), 2)
        doc_tax = round(abs(h_tax), 2)
        doc_total = round(abs(h_total), 2)
        
        # ตรวจสอบว่าเอกสารมีภาษีมูลค่าเพิ่มหรือไม่
        # หากไม่มี VAT: TAXGROUP เป็น OS และ TAXCODE เป็น O
        has_vat = (doc_tax > 0)
        tax_group_val = AUTOMATED_CONFIG["tax_group"] if has_vat else AUTOMATED_CONFIG.get("tax_group_non_vat", "OS")
        crcode_val = AUTOMATED_CONFIG.get("credit_reason_code", "00000004") if is_cn else None
        
        cust_name = str(h_row.get("customer_name", "")).strip() if pd.notna(h_row.get("customer_name")) else ""
        if cust_name.lower() == "nan":
            cust_name = ""
        ro_no = str(h_row.get("repair_order_number", "")).strip() if pd.notna(h_row.get("repair_order_number")) else ""
        if ro_no.lower() == "nan":
            ro_no = ""
        
        # Col N: MISCREFERENCE
        # สำหรับ CN: รวม prev_doc ด้วย -> f"{inv_no}_{prev_doc}_{cust_name}_{ro_no}"
        # สำหรับ Invoice ปกติ: f"{inv_no}_{cust_name}_{ro_no}"
        if is_cn and prev_doc:
            parts = [inv_no, prev_doc, cust_name, ro_no]
        else:
            parts = [inv_no, cust_name, ro_no]
        misc_ref = "_".join([p for p in parts if p])
        narrative = misc_ref
        
        subaccount = AUTOMATED_CONFIG["subaccount_default"]
        terms_val = AUTOMATED_CONFIG["terms"]
        src_branch = resolve_branch(h_row=h_row, invoice_number=inv_no)
        
        # Detail lookup: สำหรับ CN ให้อ้างอิง Preview_Document ใน Detail ก่อน
        lookup_inv = prev_doc if (is_cn and prev_doc and prev_doc in detail_by_inv) else inv_no
        d_group = detail_by_inv.get(lookup_inv, pd.DataFrame())
        
        cat_sales = {}
        if not d_group.empty:
            cat_sums = d_group.groupby("category")["sales"].sum().to_dict()
            cat_tax_sums = d_group.groupby("category")["sales_tax"].sum().to_dict()
            
            valid_cats = {k: v for k, v in cat_sums.items() if v > 0 and k in ["P", "L", "S", "C"]}
            valid_tax_cats = {k: v for k, v in cat_tax_sums.items() if v > 0 and k in ["P", "L", "S", "C"]}
            
            # Map C -> S (Sublet)
            target_tax = {}
            for k, v in valid_tax_cats.items():
                mapped_k = "S" if k == "C" else k
                target_tax[mapped_k] = target_tax.get(mapped_k, 0.0) + v
                
            target_cats = {}
            for k, v in valid_cats.items():
                mapped_k = "S" if k == "C" else k
                target_cats[mapped_k] = target_cats.get(mapped_k, 0.0) + v
            
            if len(target_cats) == 1:
                cat_name = list(target_cats.keys())[0]
                cat_sales[cat_name] = doc_nett
            elif len(target_cats) > 1:
                tot_tax = sum(target_tax.values())
                if tot_tax > 0:
                    allocated = 0.0
                    c_list = list(target_tax.keys())
                    for c in c_list[:-1]:
                        c_amt = round(doc_nett * (target_tax[c] / tot_tax), 2)
                        cat_sales[c] = c_amt
                        allocated += c_amt
                    cat_sales[c_list[-1]] = round(doc_nett - allocated, 2)
                else:
                    tot_s = sum(target_cats.values())
                    allocated = 0.0
                    c_list = list(target_cats.keys())
                    for c in c_list[:-1]:
                        c_amt = round(doc_nett * (target_cats[c] / tot_s), 2)
                        cat_sales[c] = c_amt
                        allocated += c_amt
                    cat_sales[c_list[-1]] = round(doc_nett - allocated, 2)
                    
        # Fallback to Header parts_amount / labor_amount / other_amount
        if not cat_sales:
            p_amt = abs(float(h_row.get("parts_amount", 0.0)))
            l_amt = abs(float(h_row.get("labor_amount", 0.0)))
            s_amt = abs(float(h_row.get("other_amount", 0.0)))
            tot_pl = p_amt + l_amt + s_amt
            if tot_pl > 0:
                sub_cats = []
                if p_amt > 0: sub_cats.append(("P", p_amt))
                if l_amt > 0: sub_cats.append(("L", l_amt))
                if s_amt > 0: sub_cats.append(("S", s_amt))
                if len(sub_cats) == 1:
                    cat_sales[sub_cats[0][0]] = doc_nett
                else:
                    allocated = 0.0
                    for c_name, c_val in sub_cats[:-1]:
                        c_amt = round(doc_nett * (c_val / tot_pl), 2)
                        cat_sales[c_name] = c_amt
                        allocated += c_amt
                    cat_sales[sub_cats[-1][0]] = round(doc_nett - allocated, 2)
            else:
                cat_sales["P"] = doc_nett
            
        # ตรวจสอบความสมดุลของเอกสารแต่ละใบ (Doc Total == Items + Tax)
        sum_items = round(sum(cat_sales.values()), 2)
        if abs(doc_total - (sum_items + doc_tax)) > 0.05:
            unbalanced_docs.append((inv_no, doc_total, sum_items, doc_tax))
            
        if is_cn:
            count_cn += 1
            cn_revenue += sum_items
            cn_tax += doc_tax
            cn_ar += doc_total
        else:
            count_inv += 1
            inv_revenue += sum_items
            inv_tax += doc_tax
            inv_ar += doc_total
            
        # 1. Header Row
        # Col A JNLDESC ใช้เลขที่เอกสาร (inv_no)
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
            "I": doc_tax,
            "J": doc_total,
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
            "U": tax_group_val,
            "V": None,
            "W": terms_val,
            "X": src_branch,
            "Y": crcode_val
        }
        rows_to_write.append(header_record)
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": f"HEADER ({doc_code})",
            "Batch": batch_idx,
            "Line": "",
            "Doc Code": doc_code,
            "Doc Seq": doc_seq,
            "Date": doc_date.strftime("%Y-%m-%d"),
            "GL Code": "",
            "Department": "",
            "Debit": "",
            "Credit": "",
            "Tax": doc_tax,
            "Total Value": doc_total,
            "Narrative": misc_ref,
            "Branch": src_branch,
            "Subaccount": subaccount
        })
        
        # 2. Blank Row
        rows_to_write.append({"type": "BLANK"})
        
        # 3. Line 1: AR Control Account
        # Normal Invoice -> DEBIT (ลูกหนี้เพิ่ม)
        # Credit Note (CN) -> CREDIT (ลูกหนี้ลด)
        ar_gl = AUTOMATED_CONFIG["ar_gl_code"]
        ar_dept = AUTOMATED_CONFIG["ar_department"]
        
        ar_debit = None if is_cn else doc_total
        ar_credit = doc_total if is_cn else None
        
        ar_record = {
            "type": "ITEM",
            "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
            "G": AUTOMATED_CONFIG["currency"], "H": 1,
            "I": src_branch, "J": ar_dept, "K": int(ar_gl), "L": None,
            "M": ar_debit, "N": ar_credit, "O": narrative, "P": None, "Q": None,
            "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
            "X": None, "Y": None, "Z": None, "AA": None, "AB": None, "AC": None, "AD": None
        }
        rows_to_write.append(ar_record)
        if ar_debit is not None:
            total_debit_sum += ar_debit
        if ar_credit is not None:
            total_credit_sum += ar_credit
            
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": "CREDIT (AR)" if is_cn else "DEBIT (AR)",
            "Batch": batch_idx,
            "Line": 1,
            "Doc Code": "",
            "Doc Seq": "",
            "Date": doc_date.strftime("%Y-%m-%d"),
            "GL Code": ar_gl,
            "Department": ar_dept,
            "Debit": ar_debit if ar_debit is not None else "",
            "Credit": ar_credit if ar_credit is not None else "",
            "Tax": "",
            "Total Value": "",
            "Narrative": narrative,
            "Branch": src_branch,
            "Subaccount": subaccount
        })
        
        # 4. Line 2+: Revenue Lines
        # Normal Invoice -> CREDIT (รายได้เพิ่ม)
        # Credit Note (CN) -> DEBIT (รายได้ลด)
        line_num = 2
        for cat, amount in cat_sales.items():
            if amount <= 0:
                continue
                
            if cat == "P":
                cfg = AUTOMATED_CONFIG["parts"]
                gl_code_val = int(cfg["gl_code"])      # 41211001
                dept_val = str(cfg["department"])      # 4002 (งานอะไหล่)
                tax_code_val = cfg["tax_code"] if has_vat else AUTOMATED_CONFIG.get("tax_code_non_vat", "O")
                aftstype_val = cfg["aftstype"]
                partfran = cfg["partfran"]
                partprod = cfg["partprod"]
                servprod = None
                servicefranc = None
            elif cat == "L":
                cfg = AUTOMATED_CONFIG["labor"]
                gl_code_val = int(cfg["gl_code"])      # 42111001
                dept_val = str(cfg["department"])      # 5002 (งานบริการ)
                tax_code_val = cfg["tax_code"] if has_vat else AUTOMATED_CONFIG.get("tax_code_non_vat", "O")
                aftstype_val = cfg["aftstype"]
                partfran = None
                partprod = None
                servprod = cfg["servprod"]             # S
                servicefranc = cfg["servicefranc"]
            else: # S (Sublet)
                cfg = AUTOMATED_CONFIG["sublet"]
                gl_code_val = int(cfg["gl_code"])      # 42111001
                dept_val = str(cfg["department"])      # 5002 (งานบริการ)
                tax_code_val = cfg["tax_code"] if has_vat else AUTOMATED_CONFIG.get("tax_code_non_vat", "O")
                aftstype_val = cfg["aftstype"]
                partfran = None
                partprod = None
                servprod = cfg["servprod"]             # S
                servicefranc = cfg["servicefranc"]
            
            rev_debit = amount if is_cn else None
            rev_credit = None if is_cn else amount
            
            rev_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": AUTOMATED_CONFIG["currency"], "H": line_num,
                "I": src_branch, "J": dept_val, "K": gl_code_val, "L": None,
                "M": rev_debit, "N": rev_credit, "O": narrative, "P": None, "Q": tax_code_val,
                "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
                "X": aftstype_val, "Y": None, "Z": partfran, "AA": servprod, "AB": partprod, "AC": servicefranc, "AD": None
            }
            rows_to_write.append(rev_record)
            if rev_debit is not None:
                total_debit_sum += rev_debit
            if rev_credit is not None:
                total_credit_sum += rev_credit
            
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": f"DEBIT ({cat})" if is_cn else f"CREDIT ({cat})",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d"),
                "GL Code": str(gl_code_val),
                "Department": dept_val,
                "Debit": rev_debit if rev_debit is not None else "",
                "Credit": rev_credit if rev_credit is not None else "",
                "Tax": tax_code_val,
                "Total Value": "",
                "Narrative": narrative,
                "Branch": src_branch,
                "Subaccount": ""
            })
            line_num += 1
            
        batch_idx += 1
        # 5. เว้น 1 บรรทัดว่างหลังจบแต่ละบิล ก่อนขึ้นบิลถัดไป
        rows_to_write.append({"type": "BLANK"})
        
    # ลบแถวว่างส่วนเกินท้ายสุดหลังบิลสุดท้าย (ถ้ามี)
    if rows_to_write and rows_to_write[-1].get("type") == "BLANK":
        rows_to_write.pop()
        
    net_revenue = round(inv_revenue - cn_revenue, 2)
    net_tax = round(inv_tax - cn_tax, 2)
    net_ar = round(inv_ar - cn_ar, 2)
    
    summary_stats = {
        "total_documents": len(df_header),
        "count_inv": count_inv,
        "count_cn": count_cn,
        "total_debit": round(total_debit_sum, 2),
        "total_credit": round(total_credit_sum, 2),
        "net_revenue": net_revenue,
        "net_tax": net_tax,
        "net_ar": net_ar,
        "inv_revenue": round(inv_revenue, 2),
        "inv_tax": round(inv_tax, 2),
        "inv_ar": round(inv_ar, 2),
        "cn_revenue": round(cn_revenue, 2),
        "cn_tax": round(cn_tax, 2),
        "cn_ar": round(cn_ar, 2),
        "is_balanced": len(unbalanced_docs) == 0,
        "unbalanced_count": len(unbalanced_docs),
        "difference": round(abs((total_debit_sum - total_credit_sum) - (inv_tax - cn_tax)), 2)
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
                elif col_letter == "Y" and val is not None:
                    cell.number_format = "@"
                    
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
        
        cn_label = f"Inv: {summary_stats['count_inv']:,} | CN: {summary_stats['count_cn']:,}" if summary_stats['count_cn'] > 0 else f"{summary_stats['count_inv']:,} ใบ"
        kpi1.metric("จำนวนเอกสารทั้งหมด", f"{summary_stats['total_documents']:,} ฉบับ", delta=cn_label if summary_stats['count_cn'] > 0 else None)
        
        kpi2.metric(
            "ยอดขายสุทธิ (Revenue)", 
            f"{summary_stats['net_revenue']:,.2f} ฿",
            delta=f"CN: -{summary_stats['cn_revenue']:,.2f} ฿" if summary_stats['count_cn'] > 0 else None
        )
        kpi3.metric(
            "ภาษี 7% (VAT)", 
            f"{summary_stats['net_tax']:,.2f} ฿",
            delta=f"CN: -{summary_stats['cn_tax']:,.2f} ฿" if summary_stats['count_cn'] > 0 else None
        )
        kpi4.metric(
            "ยอดลูกหนี้สุทธิ (AR)", 
            f"{summary_stats['net_ar']:,.2f} ฿",
            delta=f"CN: -{summary_stats['cn_ar']:,.2f} ฿" if summary_stats['count_cn'] > 0 else None
        )
        
        if summary_stats["is_balanced"]:
            kpi5.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta=f"ดุลครบทุกบิล ({summary_stats['total_documents']}/{summary_stats['total_documents']})")
        else:
            kpi5.metric("สถานะดุลบัญชี", "พบเอกสารไม่ดุล ⚠️", delta=f"{summary_stats['unbalanced_count']} ฉบับ")

        # Download Button
        st.markdown("---")
        col_dl1, col_dl2 = st.columns([2, 1])
        with col_dl1:
            st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline")
            st.caption("ไฟล์ Excel ในโครงสร้าง Autoline AR/AP (รองรับทั้งบิล ARI และใบลดหนี้ ARC พร้อมเว้นบรรทัดตามมาตรฐาน)")
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
        search_kw = st.text_input("ค้นหาเอกสาร (เลขที่บิล, ใบลดหนี้, เลขที่อ้างอิง, ชื่อลูกค้า หรือ Doc Code)", placeholder="เช่น 01SC26050001, ARC, HA0027, SINVOICV")
        
        display_df = preview_df
        if search_kw.strip():
            kw = search_kw.strip().lower()
            mask = (
                preview_df["Invoice"].astype(str).str.lower().str.contains(kw, na=False) |
                preview_df["Narrative"].astype(str).str.lower().str.contains(kw, na=False) |
                preview_df["Doc Code"].astype(str).str.lower().str.contains(kw, na=False) |
                preview_df["Doc Seq"].astype(str).str.lower().str.contains(kw, na=False) |
                preview_df["Row Type"].astype(str).str.lower().str.contains(kw, na=False)
            )
            display_df = preview_df[mask]
            
        st.dataframe(display_df, use_container_width=True, height=450)
        
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
else:
    st.info("💡 กรุณาอัปโหลดไฟล์ HEADER และ DETAIL ด้านบนเพื่อเริ่มต้นแปลงข้อมูล")
