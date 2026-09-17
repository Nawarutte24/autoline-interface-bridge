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
    "subaccount_default": "X0004",
    "credit_reason_code": "00000004",  # Incorrect (Credit Reason for Autoline ARC)
    "ar_gl_code": "11311001",
    "ar_department": "0000",
    "parts_cogs_gl_code": "51211006",  # บัญชีต้นทุนขายอะไหล่ (COGS)
    "parts_cogs_tax_code": "O",        # Non-VAT (0%)
    "parts_cogs_aftstype": "R",
    "parts_cogs_partfran": "J",
    "parts_cogs_partprod": "A",
    "inventory_gl_code": "11511112",   # บัญชีสต็อกอะไหล่ (Inventory)
    "inventory_department": "0000",    # แผนกสต็อก (Control Account)
    "inventory_tax_code": "O",         # Non-VAT (0%) ป้องกันไม่ให้คิดภาษีซ้ำซ้อน
    "inventory_manufact": "JEEP",      # Fixed ค่าคงที่
    "inventory_model": "JEEP-G6",      # Fixed ค่าคงที่
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
def read_input_file(file_obj):
    """
    Smart File Reader:
    - รองรับทั้งไฟล์ Excel (.xlsx, .xls) และ CSV (.csv)
    - ตรวจจับและรองรับการเข้ารหัสภาษาไทยอัตโนมัติ (utf-8-sig, utf-8, cp874, tis-620)
    - กำหนด dtype=str สำหรับ CSV เพื่อรักษาเลข 0 ด้านหน้า (เช่น รหัสสาขา '009', '02')
    """
    file_name = getattr(file_obj, "name", "")
    if isinstance(file_obj, str):
        file_name = file_obj
        
    is_csv = file_name.lower().endswith(".csv")
    
    if is_csv:
        encodings = ["utf-8-sig", "utf-8", "cp874", "tis-620"]
        for enc in encodings:
            try:
                if hasattr(file_obj, "seek"):
                    file_obj.seek(0)
                return pd.read_csv(file_obj, encoding=enc, dtype=str)
            except Exception:
                continue
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
        return pd.read_csv(file_obj, dtype=str)
    else:
        try:
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)
            return pd.read_excel(file_obj)
        except Exception:
            # Fallback กรณีไฟล์ CSV แต่เปลี่ยนนามสกุล หรือไม่มีนามสกุล
            for enc in ["utf-8-sig", "utf-8", "cp874", "tis-620"]:
                try:
                    if hasattr(file_obj, "seek"):
                        file_obj.seek(0)
                    return pd.read_csv(file_obj, encoding=enc, dtype=str)
                except Exception:
                    continue
            raise

def load_and_validate_inputs(header_file, detail_file):
    df_header = read_input_file(header_file)
    df_detail = read_input_file(detail_file)
        
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
    
    # ไม่นำเข้าบิลที่ขึ้นต้นด้วย CA หรือ HA
    excluded_prefixes = ("CA", "HA")
    df_header = df_header[~df_header["invoice_number"].str.upper().str.startswith(excluded_prefixes)].copy()
    df_detail = df_detail[~df_detail["invoice_number"].str.upper().str.startswith(excluded_prefixes)].copy()
    
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
    
    for col in ["sales", "sales_tax", "nett_price", "discount_amount", "cost_of_sales"]:
        if col in df_detail.columns:
            df_detail[col] = pd.to_numeric(df_detail[col], errors="coerce").fillna(0.0)
            
    for col in ["sales", "sales_tax", "total", "labor_amount", "parts_amount", "other_amount", "nett_price", "cost_of_sales"]:
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
    
    total_parts_cogs = 0.0
    
    count_inv = 0
    count_cn = 0
    unbalanced_docs = []
    
    detail_by_inv = {inv: grp for inv, grp in df_detail.groupby("invoice_number")}
        
    for _, h_row in df_header.iterrows():
        inv_no = str(h_row["invoice_number"]).strip()
        if not inv_no or inv_no.upper().startswith(("CA", "HA")):
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
            
        # 5. Line COGS & Inventory (สร้างคู่ขนานเมื่อมีการขายอะไหล่และมีต้นทุน parts_cost > 0)
        parts_cost = 0.0
        if not d_group.empty and "cost_of_sales" in d_group.columns:
            parts_cost = float(d_group[d_group["category"] == "P"]["cost_of_sales"].sum())
        if parts_cost == 0.0 and cat_sales.get("P", 0.0) > 0:
            parts_cost = float(h_row.get("cost_of_sales", 0.0))
        parts_cost = round(abs(parts_cost), 2)
        
        if parts_cost > 0:
            total_parts_cogs += parts_cost
            cogs_gl = int(AUTOMATED_CONFIG.get("parts_cogs_gl_code", "51211006"))
            cogs_dept = AUTOMATED_CONFIG["parts"]["department"]
            cogs_tax = AUTOMATED_CONFIG.get("parts_cogs_tax_code", "O")
            cogs_afts = AUTOMATED_CONFIG.get("parts_cogs_aftstype", "R")
            cogs_fran = AUTOMATED_CONFIG.get("parts_cogs_partfran", "J")
            cogs_prod = AUTOMATED_CONFIG.get("parts_cogs_partprod", "A")
            
            # Normal Invoice (ARI): DEBIT COGS, CREDIT Stock
            # Credit Note (ARC): CREDIT COGS, DEBIT Stock
            cogs_debit = parts_cost if not is_cn else None
            cogs_credit = None if not is_cn else parts_cost
            
            cogs_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": AUTOMATED_CONFIG["currency"], "H": line_num,
                "I": src_branch, "J": cogs_dept, "K": cogs_gl, "L": None,
                "M": cogs_debit, "N": cogs_credit, "O": narrative, "P": None, "Q": cogs_tax,
                "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
                "X": cogs_afts, "Y": None, "Z": cogs_fran, "AA": None, "AB": cogs_prod, "AC": None, "AD": None
            }
            rows_to_write.append(cogs_record)
            if cogs_debit is not None:
                total_debit_sum += cogs_debit
            if cogs_credit is not None:
                total_credit_sum += cogs_credit
                
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": "CREDIT (COGS)" if is_cn else "DEBIT (COGS)",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d"),
                "GL Code": str(cogs_gl),
                "Department": cogs_dept,
                "Debit": cogs_debit if cogs_debit is not None else "",
                "Credit": cogs_credit if cogs_credit is not None else "",
                "Tax": cogs_tax,
                "Total Value": "",
                "Narrative": narrative,
                "Branch": src_branch,
                "Subaccount": ""
            })
            line_num += 1
            
            # Line Stock / Inventory
            inv_gl = int(AUTOMATED_CONFIG.get("inventory_gl_code", "11511112"))
            inv_dept = AUTOMATED_CONFIG.get("inventory_department", "0000")
            stock_tax_code = AUTOMATED_CONFIG.get("inventory_tax_code", "O")
            stock_mfg = AUTOMATED_CONFIG.get("inventory_manufact", "JEEP")
            stock_model = AUTOMATED_CONFIG.get("inventory_model", "JEEP-G6")
            
            stock_debit = parts_cost if is_cn else None
            stock_credit = None if is_cn else parts_cost
            
            stock_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": AUTOMATED_CONFIG["currency"], "H": line_num,
                "I": src_branch, "J": inv_dept, "K": inv_gl, "L": None,
                "M": stock_debit, "N": stock_credit, "O": narrative, "P": None, "Q": stock_tax_code,
                "R": None, "S": None, "T": None, "U": stock_mfg, "V": stock_model, "W": None,
                "X": None, "Y": None, "Z": None, "AA": None, "AB": None, "AC": None, "AD": None
            }
            rows_to_write.append(stock_record)
            if stock_debit is not None:
                total_debit_sum += stock_debit
            if stock_credit is not None:
                total_credit_sum += stock_credit
                
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": "DEBIT (Stock)" if is_cn else "CREDIT (Stock)",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d"),
                "GL Code": str(inv_gl),
                "Department": inv_dept,
                "Debit": stock_debit if stock_debit is not None else "",
                "Credit": stock_credit if stock_credit is not None else "",
                "Tax": stock_tax_code,
                "Total Value": "",
                "Narrative": narrative,
                "Branch": src_branch,
                "Subaccount": ""
            })
            line_num += 1
            
        batch_idx += 1
        # 6. เว้น 1 บรรทัดว่างหลังจบแต่ละบิล ก่อนขึ้นบิลถัดไป
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
        "total_parts_cogs": round(total_parts_cogs, 2),
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
# 3.5. SALES TRANSFORMER (VEHICLE SALES: 0XD / 0XDC / 0XWG / 0XWGCN)
# =============================================================================
DEFAULT_SALES_CONFIG = {
    "subaccount_default": "X0003",
    "subaccount_finance": "ARCODE FINANCE",
    "branch_default": "0001",
    "currency": "THB",
    "tax_code": "S",
    "terms": 30,
    "categories": {
        "DEPOSIT": {
            "name": "เงินจองรถยนต์",
            "ar_gl": "11314001",      # ยอดเงินจองรวม Vat
            "ar_dept": "0000",
            "rev_gl": "21931104",     # เงินรับล่วงหน้าค่าจองก่อน Vat
            "rev_dept": "0000",
            "subaccount": "X0003"
        },
        "RED_PLATE": {
            "name": "มัดจำป้ายแดง",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "21931003",     # มัดจำป้ายแดงก่อน Vat
            "rev_dept": "0000",
            "subaccount": "X0003"
        },
        "REGISTRATION": {
            "name": "รายได้ค่าจดทะเบียน",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "49121004",     # รายได้ค่าจด/ค่าดำเนินการก่อน Vat
            "rev_dept": "2002",
            "subaccount": "X0003"
        },
        "ACCESSORIES": {
            "name": "รายได้ค่าอุปกรณ์ตกแต่งต่างๆ",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "41111003",     # รายได้ค่าอุปกรณ์ต่างๆก่อน Vat
            "rev_dept": "2002",
            "subaccount": "X0003"
        },
        "COMM_FINANCE": {
            "name": "รายได้ค่าคอมไฟแนนซ์",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "49121001",     # รายได้ commission finance ก่อน Vat
            "rev_dept": "2002",
            "subaccount": "ARCODE FINANCE"
        },
        "CAR_SALE_CASH": {
            "name": "ขายรถยนต์ (ขายสด)",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "41111001",     # รายได้ขายรถยอดก่อน Vat
            "rev_dept": "2002",
            "subaccount": "X0003"
        },
        "CAR_SALE_CREDIT": {
            "name": "ขายรถยนต์ (ขายเชื่อ/ไฟแนนซ์)",
            "ar_gl": "11311001",      # ลูกหนี้การค้าไฟแนนซ์
            "ar_dept": "0000",
            "rev_gl": "41111001",     # รายได้ขายรถยอดก่อน Vat
            "rev_dept": "2002",
            "subaccount": "ARCODE FINANCE"
        }
    }
}

FINANCE_KEYWORDS = [
    'ธนาคาร', 'ลิสซิ่ง', 'ลีสซิ่ง', 'แคปปิตอล', 'ไฟแนนซ์',
    'bank', 'leasing', 'capital', 'finance', 'auto'
]

def resolve_sales_branch(invoice_number):
    inv_str = str(invoice_number or "").strip()
    if inv_str.startswith("02"):
        return "0002"
    return "0001"

def is_finance_customer(customer_name):
    c_str = str(customer_name or "").lower().strip()
    return any(k in c_str for k in FINANCE_KEYWORDS)

def classify_description_text(desc):
    d = str(desc or '').lower().strip()
    if not d:
        return 'DEPOSIT'
        
    if any(k in d for k in ['ป้ายแดง', 'ป้ายเเดง', 'มัดจำป้าย']):
        return 'RED_PLATE'
        
    if any(k in d for k in ['คอมไฟแนนซ์', 'ค่าคอมไฟแนนซ์', 'คอมมิชชั่นไฟแนนซ์', 'ค่าคอมมิชชั่นไฟแนนซ์', 'comm', 'finan']):
        return 'COMM_FINANCE'
        
    if any(k in d for k in ['จดทะเบียน', 'ค่าจด', 'ค่าดำเนินการ', 'คัดป้าย', 'ป้ายขาว', 'รูดบัตร', 'ธรรมเนียม']):
        return 'REGISTRATION'
        
    if any(k in d for k in ['ฟิล์ม', 'film', 'เคลือบ', 'อุปกรณ์', 'แคมเปญ', 'มัดจำส่วนลด', 'ดาวน์', 'ผ้ายาง', 'กล้อง', 'wall box', 'wallbox', 'ม่าน', 'เบาะ', 'wrap', 'อะไหล่']):
        return 'ACCESSORIES'
        
    if any(k in d for k in ['เงินจอง', 'จองสิทธิ์', 'ลดหนี้เงินจอง', 'หัก เงินจอง', 'จองรถ', 'มัดจำจอง', 'มัดจำ', 'ลดหนี้']):
        return 'DEPOSIT'
        
    if any(k in d for k in ['สินค้ารถ', 'discount', 'รถยนต์', 'ค่างวดแรก', 'วันส่งมอบ']):
        return 'CAR_SALE_CASH'
        
    return 'DEPOSIT'

def load_sales_vat_report(file_or_path):
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    
    records = []
    target_prefixes = ('01D', '02D', '01DC', '02DC', '01WG', '02WG')
    
    for r in range(2, ws.max_row + 1):
        seq = ws.cell(row=r, column=1).value
        inv = ws.cell(row=r, column=2).value
        
        if inv is None or str(seq).strip() in ['รวม', 'Total'] or str(inv).strip() in ['รวม', 'Total']:
            continue
            
        inv_str = str(inv).strip()
        
        if not inv_str.startswith(target_prefixes):
            continue
            
        if inv_str.startswith(('HA', 'CA')):
            continue
            
        dt_val = ws.cell(row=r, column=3).value
        cust_name = str(ws.cell(row=r, column=4).value or '').strip()
        tax_id = str(ws.cell(row=r, column=5).value or '').strip()
        
        net_val = ws.cell(row=r, column=7).value
        vat_val = ws.cell(row=r, column=8).value
        gross_val = ws.cell(row=r, column=9).value
        
        try:
            f_net = float(net_val) if net_val is not None else 0.0
            f_vat = float(vat_val) if vat_val is not None else 0.0
            f_gross = float(gross_val) if gross_val is not None else 0.0
        except (ValueError, TypeError):
            continue
            
        if f_net == 0.0 and f_vat == 0.0 and f_gross == 0.0:
            continue
            
        parsed_date = None
        if dt_val:
            if isinstance(dt_val, (datetime.date, datetime.datetime)):
                parsed_date = dt_val
            else:
                parts = str(dt_val).strip().split('/')
                if len(parts) == 3:
                    try:
                        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                        year = 2000 + y if y < 100 else y
                        parsed_date = datetime.datetime(year, m, d)
                    except:
                        pass
        if not parsed_date:
            parsed_date = datetime.datetime.now()
            
        records.append({
            "seq": seq,
            "invoice_number": inv_str,
            "doc_date": parsed_date,
            "doc_date_str": parsed_date.strftime("%d/%m/%y"),
            "customer_name": cust_name,
            "tax_id": tax_id,
            "net_amount": f_net,
            "vat_amount": f_vat,
            "gross_amount": f_gross,
            "is_credit_note": (f_net < 0 or 'DC' in inv_str or 'WGCN' in inv_str)
        })
        
    return pd.DataFrame(records)

def load_gl_descriptions(file_or_path):
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    
    gl_map = {}
    for r in range(3, ws.max_row + 1):
        doc = ws.cell(row=r, column=10).value
        if not doc:
            continue
        doc_str = str(doc).strip()
        desc_str = str(ws.cell(row=r, column=12).value or '').strip()
        cr = ws.cell(row=r, column=9).value
        dr = ws.cell(row=r, column=8).value
        
        f_cr = float(cr) if cr is not None and isinstance(cr, (int, float)) else None
        f_dr = float(dr) if dr is not None and isinstance(dr, (int, float)) else None
        
        if doc_str not in gl_map:
            gl_map[doc_str] = []
        gl_map[doc_str].append({
            "desc": desc_str,
            "cr": f_cr,
            "dr": f_dr,
            "row": r
        })
        
    return gl_map

def transform_sales_to_autoline(df_vat, gl_dict, user_config=None):
    import copy
    cfg = copy.deepcopy(DEFAULT_SALES_CONFIG)
    if user_config:
        for k, v in user_config.items():
            if k == "categories" and isinstance(v, dict):
                for cat_name, cat_overrides in v.items():
                    if cat_name in cfg["categories"]:
                        cfg["categories"][cat_name].update(cat_overrides)
            else:
                cfg[k] = v
                
    rows_to_write = []
    preview_rows = []
    
    batch_idx = 1
    total_debit_sum = 0.0
    total_credit_sum = 0.0
    total_tax_sum = 0.0
    
    cat_counts = {k: 0 for k in cfg["categories"].keys()}
    cat_values = {k: 0.0 for k in cfg["categories"].keys()}
    unmatched_gl_count = 0
    
    for _, row in df_vat.iterrows():
        inv_no = row["invoice_number"]
        doc_date = row["doc_date"]
        doc_date_str = row["doc_date_str"]
        cust_name = row["customer_name"]
        
        net_val = float(row["net_amount"])
        vat_val = float(row["vat_amount"])
        gross_val = float(row["gross_amount"])
        is_cn = bool(row["is_credit_note"])
        
        abs_tax = abs(vat_val)
        abs_gross = abs(gross_val)
        abs_net = abs(net_val)
        
        branch = resolve_sales_branch(inv_no)
        narrative = f"{inv_no}_{cust_name}" if cust_name else inv_no
        if len(narrative) > 75:
            narrative = narrative[:75]
            
        doc_code = "ARC" if is_cn else "ARI"
        doc_seq = "SCREDITV" if is_cn else "SINVOICV"
        
        gl_lines = gl_dict.get(inv_no, [])
        if not gl_lines and (inv_no.startswith(('01D', '02D', '01DC', '02DC'))):
            unmatched_gl_count += 1
            
        is_wg = inv_no.startswith(('01WG', '02WG'))
        itemized_lines = []
        
        if is_wg:
            if is_finance_customer(cust_name):
                cat_key = "CAR_SALE_CREDIT"
            else:
                cat_key = "CAR_SALE_CASH"
            cat_cfg = cfg["categories"][cat_key]
            subaccount = cat_cfg["subaccount"]
            
            itemized_lines.append({
                "category": cat_key,
                "gl": cat_cfg["rev_gl"],
                "dept": cat_cfg["rev_dept"],
                "net_amount": abs_net,
                "line_narrative": narrative
            })
            
        else:
            if len(gl_lines) > 1:
                total_gl_tax = sum(abs(l["cr"] or l["dr"] or 0.0) for l in gl_lines)
                if total_gl_tax > 0:
                    allocated_net = 0.0
                    for idx_l, gl_l in enumerate(gl_lines):
                        l_desc = gl_l["desc"]
                        l_tax = abs(gl_l["cr"] or gl_l["dr"] or 0.0)
                        l_cat = classify_description_text(l_desc)
                        l_cat_cfg = cfg["categories"].get(l_cat, cfg["categories"]["DEPOSIT"])
                        
                        if idx_l == len(gl_lines) - 1:
                            l_net = round(abs_net - allocated_net, 2)
                        else:
                            l_net = round(abs_net * (l_tax / total_gl_tax), 2)
                            allocated_net += l_net
                            
                        itemized_lines.append({
                            "category": l_cat,
                            "gl": l_cat_cfg["rev_gl"],
                            "dept": l_cat_cfg["rev_dept"],
                            "net_amount": l_net,
                            "line_narrative": f"{inv_no}_{l_desc}"[:75]
                        })
                else:
                    primary_cat = classify_description_text(gl_lines[0]["desc"])
                    cat_cfg = cfg["categories"].get(primary_cat, cfg["categories"]["DEPOSIT"])
                    itemized_lines.append({
                        "category": primary_cat,
                        "gl": cat_cfg["rev_gl"],
                        "dept": cat_cfg["rev_dept"],
                        "net_amount": abs_net,
                        "line_narrative": narrative
                    })
            else:
                primary_desc = gl_lines[0]["desc"] if gl_lines else ""
                cat_key = classify_description_text(primary_desc)
                cat_cfg = cfg["categories"].get(cat_key, cfg["categories"]["DEPOSIT"])
                itemized_lines.append({
                    "category": cat_key,
                    "gl": cat_cfg["rev_gl"],
                    "dept": cat_cfg["rev_dept"],
                    "net_amount": abs_net,
                    "line_narrative": narrative
                })
                
            primary_cat_key = itemized_lines[0]["category"]
            cat_cfg = cfg["categories"].get(primary_cat_key, cfg["categories"]["DEPOSIT"])
            subaccount = cat_cfg["subaccount"]
            
        for itm in itemized_lines:
            c_key = itm["category"]
            if c_key in cat_counts:
                cat_counts[c_key] += 1
                cat_values[c_key] += itm["net_amount"]
                
        main_cat_key = itemized_lines[0]["category"]
        main_cat_cfg = cfg["categories"].get(main_cat_key, cfg["categories"]["DEPOSIT"])
        ar_gl = main_cat_cfg["ar_gl"]
        ar_dept = main_cat_cfg["ar_dept"]
        
        # Header Row
        header_record = {
            "type": "HEADER",
            "A": inv_no,
            "B": inv_no,
            "C": 1,
            "D": batch_idx,
            "E": None,
            "F": doc_code,
            "G": cfg["currency"],
            "H": doc_seq,
            "I": abs_tax,
            "J": abs_gross,
            "K": doc_date,
            "L": None,
            "M": inv_no,
            "N": narrative,
            "O": doc_date_str,
            "P": subaccount,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "U": "U",
            "V": None,
            "W": int(cfg.get("terms", 30)),
            "X": branch,
            "Y": None
        }
        rows_to_write.append(header_record)
        
        # Line 1: AR / Bank Line
        line1_debit = abs_gross if not is_cn else None
        line1_credit = abs_gross if is_cn else None
        
        line1_record = {
            "type": "DETAIL",
            "D": batch_idx,
            "G": cfg["currency"],
            "H": 1,
            "I": branch,
            "J": ar_dept,
            "K": ar_gl,
            "L": None,
            "M": line1_debit,
            "N": line1_credit,
            "O": narrative
        }
        rows_to_write.append(line1_record)
        
        # Line 2+: Revenue / Liability Lines
        line_num = 2
        for itm in itemized_lines:
            itm_net = itm["net_amount"]
            itm_debit = itm_net if is_cn else None
            itm_credit = itm_net if not is_cn else None
            
            line_record = {
                "type": "DETAIL",
                "D": batch_idx,
                "G": cfg["currency"],
                "H": line_num,
                "I": branch,
                "J": itm["dept"],
                "K": itm["gl"],
                "L": None,
                "M": itm_debit,
                "N": itm_credit,
                "O": itm["line_narrative"],
                "Q": cfg["tax_code"]
            }
            rows_to_write.append(line_record)
            line_num += 1
            
        rows_to_write.append({"type": "BLANK"})
        
        total_tax_sum += abs_tax
        if not is_cn:
            total_debit_sum += abs_gross
            total_credit_sum += abs_net
        else:
            total_credit_sum += abs_gross
            total_debit_sum += abs_net
            
        preview_rows.append({
            "Batch": batch_idx,
            "Invoice": inv_no,
            "Type": doc_code,
            "Date": doc_date_str,
            "Branch": branch,
            "Category": main_cat_cfg["name"],
            "Customer": cust_name,
            "Subaccount": subaccount,
            "Net": abs_net if not is_cn else -abs_net,
            "VAT": abs_tax if not is_cn else -abs_tax,
            "Gross": abs_gross if not is_cn else -abs_gross,
            "Items": len(itemized_lines)
        })
        
        batch_idx += 1
        
    summary_stats = {
        "total_invoices": len(df_vat),
        "total_batches": batch_idx - 1,
        "total_debit": total_debit_sum,
        "total_credit": total_credit_sum,
        "total_tax": total_tax_sum,
        "unmatched_gl_count": unmatched_gl_count,
        "category_counts": {cfg["categories"][k]["name"]: cat_counts[k] for k in cat_counts},
        "category_values": {cfg["categories"][k]["name"]: cat_values[k] for k in cat_values}
    }
    
    return rows_to_write, summary_stats, pd.DataFrame(preview_rows)

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
st.markdown('<div class="sub-header">ระบบแปลงข้อมูลบัญชีและภาษีเข้าสู่แบบฟอร์ม Autoline AR/AP Journal Import (รองรับทั้งฝ่ายบริการ Aftersales และ ฝ่ายขาย Vehicle Sales)</div>', unsafe_allow_html=True)

tab_aftersales, tab_sales = st.tabs([
    "🔧 ฝ่ายบริการหลังการขาย (Aftersales: 01S / 02S)",
    "🚗 ฝ่ายขายรถยนต์ (Vehicle Sales: 0XD / 0XWG)"
])

# =============================================================================
# TAB 1: AFTERSALES INTERFACE
# =============================================================================
with tab_aftersales:
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.subheader("1. อัปโหลดไฟล์ HEADER")
        st.caption("ไฟล์ที่มีเลขที่บิล, วันที่, ข้อมูลลูกค้า, และยอดรวม (.xlsx หรือ .csv)")
        header_file = st.file_uploader("เลือกไฟล์ HEADER (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="header_upload")
        
    with col_up2:
        st.subheader("2. อัปโหลดไฟล์ DETAIL")
        st.caption("ไฟล์ที่มีรายการแยกตามหมวด อะไหล่ P, ค่าแรง L, บริการ S (.xlsx หรือ .csv)")
        detail_file = st.file_uploader("เลือกไฟล์ DETAIL (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="detail_upload")

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
            kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
            
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
            kpi5.metric(
                "ต้นทุนอะไหล่ (COGS)",
                f"{summary_stats.get('total_parts_cogs', 0.0):,.2f} ฿",
                delta="Dr 51211006 / Cr 11511112"
            )
            
            if summary_stats["is_balanced"]:
                kpi6.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta=f"ดุลครบทุกบิล ({summary_stats['total_documents']}/{summary_stats['total_documents']})")
            else:
                kpi6.metric("สถานะดุลบัญชี", "พบเอกสารไม่ดุล ⚠️", delta=f"{summary_stats['unbalanced_count']} ฉบับ")

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
            search_kw = st.text_input("ค้นหาเอกสาร (เลขที่บิล, ใบลดหนี้, เลขที่อ้างอิง, ชื่อลูกค้า หรือ Doc Code)", placeholder="เช่น 01SC26050001, ARC, HA0027, SINVOICV", key="af_search")
            
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
        st.info("💡 กรุณาอัปโหลดไฟล์ HEADER และ DETAIL ด้านบนเพื่อเริ่มต้นแปลงข้อมูลงานบริการหลังการขาย")

# =============================================================================
# TAB 2: VEHICLE SALES INTERFACE (0XD / 0XDC / 0XWG / 0XWGCN)
# =============================================================================
with tab_sales:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.subheader("1. อัปโหลดรายงานภาษีขาย (Sales VAT Report)")
        st.caption("ไฟล์รายงานภาษีขายประจำงวด (เช่น 2026VatReport.xlsx)")
        vat_file = st.file_uploader("เลือกไฟล์รายงานภาษีขาย (.xlsx / .xls)", type=["xlsx", "xls"], key="sales_vat_upload")
        
    with col_s2:
        st.subheader("2. อัปโหลดรายงานบัญชีแยกประเภท (GL Report)")
        st.caption("ไฟล์บัญชีแยกประเภทสำหรับแยกหมวดเงินจอง/ป้ายแดง/ค่าจด/อุปกรณ์/คอม (เช่น บัญชีแยกประเภท2026.xlsx)")
        gl_file = st.file_uploader("เลือกไฟล์บัญชีแยกประเภท (.xlsx / .xls)", type=["xlsx", "xls"], key="sales_gl_upload")
        
    # Optional Config Expander
    with st.expander("⚙️ ตั้งค่า Fixed Values สำหรับฝ่ายขาย (Sales Settings)", expanded=False):
        cfg_col1, cfg_col2, cfg_col3 = st.columns(3)
        with cfg_col1:
            sales_subacc_default = st.text_input("Subaccount ลูกค้าทั่วไป (Default)", value="X0003")
        with cfg_col2:
            sales_subacc_finance = st.text_input("Subaccount สถาบันการเงิน (Finance)", value="ARCODE FINANCE")
        with cfg_col3:
            sales_dept = st.text_input("รหัสแผนกฝ่ายขาย (Sales Department)", value="2002")
            
    sales_config_override = {
        "subaccount_default": sales_subacc_default,
        "subaccount_finance": sales_subacc_finance
    }
    # Update category configs if changed
    sales_config_override["categories"] = {
        "DEPOSIT": {"subaccount": sales_subacc_default},
        "RED_PLATE": {"subaccount": sales_subacc_default},
        "REGISTRATION": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "ACCESSORIES": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "COMM_FINANCE": {"subaccount": sales_subacc_finance, "rev_dept": sales_dept},
        "CAR_SALE_CASH": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "CAR_SALE_CREDIT": {"subaccount": sales_subacc_finance, "rev_dept": sales_dept}
    }
    
    if vat_file and gl_file:
        try:
            with st.spinner("กำลังประมวลผลข้อมูลฝ่ายขายและจัดหมวดหมู่ทางบัญชี..."):
                df_vat_sales = load_sales_vat_report(vat_file)
                gl_dict_sales = load_gl_descriptions(gl_file)
                rows_sales, stats_sales, preview_sales = transform_sales_to_autoline(
                    df_vat_sales, gl_dict_sales, user_config=sales_config_override
                )
                
            st.success("✅ แปลงข้อมูลฝ่ายขายสำเร็จเรียบร้อย!")
            
            # Sales KPI Dashboard
            st.markdown("### 📊 สรุปข้อมูลธุรกรรมฝ่ายขาย (Vehicle Sales Summary)")
            skpi1, skpi2, skpi3, skpi4 = st.columns(4)
            skpi1.metric("จำนวนบิลที่แปลง (Invoices)", f"{stats_sales['total_invoices']:,} ฉบับ")
            skpi2.metric("จำนวนรายการบัญชี (Lines)", f"{len(rows_sales):,} บรรทัด")
            skpi3.metric("ยอดรวมเดบิต (Total Debit)", f"{stats_sales['total_debit']:,.2f} ฿")
            skpi4.metric("ยอดรวมภาษี (Total Tax)", f"{stats_sales['total_tax']:,.2f} ฿")
            
            # Breakdown by Category
            st.markdown("### 🚗 จำแนกยอดตามหมวดหมู่ธุรกรรม")
            cat_df = pd.DataFrame([
                {
                    "หมวดหมู่ธุรกรรม (Category)": k,
                    "จำนวนรายการ (Items)": stats_sales["category_counts"].get(k, 0),
                    "มูลค่าก่อนภาษี (Net Amount ฿)": f"{stats_sales['category_values'].get(k, 0.0):,.2f}"
                }
                for k in stats_sales["category_counts"]
            ])
            st.dataframe(cat_df, use_container_width=True, hide_index=True)
            
            if stats_sales.get("unmatched_gl_count", 0) > 0:
                st.warning(f"⚠️ พบเอกสารกลุ่ม D ที่ไม่พบในบัญชีแยกประเภทจำนวน {stats_sales['unmatched_gl_count']} ฉบับ (ระบบใช้ค่าเริ่มต้นเป็นเงินจอง)")

            # Download Button for Sales
            st.markdown("---")
            col_sdl1, col_sdl2 = st.columns([2, 1])
            with col_sdl1:
                st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline (ฝ่ายขาย)")
                st.caption("ไฟล์ Excel พร้อมสำหรับการ Import เข้าสู่ระบบ Autoline")
            with col_sdl2:
                excel_sales_output = generate_output_excel(template_path, rows_sales)
                now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="⬇️ ดาวน์โหลด Autoline_Import_SALES.xlsx",
                    data=excel_sales_output,
                    file_name=f"Autoline_Import_SALES_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            # Table Preview for Sales
            st.markdown("### 🔍 ตรวจสอบรายการเอกสาร (Sales Data Preview)")
            search_sales_kw = st.text_input("ค้นหาเอกสารฝ่ายขาย (เลขที่บิล, ชื่อลูกค้า, หรือหมวดหมู่)", placeholder="เช่น 01WG26020001, 01D26010001, เงินจอง, ธนาคาร, X0003", key="sales_search")
            
            display_sales_df = preview_sales
            if search_sales_kw.strip():
                skw = search_sales_kw.strip().lower()
                smask = (
                    preview_sales["Invoice"].astype(str).str.lower().str.contains(skw, na=False) |
                    preview_sales["Customer"].astype(str).str.lower().str.contains(skw, na=False) |
                    preview_sales["Category"].astype(str).str.lower().str.contains(skw, na=False) |
                    preview_sales["Subaccount"].astype(str).str.lower().str.contains(skw, na=False)
                )
                display_sales_df = preview_sales[smask]
                
            st.dataframe(display_sales_df, use_container_width=True, height=450)
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลฝ่ายขาย: {str(e)}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ รายงานภาษีขาย และ รายงานบัญชีแยกประเภท ด้านบนเพื่อเริ่มต้นแปลงข้อมูลฝ่ายขาย")
