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

# -----------------------------------------------------------------------------
# UNIFIED CUSTOMER MASTER & ARCODE MAPPINGS (SALE & AFTERSALE)
# DealerPro Code = Autoline Code = Customer Name (3-Way Mapping)
# -----------------------------------------------------------------------------
UNIFIED_ARCODE_BY_DEALERPRO = {
    "1005": "A0005",  # ธนาคารไทยพาณิชย์ จำกัด (มหาชน)
    "1011": "A0011",  # บริษัท ลีสซิ่งกสิกรไทย จำกัด
    "1104": "A0107",  # ธนาคารทหารไทยธนชาต จำกัด (มหาชน)
    "1163": "A0168",  # บริษัท ไทยโอริกซ์ลีสซิ่ง จำกัด (สำนักงานใหญ่)
    "1333": "A0339",  # บริษัท ภัทรลิสซิ่ง จำกัด (มหาชน)
    "1391": "A0397",  # บริษัท เอ็มยูเอฟจี เอฟแอนด์แอล (ประเทศไทย) จำกัด
    "1152": "A0465",  # บริษัท อัลฟ่า เอกซ์ จำกัด (สำนักงานใหญ่)
    "1485": "K0001",  # ธนาคารกรุงศรีอยุธยา จำกัด (มหาชน)
    "1102": "T0006",  # ธนาคารทิสโก้ จำกัด(มหาชน)
    "3508": "K0009",  # บริษัท กรุงเทพแกรนด์แปซิฟิคลีส จำกัด (มหาชน) (สำนักงานใหญ่)
    "2449": "B0001",  # บริษัท บีเอ็มดับเบิลยู ลิสซึ่ง (ประเทศไทย) จำกัด ( สำนักงานใหญ่ )
    "2470": "X0005",  # บริษัท เอ็กซ์ โมบิลิตี้ พลัส จำกัด
    "1150": "A0155",  # บริษัท ทิพยประกันภัย จำกัด (มหาชน)
    "1193": "A0198",  # บริษัท กรุงเทพประกันภัย จำกัด (มหาชน)
    "1213": "A0218",  # บมจ.เอ็ม เอส ไอ จี ประกันภัย (ประเทศไทย)
    "1223": "A0228",  # บริษัท แอกซ่าประกันภัย จำกัด (มหาชน)
    "2835": "C0007",  # บริษัท ชับบ์สามัคคีประกันภัย จำกัด (มหาชน)
    "1456": "F0001",  # บริษัท ฟอลคอนประกันภัย จำกัด (มหาชน)
    "1491": "K0007",  # บริษัท กรุงไทยพานิชประกันภัย จำกัด(มหาชน)
    "1493": "L0001",  # บริษัท แอลเอ็มจี ประกันภัย จำกัด (มหาชน)
    "1509": "M0019",  # บริษัท เมืองไทยประกันภัย จำกัด (มหาชน)
    "1625": "M0012",  # บริษัท มาสเตอร์ มอเตอร์ เซอร์วิสเซส (ประเทศไทย) จำกัด
    "2419": "A0025",  # บริษัท มาสเตอร์ คาร์เร้นเทิล จำกัด
    "1527": "X0005",  # บริษัท เอ็กซ์ โมบิลิตี้ พลัส จำกัด
    "1526": "X0006",  # บริษัท เอ็กซ์ โมบิลิตี้ (ประเทศไทย) จำกัด
    "2907": "M0012",  # MMS BOSCH สาขา ภูเก็ต
}

UNIFIED_CUSTOMER_ENTRIES = [
    ("A0005", "ธนาคารไทยพาณิชย์ จำกัด (มหาชน)", ["ไทยพาณิชย์", "scb"]),
    ("A0011", "บริษัท ลีสซิ่งกสิกรไทย จำกัด", ["ลีสซิ่งกสิกรไทย", "กสิกรไทย", "k-leasing", "k leasing"]),
    ("A0107", "ธนาคารทหารไทยธนชาต จำกัด (มหาชน)", ["ทหารไทยธนชาต", "ttb", "tmb"]),
    ("A0168", "บริษัท ไทยโอริกซ์ลีสซิ่ง จำกัด (สำนักงานใหญ่)", ["ไทยโอริกซ์", "orix"]),
    ("A0339", "บริษัท ภัทรลิสซิ่ง จำกัด (มหาชน)", ["ภัทรลิสซิ่ง", "ภัทร ลิสซิ่ง", "phatra leasing", "phatra"]),
    ("A0397", "บริษัท เอ็มยูเอฟจี เอฟแอนด์แอล (ประเทศไทย) จำกัด", ["เอ็มยูเอฟจี", "mufg"]),
    ("A0465", "บริษัท อัลฟ่า เอกซ์ จำกัด (สำนักงานใหญ่)", ["อัลฟ่า เอกซ์", "อัลฟ่าเอกซ์", "alpha x", "alpha-x"]),
    ("K0001", "ธนาคารกรุงศรีอยุธยา จำกัด (มหาชน)", ["กรุงศรีอยุธยา", "กรุงศรี", "krungsri", "bay"]),
    ("T0006", "ธนาคารทิสโก้ จำกัด(มหาชน)", ["ทิสโก้", "tisco"]),
    ("K0009", "บริษัท กรุงเทพแกรนด์แปซิฟิคลีส จำกัด (มหาชน) (สำนักงานใหญ่)", ["กรุงเทพแกรนด์แปซิฟิค", "แกรนด์แปซิฟิค", "bgpl"]),
    ("B0001", "บริษัท บีเอ็มดับเบิลยู ลิสซึ่ง (ประเทศไทย) จำกัด ( สำนักงานใหญ่ )", ["บีเอ็มดับเบิลยู", "bmw"]),
    ("X0005", "บริษัท เอ็กซ์ โมบิลิตี้ พลัส จำกัด", ["เอ็กซ์ โมบิลิตี้ พลัส", "เอ็กซ์โมบิลิตี้ พลัส", "x mobility plus"]),
    ("A0155", "บริษัท ทิพยประกันภัย จำกัด (มหาชน)", ["ทิพยประกันภัย", "ทิพย ประกันภัย", "dhipaya"]),
    ("A0198", "บริษัท กรุงเทพประกันภัย จำกัด (มหาชน)", ["กรุงเทพประกันภัย", "กรุงเทพ ประกันภัย", "bangkok insurance"]),
    ("A0218", "บมจ.เอ็ม เอส ไอ จี ประกันภัย (ประเทศไทย)", ["เอ็ม เอส ไอ จี", "เอ็มเอสไอจี", "msig"]),
    ("A0228", "บริษัท แอกซ่าประกันภัย จำกัด (มหาชน)", ["แอกซ่าประกันภัย", "แอกซ่า ประกันภัย", "axa"]),
    ("C0007", "บริษัท ชับบ์สามัคคีประกันภัย จำกัด (มหาชน)", ["ชับบ์สามัคคีประกันภัย", "ชับบ์สามัคคี", "ชับบ์", "chubb"]),
    ("F0001", "บริษัท ฟอลคอนประกันภัย จำกัด (มหาชน)", ["ฟอลคอนประกันภัย", "ฟอลคอน ประกันภัย", "falcon"]),
    ("K0007", "บริษัท กรุงไทยพานิชประกันภัย จำกัด(มหาชน)", ["กรุงไทยพานิชประกันภัย", "กรุงไทยพานิช", "กรุงไทยพาณิช", "kpi"]),
    ("L0001", "บริษัท แอลเอ็มจี ประกันภัย จำกัด (มหาชน)", ["แอลเอ็มจี ประกันภัย", "แอลเอ็มจีประกันภัย", "lmg"]),
    ("M0019", "บริษัท เมืองไทยประกันภัย จำกัด (มหาชน)", ["เมืองไทยประกันภัย", "เมืองไทย ประกันภัย", "muang thai insurance"]),
    ("M0012", "บริษัท มาสเตอร์ มอเตอร์ เซอร์วิสเซส (ประเทศไทย) จำกัด", ["มาสเตอร์ มอเตอร์ เซอร์วิสเซส", "มาสเตอร์ มอเตอร์", "มาสเตอร์ คาร์เร้นเทิล เซอร์วิส", "mms"]),
    ("A0025", "บริษัท มาสเตอร์ คาร์เร้นเทิล จำกัด", ["มาสเตอร์ คาร์เร้นเทิล", "master car rental"]),
    ("X0006", "บริษัท เอ็กซ์ โมบิลิตี้ (ประเทศไทย) จำกัด", ["เอ็กซ์ โมบิลิตี้", "เอ็กซ์โมบิลิตี้", "x mobility"]),
]

def resolve_unified_arcode(dealerpro_code=None, customer_name=None, context="aftersale", default_override=None):
    """
    จับคู่ ARCODE 3 ทาง (DealerPro Code = Autoline Code = Customer Name):
    1. ตรวจสอบรหัสจาก DealerPro (ถ้ามีระบุ เช่น customer_code ใน Aftersales หรือ fin_coy ใน Sales)
    2. ตรวจสอบจากชื่อลูกค้า (Customer Name)
    3. ค่าเริ่มต้นตามบริบท (Context):
       - ฝั่ง Sale: ค่าเริ่มต้นเป็น 'X0003' (Vehicle Xpeng)
       - ฝั่ง Aftersale: ค่าเริ่มต้นเป็น 'X0004' (Service Xpeng)
    """
    # 1. Match from DealerPro code
    if dealerpro_code is not None:
        d_code = str(dealerpro_code).strip().split('.')[0]
        if d_code in UNIFIED_ARCODE_BY_DEALERPRO:
            return UNIFIED_ARCODE_BY_DEALERPRO[d_code]
            
    # 2. Match from Customer Name
    c = str(customer_name or "").strip().lower()
    if c and c != "nan" and c != "none":
        if "มัดจำอะไหล่" in c or ("มัดจำ" in c and "อะไหล่" in c) or "parts deposit" in c:
            return "D0001"
            
        if any(k in c for k in ["ซีเจ", "ทีจี", "ซูมิ", "มิซูโฮ"]):
            return "X0003"
            
        # Specific disambiguation rules:
        # X Mobility Plus (X0005) vs X Mobility Thailand (X0006)
        if ("เอ็กซ์" in c or "x mobility" in c) and ("พลัส" in c or "plus" in c):
            return "X0005"
        if "เอ็กซ์ โมบิลิตี้" in c or "เอ็กซ์โมบิลิตี้" in c or "x mobility" in c:
            return "X0006"
            
        # Master Motor Services (M0012) vs Master Car Rental (A0025)
        if "มาสเตอร์" in c:
            if "เซอร์วิส" in c or "service" in c or "motor" in c or "mms" in c:
                return "M0012"
            if "คาร์เร้นเทิล" in c or "rental" in c:
                return "A0025"
                
        # Non-life Insurance vs Life Insurance
        if "ทิพย" in c and "ประกัน" in c:
            return "A0155"
        if "กรุงเทพ" in c and "ประกัน" in c and "ชีวิต" not in c:
            return "A0198"
        if "เมืองไทย" in c and "ประกัน" in c and "ชีวิต" not in c:
            return "M0019"
            
        for autoline_code, off_name, keywords in UNIFIED_CUSTOMER_ENTRIES:
            clean_off = off_name.replace("บริษัท", "").replace("จำกัด", "").replace("จํากัด", "").replace("(มหาชน)", "").replace("(สำนักงานใหญ่)", "").replace("( สำนักงานใหญ่ )", "").replace("ธนาคาร", "").strip().lower()
            if clean_off and clean_off in c:
                return autoline_code
            for kw in keywords:
                if kw.lower() in c:
                    return autoline_code
                    
    # 3. Fallback default by context
    if default_override is not None:
        return default_override
    if context == "sale":
        return "X0003"
    return "X0004"

def resolve_aftersales_subaccount(customer_name, dealerpro_code=None, default="X0004"):
    return resolve_unified_arcode(dealerpro_code=dealerpro_code, customer_name=customer_name, context="aftersale", default_override=default)

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
        
    # กรองแถวว่าง / แถวสรุปผลรวมท้ายไฟล์ที่ไม่มีเลขที่ใบแจ้งหนี้
    df_header = df_header[df_header["invoice_number"].notna()].copy()
    df_detail = df_detail[df_detail["invoice_number"].notna()].copy()
    
    df_header["invoice_number"] = df_header["invoice_number"].astype(str).str.strip()
    df_detail["invoice_number"] = df_detail["invoice_number"].astype(str).str.strip()
    
    df_header = df_header[~df_header["invoice_number"].str.lower().isin(["", "nan", "none"])].copy()
    df_detail = df_detail[~df_detail["invoice_number"].str.lower().isin(["", "nan", "none"])].copy()
    
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
        if not inv_no or inv_no.lower() in ("nan", "none") or inv_no.upper().startswith(("CA", "HA")):
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
        
        cust_code = str(h_row.get("customer_code", "")).strip()
        subaccount = resolve_aftersales_subaccount(customer_name=cust_name, dealerpro_code=cust_code, default=AUTOMATED_CONFIG.get("subaccount_default", "X0004"))
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
# 3.2. PARTS OVER-THE-COUNTER SALES TRANSFORMER (01P / 02P / 01PC / 02PC)
# =============================================================================
def parse_parts_sales_file(file_obj_or_path):
    """
    อ่านและแปลงโครงสร้างรายงานการขายอะไหล่ (เช่น AftersalePart2026.xlsx)
    สกัดเฉพาะบิลขายอะไหล่หน้าร้าน (01P, 02P) และใบลดหนี้ (01PC, 02PC)
    """
    import re
    wb = openpyxl.load_workbook(file_obj_or_path, data_only=True)
    ws = wb.active

    current_cust_code = ""
    current_cust_name = ""
    current_branch_code = "0001"
    invoices = {}

    for r in range(1, ws.max_row + 1):
        c1 = str(ws.cell(row=r, column=1).value or "").strip()
        c2 = str(ws.cell(row=r, column=2).value or "").strip()
        
        if c1.startswith("สาขา:"):
            br = c1.replace("สาขา:", "").strip()
            if br == "01": current_branch_code = "0001"
            elif br == "02": current_branch_code = "0002"
            else: current_branch_code = br.zfill(4)
            continue
            
        if c1.startswith("ลูกค้า:"):
            current_cust_code = c1.replace("ลูกค้า:", "").strip()
            current_cust_name = c2
            continue
            
        if c1.startswith("รวม") or c1.lower().startswith("total") or c1.startswith("<====") or c1.startswith("บริษัท:") or c1.startswith("ฝ่าย:"):
            continue
            
        if c1.startswith("เอกสารเลขที่:"):
            continue

        date_raw = ws.cell(row=r, column=3).value
        # Match invoice patterns like 01P..., 02P..., 01PC..., 02PC...
        if c1 and (re.search(r'^\d{2}P', c1, re.IGNORECASE) or re.search(r'^\d{2}PC', c1, re.IGNORECASE)) and date_raw:
            inv_no = c1
            cust_code = str(ws.cell(row=r, column=2).value or current_cust_code).strip()
            part_desc = str(ws.cell(row=r, column=4).value or "").strip()
            qty = float(ws.cell(row=r, column=6).value or 0.0)
            retail_price = float(ws.cell(row=r, column=7).value or 0.0)
            net_sales = float(ws.cell(row=r, column=8).value or 0.0)
            cost = float(ws.cell(row=r, column=9).value or 0.0)
            vat = float(ws.cell(row=r, column=10).value or 0.0)
            
            # Parse Date
            if isinstance(date_raw, (datetime.datetime, pd.Timestamp)):
                doc_date = date_raw.date() if hasattr(date_raw, 'date') else date_raw
            elif isinstance(date_raw, str):
                try:
                    doc_date = datetime.datetime.strptime(date_raw.strip(), "%d/%m/%y").date()
                except:
                    try:
                        doc_date = datetime.datetime.strptime(date_raw.strip(), "%d/%m/%Y").date()
                    except:
                        doc_date = datetime.date.today()
            else:
                doc_date = datetime.date.today()
                
            branch = current_branch_code
            if inv_no.startswith("01"): branch = "0001"
            elif inv_no.startswith("02"): branch = "0002"
            
            if inv_no not in invoices:
                invoices[inv_no] = {
                    "inv_no": inv_no,
                    "cust_code": cust_code,
                    "cust_name": current_cust_name,
                    "branch": branch,
                    "doc_date": doc_date,
                    "items": [],
                    "total_net": 0.0,
                    "total_vat": 0.0,
                    "total_cost": 0.0
                }
            invoices[inv_no]["items"].append({
                "part": part_desc,
                "qty": qty,
                "net_sales": net_sales,
                "cost": cost,
                "vat": vat
            })
            invoices[inv_no]["total_net"] += net_sales
            invoices[inv_no]["total_vat"] += vat
            invoices[inv_no]["total_cost"] += cost
            
    return invoices


def transform_parts_sales_to_autoline(invoices_dict, config=None):
    """
    แปลงข้อมูลขายอะไหล่หน้าร้าน (01P / 01PC) เป็นรายการ Autoline AR Journal Import
    พร้อมลงบันทึกต้นทุนและสต็อกอะไหล่ (COGS & Inventory)
    """
    if config is None:
        config = {}
        
    rows_to_write = []
    preview_records = []
    
    total_debit_sum = 0.0
    total_credit_sum = 0.0
    total_parts_cogs = 0.0
    
    inv_revenue = 0.0
    inv_tax = 0.0
    inv_ar = 0.0
    cn_revenue = 0.0
    cn_tax = 0.0
    cn_ar = 0.0
    
    count_inv = 0
    count_cn = 0
    
    batch_idx = 1
    sorted_inv_keys = sorted(invoices_dict.keys())
    
    for inv_no in sorted_inv_keys:
        inv_data = invoices_dict[inv_no]
        cust_code = inv_data.get("cust_code", "")
        cust_name = inv_data.get("cust_name", "")
        doc_date = inv_data.get("doc_date", datetime.date.today())
        branch = inv_data.get("branch", "0001")
        
        raw_net = inv_data.get("total_net", 0.0)
        raw_vat = inv_data.get("total_vat", 0.0)
        raw_cost = inv_data.get("total_cost", 0.0)
        
        is_cn = ("PC" in inv_no.upper()) or (raw_net < 0) or (raw_vat < 0)
        
        net_amt = round(abs(raw_net), 2)
        vat_amt = round(abs(raw_vat), 2)
        gross_amt = round(net_amt + vat_amt, 2)
        cost_amt = round(abs(raw_cost), 2)
        
        doc_code = "ARC" if is_cn else "ARI"
        doc_seq = "SCREDITV" if is_cn else "SINVOICV"
        has_vat = (vat_amt > 0)
        tax_group_val = "U" if has_vat else "OS"
        crcode_val = "00000004" if is_cn else None
        
        narrative = f"{inv_no}_{cust_name}" if cust_name else inv_no
        subaccount = resolve_unified_arcode(dealerpro_code=cust_code, customer_name=cust_name, context="aftersale")
        
        if is_cn:
            count_cn += 1
            cn_revenue += net_amt
            cn_tax += vat_amt
            cn_ar += gross_amt
        else:
            count_inv += 1
            inv_revenue += net_amt
            inv_tax += vat_amt
            inv_ar += gross_amt
            
        # 1. Header Row
        header_record = {
            "type": "HEADER",
            "A": inv_no,
            "B": inv_no,
            "C": 1,
            "D": batch_idx,
            "E": None,
            "F": doc_code,
            "G": "THB",
            "H": doc_seq,
            "I": vat_amt,
            "J": gross_amt,
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
            "U": tax_group_val,
            "V": None,
            "W": 30,
            "X": branch,
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
            "Date": doc_date.strftime("%Y-%m-%d") if hasattr(doc_date, "strftime") else str(doc_date),
            "GL Code": "",
            "Department": "",
            "Debit": "",
            "Credit": "",
            "Tax": vat_amt,
            "Total Value": gross_amt,
            "Narrative": narrative,
            "Branch": branch,
            "Subaccount": subaccount
        })
        
        # 2. Blank Row
        rows_to_write.append({"type": "BLANK"})
        
        # 3. Line 1: AR Control Account (11311001 / Dept 0000)
        ar_gl = 11311001
        ar_dept = "0000"
        ar_debit = None if is_cn else gross_amt
        ar_credit = gross_amt if is_cn else None
        
        ar_record = {
            "type": "ITEM",
            "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
            "G": "THB", "H": 1,
            "I": branch, "J": ar_dept, "K": ar_gl, "L": None,
            "M": ar_debit, "N": ar_credit, "O": narrative, "P": None, "Q": None,
            "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
            "X": None, "Y": None, "Z": None, "AA": None, "AB": None, "AC": None, "AD": None
        }
        rows_to_write.append(ar_record)
        if ar_debit is not None: total_debit_sum += ar_debit
        if ar_credit is not None: total_credit_sum += ar_credit
        
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": "CREDIT (AR)" if is_cn else "DEBIT (AR)",
            "Batch": batch_idx,
            "Line": 1,
            "Doc Code": "",
            "Doc Seq": "",
            "Date": doc_date.strftime("%Y-%m-%d") if hasattr(doc_date, "strftime") else str(doc_date),
            "GL Code": str(ar_gl),
            "Department": ar_dept,
            "Debit": ar_debit if ar_debit is not None else "",
            "Credit": ar_credit if ar_credit is not None else "",
            "Tax": "",
            "Total Value": "",
            "Narrative": narrative,
            "Branch": branch,
            "Subaccount": subaccount
        })
        
        # 4. Line 2: Parts Revenue (41211001 / Dept 4002)
        parts_rev_gl = 41211001
        parts_dept = "4002"
        tax_code_val = "S" if has_vat else "O"
        rev_debit = net_amt if is_cn else None
        rev_credit = None if is_cn else net_amt
        
        rev_record = {
            "type": "ITEM",
            "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
            "G": "THB", "H": 2,
            "I": branch, "J": parts_dept, "K": parts_rev_gl, "L": None,
            "M": rev_debit, "N": rev_credit, "O": narrative, "P": None, "Q": tax_code_val,
            "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
            "X": "R", "Y": None, "Z": "J", "AA": None, "AB": "A", "AC": None, "AD": None
        }
        rows_to_write.append(rev_record)
        if rev_debit is not None: total_debit_sum += rev_debit
        if rev_credit is not None: total_credit_sum += rev_credit
        
        preview_records.append({
            "Invoice": inv_no,
            "Row Type": "DEBIT (Parts Rev)" if is_cn else "CREDIT (Parts Rev)",
            "Batch": batch_idx,
            "Line": 2,
            "Doc Code": "",
            "Doc Seq": "",
            "Date": doc_date.strftime("%Y-%m-%d") if hasattr(doc_date, "strftime") else str(doc_date),
            "GL Code": str(parts_rev_gl),
            "Department": parts_dept,
            "Debit": rev_debit if rev_debit is not None else "",
            "Credit": rev_credit if rev_credit is not None else "",
            "Tax": tax_code_val,
            "Total Value": "",
            "Narrative": narrative,
            "Branch": branch,
            "Subaccount": ""
        })
        
        # 5. Line 3 & Line 4: COGS & Inventory (if cost > 0)
        line_num = 3
        if cost_amt > 0:
            total_parts_cogs += cost_amt
            cogs_gl = 51211006
            cogs_dept = "4002"
            cogs_debit = cost_amt if not is_cn else None
            cogs_credit = None if not is_cn else cost_amt
            
            cogs_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": "THB", "H": line_num,
                "I": branch, "J": cogs_dept, "K": cogs_gl, "L": None,
                "M": cogs_debit, "N": cogs_credit, "O": narrative, "P": None, "Q": "O",
                "R": None, "S": None, "T": None, "U": None, "V": None, "W": None,
                "X": "R", "Y": None, "Z": "J", "AA": None, "AB": "A", "AC": None, "AD": None
            }
            rows_to_write.append(cogs_record)
            if cogs_debit is not None: total_debit_sum += cogs_debit
            if cogs_credit is not None: total_credit_sum += cogs_credit
            
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": "CREDIT (COGS)" if is_cn else "DEBIT (COGS)",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d") if hasattr(doc_date, "strftime") else str(doc_date),
                "GL Code": str(cogs_gl),
                "Department": cogs_dept,
                "Debit": cogs_debit if cogs_debit is not None else "",
                "Credit": cogs_credit if cogs_credit is not None else "",
                "Tax": "O",
                "Total Value": "",
                "Narrative": narrative,
                "Branch": branch,
                "Subaccount": ""
            })
            line_num += 1
            
            inv_gl = 11511112
            inv_dept = "0000"
            inv_debit = cost_amt if is_cn else None
            inv_credit = None if is_cn else cost_amt
            
            inv_stock_record = {
                "type": "ITEM",
                "B": None, "C": None, "D": batch_idx, "E": None, "F": None,
                "G": "THB", "H": line_num,
                "I": branch, "J": inv_dept, "K": inv_gl, "L": None,
                "M": inv_debit, "N": inv_credit, "O": narrative, "P": None, "Q": "O",
                "R": None, "S": None, "T": None, "U": "JEEP", "V": "JEEP-G6", "W": None,
                "X": None, "Y": None, "Z": None, "AA": None, "AB": None, "AC": None, "AD": None
            }
            rows_to_write.append(inv_stock_record)
            if inv_debit is not None: total_debit_sum += inv_debit
            if inv_credit is not None: total_credit_sum += inv_credit
            
            preview_records.append({
                "Invoice": inv_no,
                "Row Type": "DEBIT (Stock)" if is_cn else "CREDIT (Stock)",
                "Batch": batch_idx,
                "Line": line_num,
                "Doc Code": "",
                "Doc Seq": "",
                "Date": doc_date.strftime("%Y-%m-%d") if hasattr(doc_date, "strftime") else str(doc_date),
                "GL Code": str(inv_gl),
                "Department": inv_dept,
                "Debit": inv_debit if inv_debit is not None else "",
                "Credit": inv_credit if inv_credit is not None else "",
                "Tax": "O",
                "Total Value": "",
                "Narrative": narrative,
                "Branch": branch,
                "Subaccount": ""
            })
            
        rows_to_write.append({"type": "BLANK"})
        batch_idx += 1
        
    if rows_to_write and rows_to_write[-1].get("type") == "BLANK":
        rows_to_write.pop()
        
    net_revenue = round(inv_revenue - cn_revenue, 2)
    net_tax = round(inv_tax - cn_tax, 2)
    net_ar = round(inv_ar - cn_ar, 2)
    
    summary_stats = {
        "total_documents": len(invoices_dict),
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
        "is_balanced": abs((total_debit_sum - total_credit_sum) - (inv_tax - cn_tax)) < 0.05,
        "difference": round(abs((total_debit_sum - total_credit_sum) - (inv_tax - cn_tax)), 2)
    }
    
    return rows_to_write, summary_stats, pd.DataFrame(preview_records)


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
    "vehicle_cogs_gl": 51111012,
    "vehicle_cogs_dept": "2002",
    "vehicle_cogs_tax": "O",
    "vehicle_cogs_manufact": "JEEP",
    "vehicle_cogs_model": "JEEP-G6",
    "vehicle_cogs_saletype": "TRADE",
    "vehicle_inv_gl": 11511111,
    "vehicle_inv_dept": "0000",
    "vehicle_inv_tax": "O",
    "vehicle_inv_manufact": "JEEP",
    "vehicle_inv_model": "JEEP-G6",
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
        "FEE": {
            "name": "รายได้ค่าธรรมเนียม",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "49291009",     # ค่าธรรมเนียม
            "rev_dept": "2002",
            "subaccount": "X0003"
        },
        "PARTS_DEPOSIT": {
            "name": "มัดจำอะไหล่",
            "ar_gl": "11311001",      # ลูกหนี้การค้า
            "ar_dept": "0000",
            "rev_gl": "21931004",     # มัดจำอะไหล่
            "rev_dept": "0000",
            "subaccount": "D0001"
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
    if len(inv_str) >= 2 and inv_str[:2].isdigit():
        return f"00{inv_str[:2]}"
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
        
    if 'มัดจำอะไหล่' in d or ('มัดจำ' in d and 'อะไหล่' in d) or 'parts deposit' in d:
        return 'PARTS_DEPOSIT'
        
    if any(k in d for k in ['ค่าธรรมเนียม', 'ธรรมเนียม', 'fee']):
        return 'FEE'
        
    if any(k in d for k in ['ป้ายแดง', 'ป้ายเเดง', 'มัดจำป้าย']):
        return 'RED_PLATE'
        
    if any(k in d for k in ['คอมไฟแนนซ์', 'ค่าคอมไฟแนนซ์', 'คอมมิชชั่นไฟแนนซ์', 'ค่าคอมมิชชั่นไฟแนนซ์', 'comm', 'finan']):
        return 'COMM_FINANCE'
        
    if any(k in d for k in ['จดทะเบียน', 'ค่าจด', 'ค่าดำเนินการ', 'คัดป้าย', 'ป้ายขาว', 'รูดบัตร']):
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
        acc = str(ws.cell(row=r, column=3).value or '').strip()
        desc_str = str(ws.cell(row=r, column=12).value or '').strip()
        cr = ws.cell(row=r, column=9).value
        dr = ws.cell(row=r, column=8).value
        
        f_cr = float(cr) if cr is not None and isinstance(cr, (int, float)) else None
        f_dr = float(dr) if dr is not None and isinstance(dr, (int, float)) else None
        
        if doc_str not in gl_map:
            gl_map[doc_str] = []
        gl_map[doc_str].append({
            "acc": acc,
            "desc": desc_str,
            "cr": f_cr,
            "dr": f_dr,
            "row": r
        })
        
    return gl_map

def clean_customer_name(name):
    if not name:
        return ""
    import re
    s = str(name).strip()
    s = re.sub(r'\s*\((สำนักงานใหญ่|สนญ\.|สาขาที่\s*\d+|สาขา\s*\d+|สนง\.\s*ใหญ่)\s*\)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+(สำนักงานใหญ่|สนญ\.|สาขาที่\s*\d+)$', '', s, flags=re.IGNORECASE)
    return s.strip()

def load_stock_numbers(file_or_path):
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    stock_by_inv = {}
    fin_coy_by_inv = {}
    vin_by_inv = {}
    for r in range(1, ws.max_row + 1):
        c1 = ws.cell(row=r, column=1).value
        c2 = ws.cell(row=r, column=2).value
        c3 = ws.cell(row=r, column=3).value
        c5 = ws.cell(row=r, column=5).value
        if c1 and str(c1).strip().startswith('N0') and c5:
            stk_no = str(c1).strip()
            inv_no = str(c5).strip()
            c3_str = str(c3).strip() if c3 is not None else ""
            c2_str = str(c2).strip() if c2 is not None else ""
            if inv_no:
                stock_by_inv[inv_no] = stk_no
                if c3_str:
                    fin_coy_by_inv[inv_no] = c3_str
                if c2_str:
                    vin_by_inv[inv_no] = c2_str
    return stock_by_inv, fin_coy_by_inv, vin_by_inv

def load_vehicle_costs(file_or_path):
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    cost_by_stock = {}
    for r in range(1, ws.max_row + 1):
        c1 = ws.cell(row=r, column=1).value
        if c1 and str(c1).strip().startswith('N0'):
            stk_no = str(c1).strip()
            cost_val = ws.cell(row=r, column=5).value
            if cost_val is not None:
                try:
                    cost_by_stock[stk_no] = float(cost_val)
                except (ValueError, TypeError):
                    pass
    return cost_by_stock

DEFAULT_FINANCE_DATA = {
    "by_code": {
        "1005": "A0005",
        "1011": "A0011",
        "1104": "A0107",
        "1163": "A0168",
        "1333": "A0339",
        "1391": "A0397",
        "1152": "A0465",
        "1485": "K0001",
        "1102": "T0006",
        "3508": "K0009",
        "2449": "B0001",
    },
    "list": [
        ("1005", "A0005", "ธนาคารไทยพาณิชย์ จำกัด (มหาชน)"),
        ("1011", "A0011", "บริษัท ลีสซิ่งกสิกรไทย จำกัด"),
        ("1104", "A0107", "ธนาคารทหารไทยธนชาต จำกัด (มหาชน)"),
        ("1163", "A0168", "บริษัท ไทยโอริกซ์ลีสซิ่ง จำกัด (สำนักงานใหญ่)"),
        ("1333", "A0339", "บริษัท ภัทรลิสซิ่ง จำกัด (มหาชน)"),
        ("1391", "A0397", "บริษัท เอ็มยูเอฟจี เอฟแอนด์แอล (ประเทศไทย) จำกัด"),
        ("1152", "A0465", "บริษัท อัลฟ่า เอกซ์ จำกัด (สำนักงานใหญ่)"),
        ("1485", "K0001", "ธนาคารกรุงศรีอยุธยา จำกัด (มหาชน)"),
        ("1102", "T0006", "ธนาคารทิสโก้ จำกัด(มหาชน)"),
        ("3508", "K0009", "บริษัท กรุงเทพแกรนด์แปซิฟิคลีส จำกัด (มหาชน) (สำนักงานใหญ่)"),
        ("2449", "B0001", "บริษัท บีเอ็มดับเบิลยู ลิสซึ่ง (ประเทศไทย) จำกัด ( สำนักงานใหญ่ )"),
        ("", "X0003", "บริษัท ซีเจ แคปปิตอล จำกัด"),
        ("", "X0003", "บริษัท ซูมิ ลิซ ลิสซิ่ง จำกัด"),
        ("", "X0003", "บริษัท ทีจี แคปปิตอล จํากัด"),
        ("", "X0003", "บริษัท กรุงไทย มิซูโฮ ลีสซิ่ง จำกัด"),
        ("", "X0003", "ซีเจ แคปปิตอล"),
        ("", "X0003", "ทีจี แคปปิตอล"),
        ("", "X0003", "ซูมิ ลิซ"),
        ("", "X0003", "มิซูโฮ ลีสซิ่ง"),
        ("", "X0003", "ซีเจ"),
        ("", "X0003", "ทีจี"),
        ("", "X0003", "ซูมิ"),
        ("", "X0003", "มิซูโฮ"),
    ]
}

def load_finance_codes(file_or_path=None):
    fin_by_code = dict(DEFAULT_FINANCE_DATA["by_code"])
    fin_list = list(DEFAULT_FINANCE_DATA["list"])
    if file_or_path:
        try:
            wb = openpyxl.load_workbook(file_or_path, data_only=True)
            ws = wb.active
            for r in range(2, ws.max_row + 1):
                d_pro = str(ws.cell(r, 1).value or '').strip()
                autoline = str(ws.cell(r, 2).value or '').strip()
                c_name = str(ws.cell(r, 3).value or '').strip()
                if autoline:
                    fin_list.append((d_pro, autoline, c_name))
                    if d_pro:
                        fin_by_code[d_pro] = autoline
        except Exception:
            pass
    return {"by_code": fin_by_code, "list": fin_list}

def resolve_finance_subaccount(inv_no, cust_name, fin_coy_code=None, fin_data=None, default_code="X0003"):
    # If custom fin_data was uploaded and differs from default, check it first
    if fin_data and fin_data != DEFAULT_FINANCE_DATA:
        by_code = fin_data.get("by_code", {})
        f_list = fin_data.get("list", [])
        if fin_coy_code and str(fin_coy_code).strip() in by_code:
            return by_code[str(fin_coy_code).strip()]
        c = str(cust_name or '').strip()
        if c:
            for d_pro, autoline, f_name in f_list:
                clean_f = f_name.replace('บริษัท', '').replace('จำกัด', '').replace('จํากัด', '').replace('(มหาชน)', '').replace('(สำนักงานใหญ่)', '').replace('ธนาคาร', '').strip()
                if clean_f and clean_f in c:
                    return autoline
            for d_pro, autoline, f_name in f_list:
                if f_name and f_name in c:
                    return autoline
                    
    # Unified 3-way resolution (DealerPro Code = Autoline Code = Customer Name)
    matched = resolve_unified_arcode(dealerpro_code=fin_coy_code, customer_name=cust_name, context="sale", default_override=None)
    if matched:
        return matched
    return default_code
            
def load_vehicle_gl_cogs(file_or_path):
    import re
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    
    # First pass: map VIN to Stock number from descriptions across the sheet
    vin_to_stock = {}
    for r in range(2, ws.max_row + 1):
        for c in (10, 11, 12):
            val = str(ws.cell(r, c).value or '').strip()
            m = re.search(r'(N0\d+)[#\-_ ]*([A-Z0-9]{6,17})', val)
            if m:
                stk, vin = m.group(1), m.group(2)
                vin_to_stock[vin] = stk
                if len(vin) >= 6:
                    vin_to_stock[vin[-6:]] = stk
                    vin_to_stock[vin[-8:]] = stk
                    
    cogs_map = {}
    for r in range(2, ws.max_row + 1):
        acc = str(ws.cell(r, 3).value or '').strip()
        tt = str(ws.cell(r, 6).value or '').strip()
        # Account 511001 with TT == '2' contains actual vehicle sales cost (both ARI and ARC)
        if (acc == '511001' or acc.startswith('511')) and tt == '2':
            doc = str(ws.cell(r, 10).value or '').strip()
            dr = ws.cell(r, 8).value
            cr = ws.cell(r, 9).value
            desc = str(ws.cell(r, 12).value or '').strip()
            sub = str(ws.cell(r, 11).value or '').strip()
            
            amt = dr if dr is not None and dr != '' else cr
            if doc and amt is not None and amt != '':
                try:
                    f_amt = abs(float(amt))
                except (ValueError, TypeError):
                    continue
                if f_amt > 0:
                    m_stk = re.search(r'(N0\d+)', desc)
                    stk_no = m_stk.group(1) if m_stk else ""
                    m_vin = re.search(r'([A-Z0-9]{8,17})', desc)
                    vin_no = m_vin.group(1) if m_vin else ""
                    if not stk_no and vin_no:
                        stk_no = vin_to_stock.get(vin_no, vin_to_stock.get(vin_no[-6:], vin_to_stock.get(vin_no[-8:], '')))
                    ident = stk_no or vin_no or ""
                    
                    if doc not in cogs_map or cogs_map[doc]["cost"] == 0:
                        cogs_map[doc] = {
                            "cost": f_amt,
                            "stock_no": stk_no,
                            "vin": vin_no,
                            "identifier": ident,
                            "desc": desc,
                            "model_code": sub
                        }
    return cogs_map

def load_vehicle_profit(file_or_path):
    wb = openpyxl.load_workbook(file_or_path, data_only=True)
    ws = wb.active
    vin_to_stk = {}
    vin_to_cost = {}
    stk_to_cost = {}
    stk_to_vin = {}
    stk_to_model = {}
    
    for r in range(1, ws.max_row + 1):
        c1 = str(ws.cell(r, 1).value or '').strip()
        c3 = str(ws.cell(r, 3).value or '').strip()
        m_name = str(ws.cell(r, 4).value or '').strip()
        m_var = str(ws.cell(r, 5).value or '').strip()
        c9 = ws.cell(r, 9).value
        
        if c1.startswith('N0'):
            if c3:
                vin_to_stk[c3] = c1
                stk_to_vin[c1] = c3
                if len(c3) >= 8:
                    vin_to_stk[c3[-8:]] = c1
            if m_name or m_var:
                stk_to_model[c1] = f"{m_name} {m_var}".strip()
            if c9 is not None:
                try:
                    f_cost = float(c9)
                    stk_to_cost[c1] = f_cost
                    if c3:
                        vin_to_cost[c3] = f_cost
                        if len(c3) >= 8:
                            vin_to_cost[c3[-8:]] = f_cost
                except (ValueError, TypeError):
                    pass
                    
    return {
        "vin_to_stk": vin_to_stk,
        "vin_to_cost": vin_to_cost,
        "stk_to_cost": stk_to_cost,
        "stk_to_vin": stk_to_vin,
        "stk_to_model": stk_to_model
    }

# =============================================================================
# TRANSFORMATION ENGINE
# =============================================================================
def transform_sales_to_autoline(df_vat, gl_dict, stock_dict=None, cost_dict=None, fin_coy_dict=None, fin_data=None, vehicle_gl_dict=None, vin_dict=None, vehicle_profit_dict=None, user_config=None):
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
                
    if gl_dict is None:
        gl_dict = {}
    if stock_dict is None:
        stock_dict = {}
    if cost_dict is None:
        cost_dict = {}
    if fin_coy_dict is None:
        fin_coy_dict = {}
    if vin_dict is None:
        vin_dict = {}
    if vehicle_profit_dict is None:
        vehicle_profit_dict = {}
        
    rows_to_write = []
    preview_rows = []
    missing_records = []
    
    from collections import defaultdict
    branch_batches = defaultdict(int)
    branch_rows = defaultdict(list)
    branch_previews = defaultdict(list)
    branch_stats_data = defaultdict(lambda: {
        "total_invoices": 0,
        "total_batches": 0,
        "total_debit": 0.0,
        "total_credit": 0.0,
        "total_tax": 0.0,
        "total_vehicle_cost": 0.0,
        "matched_cost_count": 0,
    })
    
    # Doc Type Grouping (0XD vs 0XWG)
    doc_type_batches = defaultdict(int)
    doc_type_rows = defaultdict(list)
    doc_type_previews = defaultdict(list)
    doc_type_stats_data = defaultdict(lambda: {
        "total_invoices": 0,
        "total_batches": 0,
        "total_debit": 0.0,
        "total_credit": 0.0,
        "total_tax": 0.0,
        "total_vehicle_cost": 0.0,
        "matched_cost_count": 0,
    })
    
    global_batch_idx = 1
    total_debit_sum = 0.0
    total_credit_sum = 0.0
    total_tax_sum = 0.0
    total_vehicle_cost = 0.0
    matched_cost_count = 0
    
    cat_counts = {k: 0 for k in cfg["categories"].keys()}
    cat_values = {k: 0.0 for k in cfg["categories"].keys()}
    unmatched_gl_count = 0
    
    for _, row in df_vat.iterrows():
        inv_no = row["invoice_number"]
        doc_date = row["doc_date"]
        doc_date_str = row["doc_date_str"]
        cust_name = row["customer_name"]
        clean_cust = clean_customer_name(cust_name)
        
        net_val = float(row["net_amount"])
        vat_val = float(row["vat_amount"])
        gross_val = float(row["gross_amount"])
        is_cn = bool(row["is_credit_note"])
        
        abs_tax = abs(vat_val)
        abs_gross = abs(gross_val)
        abs_net = abs(net_val)
        
        branch = resolve_sales_branch(inv_no)
        branch_batches[branch] += 1
        b_batch_idx = branch_batches[branch]
        g_batch_idx = global_batch_idx
        
        is_wg = inv_no.startswith(('01WG', '02WG', '03WG', '04WG'))
        doc_group = "0XWG" if is_wg else "0XD"
        doc_type_batches[doc_group] += 1
        dt_batch_idx = doc_type_batches[doc_group]
        
        def add_record(rec):
            if rec.get("type") == "BLANK":
                rows_to_write.append({"type": "BLANK"})
                branch_rows[branch].append({"type": "BLANK"})
                doc_type_rows[doc_group].append({"type": "BLANK"})
            else:
                rows_to_write.append(dict(rec, D=g_batch_idx))
                branch_rows[branch].append(dict(rec, D=b_batch_idx))
                doc_type_rows[doc_group].append(dict(rec, D=dt_batch_idx))
                
        doc_code = "ARC" if is_cn else "ARI"
        doc_seq = "SCREDITV" if is_cn else "SINVOICV"
        
        gl_lines = gl_dict.get(inv_no, [])
        if not gl_lines and (inv_no.startswith(('01D', '02D', '01DC', '02DC'))):
            unmatched_gl_count += 1
            
        # Fixed Defaults as required by Autoline constraints
        cogs_mfg = str(cfg.get("vehicle_cogs_manufact", "JEEP"))
        cogs_model = str(cfg.get("vehicle_cogs_model", "JEEP-G6"))
        cogs_sale = str(cfg.get("vehicle_cogs_saletype", "TRADE"))
        inv_mfg = str(cfg.get("vehicle_inv_manufact", "JEEP"))
        inv_model = str(cfg.get("vehicle_inv_model", "JEEP-G6"))
        
        # Resolve Vehicle Data (Stock No, VIN, and Cost)
        car_cost = 0.0
        stk_no = ""
        vin_no = ""
        
        if is_wg:
            # 1. Primary from Vehicle GL (511001)
            if vehicle_gl_dict and inv_no in vehicle_gl_dict:
                vgl_entry = vehicle_gl_dict[inv_no]
                car_cost = vgl_entry.get("cost", 0.0)
                stk_no = vgl_entry.get("stock_no") or ""
                vin_no = vgl_entry.get("vin") or ""
                if not stk_no and vgl_entry.get("identifier", "").startswith("N0"):
                    stk_no = vgl_entry.get("identifier")
                    
            # 2. Supplementary / Fallback from Stock and VIN dictionaries (from StockNumber file)
            if not stk_no and stock_dict:
                stk_no = stock_dict.get(inv_no, "")
                
            if not vin_no and vin_dict:
                vin_no = vin_dict.get(inv_no, "")
                
            # 3. Supplementary from Vehicle Profit (via VIN or Stock No cross-referencing)
            if vehicle_profit_dict:
                v_stk_map = vehicle_profit_dict.get("vin_to_stk", {})
                s_vin_map = vehicle_profit_dict.get("stk_to_vin", {})
                v_cost_map = vehicle_profit_dict.get("vin_to_cost", {})
                s_cost_map = vehicle_profit_dict.get("stk_to_cost", {})
                
                # Map VIN -> Stock No
                if not stk_no and vin_no:
                    stk_no = v_stk_map.get(vin_no, v_stk_map.get(vin_no[-8:] if len(vin_no) >= 8 else "")) or ""
                    
                # Map Stock No -> VIN
                if not vin_no and stk_no:
                    vin_no = s_vin_map.get(stk_no, "") or ""
                    
                # Recover missing Cost
                if car_cost <= 0:
                    if vin_no and vin_no in v_cost_map:
                        car_cost = v_cost_map[vin_no]
                    elif vin_no and len(vin_no) >= 8 and vin_no[-8:] in v_cost_map:
                        car_cost = v_cost_map[vin_no[-8:]]
                    elif stk_no and stk_no in s_cost_map:
                        car_cost = s_cost_map[stk_no]
                        
            # 4. Fallback from Legacy Cost dict
            if car_cost <= 0 and stk_no and cost_dict:
                car_cost = cost_dict.get(stk_no, 0.0)
                
            # Da's Template: {inv_no}_{Customer_Name}_{VIN}
            if vin_no:
                narrative = f"{inv_no}_{clean_cust}_{vin_no}" if clean_cust else f"{inv_no}_{vin_no}"
            else:
                narrative = f"{inv_no}_{clean_cust}" if clean_cust else inv_no
        else:
            narrative = f"{inv_no}_{cust_name}" if cust_name else inv_no
            
        if len(narrative) > 75:
            narrative = narrative[:75]
            
        itemized_lines = []
        
        if is_wg:
            fin_coy = fin_coy_dict.get(inv_no, "") if fin_coy_dict else ""
            fin_code = resolve_finance_subaccount(inv_no, cust_name, fin_coy, fin_data, default_code=None)
            
            if fin_code:
                if fin_code == "X0003":
                    cat_key = "CAR_SALE_CASH"
                    subaccount = "X0003"
                else:
                    cat_key = "CAR_SALE_CREDIT"
                    subaccount = fin_code
            elif is_finance_customer(cust_name):
                cat_key = "CAR_SALE_CREDIT"
                subaccount = cfg.get("subaccount_finance", "ARCODE FINANCE")
            else:
                cat_key = "CAR_SALE_CASH"
                subaccount = cfg["categories"]["CAR_SALE_CASH"]["subaccount"]
                
            cat_cfg = cfg["categories"][cat_key]
            itemized_lines.append({
                "category": cat_key,
                "gl": cat_cfg["rev_gl"],
                "dept": cat_cfg["rev_dept"],
                "net_amount": abs_net,
                "line_narrative": narrative
            })
            
        else:
            # 0XD / 0XDC: Check if fee line (FEE) is present along with other revenue lines
            rev_candidates = [
                l for l in gl_lines
                if l.get("cr") is not None and l["cr"] < 0 and (
                    str(l.get("acc", "")).startswith(("4", "216", "219")) or
                    not str(l.get("acc", "")).startswith(("213", "112", "G"))
                )
            ]
            
            has_fee = any(classify_description_text(l.get("desc", "")) == "FEE" for l in rev_candidates)
            split_lines_done = False
            
            if has_fee and len(rev_candidates) > 1:
                cat_amounts = {}
                for l in rev_candidates:
                    c_key = classify_description_text(l.get("desc", ""))
                    c_amt = abs(float(l.get("cr", 0.0)))
                    cat_amounts[c_key] = cat_amounts.get(c_key, 0.0) + c_amt
                    
                tot_rev = sum(cat_amounts.values())
                if tot_rev > 0 and "FEE" in cat_amounts and len(cat_amounts) > 1:
                    allocated = 0.0
                    # Put other categories first (e.g. ACCESSORIES), and FEE last
                    ordered_cats = [c for c in cat_amounts if c != "FEE"] + ["FEE"]
                    for idx_c, c_key in enumerate(ordered_cats):
                        if idx_c == len(ordered_cats) - 1:
                            c_net = round(abs_net - allocated, 2)
                        else:
                            c_net = round(abs_net * (cat_amounts[c_key] / tot_rev), 2)
                            allocated += c_net
                            
                        c_cfg = cfg["categories"].get(c_key, cfg["categories"]["ACCESSORIES"])
                        itemized_lines.append({
                            "category": c_key,
                            "gl": c_cfg["rev_gl"],
                            "dept": c_cfg["rev_dept"],
                            "net_amount": c_net,
                            "line_narrative": narrative
                        })
                    split_lines_done = True
                    subaccount = "X0003"
                    matched_corp = resolve_unified_arcode(dealerpro_code=None, customer_name=cust_name, context="sale", default_override=None)
                    if matched_corp and matched_corp != "X0003":
                        subaccount = matched_corp
                        
            if not split_lines_done:
                cat_priority = ['PARTS_DEPOSIT', 'FEE', 'COMM_FINANCE', 'REGISTRATION', 'ACCESSORIES', 'RED_PLATE', 'DEPOSIT']
                main_cat = 'DEPOSIT'
                if gl_lines:
                    found_cats = set()
                    for l in gl_lines:
                        d = l.get('desc', '')
                        if d:
                            found_cats.add(classify_description_text(d))
                    for p in cat_priority:
                        if p in found_cats:
                            main_cat = p
                            break
                else:
                    c_cat = classify_description_text(cust_name)
                    if c_cat in ['PARTS_DEPOSIT', 'FEE', 'COMM_FINANCE', 'REGISTRATION', 'ACCESSORIES', 'RED_PLATE']:
                        main_cat = c_cat
                            
                main_cat_cfg = cfg["categories"].get(main_cat, cfg["categories"]["DEPOSIT"])
                itemized_lines.append({
                    "category": main_cat,
                    "gl": main_cat_cfg["rev_gl"],
                    "dept": main_cat_cfg["rev_dept"],
                    "net_amount": abs_net,
                    "line_narrative": narrative
                })
                
                if main_cat == "PARTS_DEPOSIT":
                    subaccount = "D0001"
                elif main_cat == "COMM_FINANCE":
                    fin_code = resolve_finance_subaccount(inv_no, cust_name, "", fin_data, default_code=None)
                    if fin_code:
                        subaccount = fin_code
                    else:
                        subaccount = main_cat_cfg["subaccount"]
                else:
                    subaccount = main_cat_cfg["subaccount"]
                    matched_corp = resolve_unified_arcode(dealerpro_code=None, customer_name=cust_name, context="sale", default_override=None)
                    if matched_corp and matched_corp != "X0003":
                        subaccount = matched_corp
                        
        main_cat = itemized_lines[0]["category"]
        main_cat_cfg = cfg["categories"].get(main_cat, cfg["categories"]["DEPOSIT"])
        
        ar_dept = main_cat_cfg.get("ar_dept", cfg.get("ar_department", "0000"))
        ar_gl = int(main_cat_cfg.get("ar_gl", cfg.get("ar_gl_code", 11311001)))
        
        for itm in itemized_lines:
            c_name = itm["category"]
            if c_name in cat_counts:
                cat_counts[c_name] += 1
                cat_values[c_name] += itm["net_amount"]
        
        # Header Record
        header_record = {
            "type": "HEADER",
            "A": inv_no,
            "B": inv_no,
            "C": 1,
            "D": None,
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
            "O": doc_date,
            "P": subaccount,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "U": "U",
            "V": None,
            "W": int(cfg.get("terms", 30)),
            "X": branch,
            "Y": "00000004" if is_cn else None
        }
        add_record(header_record)
        
        # Blank Row immediately following Header (Autoline specification)
        add_record({"type": "BLANK"})
        
        # Line 1: AR / Bank Line
        line1_debit = abs_gross if not is_cn else None
        line1_credit = abs_gross if is_cn else None
        
        line1_record = {
            "type": "DETAIL",
            "D": None,
            "G": cfg["currency"],
            "H": 1,
            "I": branch,
            "J": ar_dept,
            "K": ar_gl,
            "L": stk_no if (is_wg and stk_no) else None,
            "M": line1_debit,
            "N": line1_credit,
            "O": narrative
        }
        add_record(line1_record)
        
        # Line 2+: Revenue / Liability Lines
        line_num = 2
        for itm in itemized_lines:
            itm_net = itm["net_amount"]
            itm_debit = itm_net if is_cn else None
            itm_credit = itm_net if not is_cn else None
            
            line_record = {
                "type": "DETAIL",
                "D": None,
                "G": cfg["currency"],
                "H": line_num,
                "I": branch,
                "J": itm["dept"],
                "K": itm["gl"],
                "L": stk_no if (is_wg and stk_no) else None,
                "M": itm_debit,
                "N": itm_credit,
                "O": narrative if is_wg else itm["line_narrative"],
                "Q": cfg["tax_code"],
                "U": cogs_mfg if is_wg else None,
                "V": cogs_model if is_wg else None,
                "W": cogs_sale if is_wg else None
            }
            add_record(line_record)
            line_num += 1
            
        # Vehicle Cost Lines (Primary: Vehicle GL 511001, Fallback: Legacy Stock & Cost files)
        if car_cost > 0:
            cogs_gl = int(cfg.get("vehicle_cogs_gl", 51111012))
            cogs_dept = str(cfg.get("vehicle_cogs_dept", "2002"))
            cogs_tax = cfg.get("vehicle_cogs_tax", "O")
            
            cogs_debit = car_cost if not is_cn else None
            cogs_credit = None if not is_cn else car_cost
            
            # Line 3: COGS (51111012)
            cogs_record = {
                "type": "DETAIL",
                "D": None,
                "G": cfg["currency"],
                "H": line_num,
                "I": branch,
                "J": cogs_dept,
                "K": cogs_gl,
                "L": stk_no if stk_no else None,
                "M": cogs_debit,
                "N": cogs_credit,
                "O": narrative,
                "Q": cogs_tax,
                "U": cogs_mfg,
                "V": cogs_model,
                "W": cogs_sale
            }
            add_record(cogs_record)
            if cogs_debit is not None:
                total_debit_sum += cogs_debit
            if cogs_credit is not None:
                total_credit_sum += cogs_credit
            line_num += 1
            
            # Line 4: Inventory (11511111)
            inv_gl = int(cfg.get("vehicle_inv_gl", 11511111))
            inv_dept = str(cfg.get("vehicle_inv_dept", "0000"))
            inv_tax = cfg.get("vehicle_inv_tax", "O")
            
            inv_debit = car_cost if is_cn else None
            inv_credit = car_cost if not is_cn else None
            
            inv_record = {
                "type": "DETAIL",
                "D": None,
                "G": cfg["currency"],
                "H": line_num,
                "I": branch,
                "J": inv_dept,
                "K": inv_gl,
                "L": stk_no if stk_no else None,
                "M": inv_debit,
                "N": inv_credit,
                "O": narrative,
                "Q": inv_tax,
                "U": inv_mfg,
                "V": inv_model,
                "W": None
            }
            add_record(inv_record)
            if inv_debit is not None:
                total_debit_sum += inv_debit
            if inv_credit is not None:
                total_credit_sum += inv_credit
            line_num += 1
            
            matched_cost_count += 1
            total_vehicle_cost += car_cost
            
        add_record({"type": "BLANK"})
        
        total_tax_sum += abs_tax
        if not is_cn:
            total_debit_sum += abs_gross
            total_credit_sum += abs_net
        else:
            total_credit_sum += abs_gross
            total_debit_sum += abs_net
            
        p_row = {
            "Batch": b_batch_idx,
            "Invoice": inv_no,
            "Type": doc_code,
            "Date": doc_date_str,
            "Branch": branch,
            "Category": main_cat_cfg["name"],
            "Customer": clean_cust if is_wg else cust_name,
            "Subaccount": subaccount,
            "Net": abs_net if not is_cn else -abs_net,
            "VAT": abs_tax if not is_cn else -abs_tax,
            "Gross": abs_gross if not is_cn else -abs_gross,
            "Stock No": stk_no or "",
            "VIN": vin_no or "",
            "Vehicle Cost": car_cost if car_cost > 0 else 0.0,
            "Items": len(itemized_lines) + (2 if car_cost > 0 else 0)
        }
        preview_rows.append(p_row)
        branch_previews[branch].append(p_row)
        doc_type_previews[doc_group].append(p_row)
        
        # Missing Data Tracking
        gaps = []
        gap_details = []
        if is_wg:
            if not stk_no:
                gaps.append("ไม่มีเลขสต๊อกรถ (Stock No)")
                gap_details.append("คอลัมน์ L (SUFFIX) ว่างเปล่า")
            if not vin_no:
                gaps.append("ไม่มีเลขตัวถัง (VIN)")
                gap_details.append("คำอธิบายไม่มีเลขตัวถัง VIN")
            if not is_cn and car_cost <= 0:
                gaps.append("ไม่มีบันทึกต้นทุนรถ (Cost = 0)")
                gap_details.append("ไม่มีการลงบันทึกต้นทุน Line 3 และ 4")
            if subaccount == "ARCODE FINANCE":
                gaps.append("ไม่พบรหัสไฟแนนซ์เฉพาะ")
                gap_details.append("ใช้ Subaccount กลางเป็น ARCODE FINANCE")
        else:
            if not gl_lines and inv_no.startswith(('01D', '02D', '01DC', '02DC', '03D', '04D')):
                gaps.append("ไม่พบในรายงาน GL ทั่วไป")
                gap_details.append("ใช้หมวดเริ่มต้นเป็นเงินจอง (GL 21311003 แผนก 0000)")
            if subaccount == "ARCODE FINANCE":
                gaps.append("ไม่พบรหัสไฟแนนซ์เฉพาะ")
                gap_details.append("ใช้ Subaccount กลางเป็น ARCODE FINANCE")
                
        if gaps:
            missing_records.append({
                "Invoice": inv_no,
                "Type": doc_code,
                "Branch": branch,
                "Group": doc_group,
                "Customer": cust_name,
                "Gross": abs_gross if not is_cn else -abs_gross,
                "Missing_Fields": ", ".join(gaps),
                "Details": " | ".join(gap_details)
            })
        
        b_st = branch_stats_data[branch]
        b_st["total_invoices"] += 1
        b_st["total_batches"] = b_batch_idx
        b_st["total_tax"] += abs_tax
        if not is_cn:
            b_st["total_debit"] += abs_gross
            b_st["total_credit"] += abs_net
        else:
            b_st["total_credit"] += abs_gross
            b_st["total_debit"] += abs_net
        if car_cost > 0:
            b_st["matched_cost_count"] += 1
            b_st["total_vehicle_cost"] += car_cost
            
        dt_st = doc_type_stats_data[doc_group]
        dt_st["total_invoices"] += 1
        dt_st["total_batches"] = dt_batch_idx
        dt_st["total_tax"] += abs_tax
        if not is_cn:
            dt_st["total_debit"] += abs_gross
            dt_st["total_credit"] += abs_net
        else:
            dt_st["total_credit"] += abs_gross
            dt_st["total_debit"] += abs_net
        if car_cost > 0:
            dt_st["matched_cost_count"] += 1
            dt_st["total_vehicle_cost"] += car_cost
            
        global_batch_idx += 1
        
    by_branch = {}
    for b in sorted(branch_batches.keys()):
        b_st = branch_stats_data[b]
        by_branch[b] = {
            "branch": b,
            "rows": branch_rows[b],
            "total_invoices": b_st["total_invoices"],
            "total_batches": b_st["total_batches"],
            "total_debit": round(b_st["total_debit"], 2),
            "total_credit": round(b_st["total_credit"], 2),
            "total_tax": round(b_st["total_tax"], 2),
            "total_vehicle_cost": round(b_st["total_vehicle_cost"], 2),
            "matched_cost_count": b_st["matched_cost_count"],
            "preview": pd.DataFrame(branch_previews[b])
        }
        
    by_doc_type = {}
    for dt_key in ["0XD", "0XWG"]:
        if dt_key in doc_type_batches:
            dt_st = doc_type_stats_data[dt_key]
            by_doc_type[dt_key] = {
                "doc_type": dt_key,
                "rows": doc_type_rows[dt_key],
                "total_invoices": dt_st["total_invoices"],
                "total_batches": dt_st["total_batches"],
                "total_debit": round(dt_st["total_debit"], 2),
                "total_credit": round(dt_st["total_credit"], 2),
                "total_tax": round(dt_st["total_tax"], 2),
                "total_vehicle_cost": round(dt_st["total_vehicle_cost"], 2),
                "matched_cost_count": dt_st["matched_cost_count"],
                "preview": pd.DataFrame(doc_type_previews[dt_key])
            }
        
    summary_stats = {
        "total_invoices": len(df_vat),
        "total_batches": global_batch_idx - 1,
        "total_debit": round(total_debit_sum, 2),
        "total_credit": round(total_credit_sum, 2),
        "total_tax": round(total_tax_sum, 2),
        "total_vehicle_cost": round(total_vehicle_cost, 2),
        "matched_cost_count": matched_cost_count,
        "unmatched_gl_count": unmatched_gl_count,
        "category_counts": {cfg["categories"][k]["name"]: cat_counts[k] for k in cat_counts},
        "category_values": {cfg["categories"][k]["name"]: cat_values[k] for k in cat_values},
        "by_branch": by_branch,
        "by_doc_type": by_doc_type,
        "missing_records": missing_records,
        "missing_count": len(missing_records),
        "missing_df": pd.DataFrame(missing_records)
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
    subtab_service, subtab_parts = st.tabs([
        "🛠️ งานบริการศูนย์บริการ (Service Workshop: 01S / 02S)",
        "📦 งานขายอะไหล่หน้าร้าน (Parts Sales: 01P / 01PC)"
    ])

    # -------------------------------------------------------------------------
    # SUB-TAB 1: SERVICE WORKSHOP (01S / 02S)
    # -------------------------------------------------------------------------
    with subtab_service:
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

    # -------------------------------------------------------------------------
    # SUB-TAB 2: PARTS OVER-THE-COUNTER SALES (01P / 02P / 01PC / 02PC)
    # -------------------------------------------------------------------------
    with subtab_parts:
        st.subheader("📦 อัปโหลดรายงานการขายอะไหล่ (Parts Sales Report)")
        st.caption("ไฟล์รายงานการขายอะไหล่ เช่น `AftersalePart2026.xlsx` (มีเลขที่เอกสาร 01P/01PC, รหัสลูกค้า, อะไหล่, จำนวน, ราคาขาย, ต้นทุน, VAT)")
        parts_file = st.file_uploader("เลือกไฟล์รายงานการขายอะไหล่ (.xlsx / .xls)", type=["xlsx", "xls"], key="parts_file_upload")
        
        if parts_file:
            try:
                with st.spinner("กำลังแปลงข้อมูลขายอะไหล่และคำนวณดุลบัญชีอัตโนมัติ..."):
                    invoices_dict = parse_parts_sales_file(parts_file)
                    if not invoices_dict:
                        st.warning("⚠️ ไม่พบข้อมูลบิลขายอะไหล่ (01P / 01PC) ในไฟล์ที่อัปโหลด กรุณาตรวจสอบรูปแบบไฟล์")
                    else:
                        rows_parts, stats_parts, preview_parts = transform_parts_sales_to_autoline(invoices_dict)
                        
                        st.success(f"✅ แปลงข้อมูลขายอะไหล่สำเร็จเรียบร้อย! พบเอกสารทั้งหมด {stats_parts['total_documents']:,} ฉบับ")
                        
                        # KPI Dashboard
                        st.markdown("### 📊 สรุปตัวเลขและความสมดุลทางบัญชี (Balance Verification)")
                        kp1, kp2, kp3, kp4, kp5, kp6 = st.columns(6)
                        
                        cn_txt = f"Inv: {stats_parts['count_inv']:,} | CN: {stats_parts['count_cn']:,}" if stats_parts['count_cn'] > 0 else f"{stats_parts['count_inv']:,} ใบ"
                        kp1.metric("จำนวนเอกสารทั้งหมด", f"{stats_parts['total_documents']:,} ฉบับ", delta=cn_txt if stats_parts['count_cn'] > 0 else None)
                        
                        kp2.metric(
                            "ยอดขายสุทธิ (Revenue)",
                            f"{stats_parts['net_revenue']:,.2f} ฿",
                            delta=f"CN: -{stats_parts['cn_revenue']:,.2f} ฿" if stats_parts['count_cn'] > 0 else None
                        )
                        kp3.metric(
                            "ภาษี 7% (VAT)",
                            f"{stats_parts['net_tax']:,.2f} ฿",
                            delta=f"CN: -{stats_parts['cn_tax']:,.2f} ฿" if stats_parts['count_cn'] > 0 else None
                        )
                        kp4.metric(
                            "ยอดลูกหนี้สุทธิ (AR)",
                            f"{stats_parts['net_ar']:,.2f} ฿",
                            delta=f"CN: -{stats_parts['cn_ar']:,.2f} ฿" if stats_parts['count_cn'] > 0 else None
                        )
                        kp5.metric(
                            "ต้นทุนอะไหล่ (COGS)",
                            f"{stats_parts['total_parts_cogs']:,.2f} ฿",
                            delta="Dr 51211006 / Cr 11511112"
                        )
                        
                        if stats_parts["is_balanced"]:
                            kp6.metric("สถานะดุลบัญชี", "สมดุล 100% ✅", delta="ผลต่าง 0.00 บาท")
                        else:
                            kp6.metric("สถานะดุลบัญชี", "พบผลต่าง ⚠️", delta=f"{stats_parts['difference']:,.2f} บาท")
                            
                        # Download Section
                        st.markdown("---")
                        col_pdl1, col_pdl2 = st.columns([2, 1])
                        with col_pdl1:
                            st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline (Parts Sales)")
                            st.caption("ไฟล์ Excel ในโครงสร้าง Autoline AR Journal Import (รองรับทั้งบิลขาย ARI และใบลดหนี้ ARC พร้อมบันทึกต้นทุนและสต็อก)")
                        with col_pdl2:
                            excel_parts_output = generate_output_excel(template_path, rows_parts)
                            now_p_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                            st.download_button(
                                label="⬇️ ดาวน์โหลด Autoline_Import_PARTS.xlsx",
                                data=excel_parts_output,
                                file_name=f"Autoline_Import_PARTS_{now_p_str}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_parts_excel_btn",
                                use_container_width=True
                            )
                            
                        # Data Preview
                        st.markdown("### 🔍 ตรวจสอบข้อมูลรายการขายอะไหล่ (Parts Data Preview)")
                        parts_search_kw = st.text_input(
                            "ค้นหาเอกสาร (เลขที่บิล, ชื่อลูกค้า, รหัส GL, Subaccount หรือ Narrative)",
                            placeholder="เช่น 01P26030001, 01PC, MMS, ไลอ้อน, 41211001",
                            key="parts_search_kw"
                        )
                        
                        display_parts_df = preview_parts.copy()
                        if parts_search_kw.strip():
                            pkw = parts_search_kw.strip().lower()
                            pmask = (
                                display_parts_df["Invoice"].astype(str).str.lower().str.contains(pkw, na=False) |
                                display_parts_df["Narrative"].astype(str).str.lower().str.contains(pkw, na=False) |
                                display_parts_df["Subaccount"].astype(str).str.lower().str.contains(pkw, na=False) |
                                display_parts_df["GL Code"].astype(str).str.lower().str.contains(pkw, na=False) |
                                display_parts_df["Row Type"].astype(str).str.lower().str.contains(pkw, na=False)
                            )
                            display_parts_df = display_parts_df[pmask]
                            
                        st.dataframe(display_parts_df, use_container_width=True, height=450)
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผลขายอะไหล่: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
        else:
            st.info("💡 กรุณาอัปโหลดไฟล์รายงานการขายอะไหล่ (.xlsx) ด้านบนเพื่อเริ่มต้นแปลงข้อมูลขายอะไหล่หน้าร้าน")

# =============================================================================
# TAB 2: VEHICLE SALES INTERFACE (0XD / 0XDC / 0XWG / 0XWGCN)
# =============================================================================
with tab_sales:
    # -------------------------------------------------------------------------
    # 1. MANDATORY INPUTS (2 FILES)
    # -------------------------------------------------------------------------
    st.markdown("### 🔴 ไฟล์จำเป็น (Mandatory Inputs - ใช้เพียง 2 ไฟล์)")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown("**1. รายงานภาษีขาย (Sales VAT Report)**")
        st.caption("เช่น `2026VatReport(AllBranch).xlsx` *จำเป็น")
        vat_file = st.file_uploader(
            "เลือกไฟล์รายงานภาษีขาย (.xlsx / .xls)", 
            type=["xlsx", "xls"], 
            key="sales_vat_upload"
        )
    with col_m2:
        st.markdown("**2. บัญชีแยกประเภททุกหมวด (All Category GL Report)**")
        st.caption("เช่น `บัญชีแยกประเภท2026(AllCategory).xlsx` (ไฟล์เดียวครอบคลุมจัดหมวดบิล 0XD และดึงต้นทุน/สต๊อก/VIN บิล 0XWG ครบ 100%) *จำเป็น")
        gl_file = st.file_uploader(
            "เลือกไฟล์บัญชีแยกประเภททุกหมวด (.xlsx / .xls)", 
            type=["xlsx", "xls"], 
            key="sales_gl_upload"
        )
        
    # -------------------------------------------------------------------------
    # 2. OPTIONAL INPUTS (2 FILES)
    # -------------------------------------------------------------------------
    st.markdown("### 🟡 ไฟล์เสริม (Optional Inputs)")
    st.caption("เสริมข้อมูลเพื่อดึงเลขสต๊อก (Stock No), เลขตัวถัง (VIN) และต้นทุนรถยนต์เพิ่มเติม (หากไม่อัปโหลด ระบบจะดึงจากโฟลเดอร์ให้อัตโนมัติ)")
    with st.expander("📁 อัปโหลดไฟล์เสริม (VehicleProfit / StockNumber)", expanded=False):
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            st.markdown("**3. รายงานกำไรการขายรถ (Vehicle Profit Report)**")
            st.caption("เช่น `VehicleProfit2026(AllBranch).xlsx` (ช่วยกู้คืนต้นทุนรถ, เลขสต๊อก, และ VIN)")
            profit_file = st.file_uploader(
                "เลือกไฟล์รายงานกำไรการขายรถ (.xlsx / .xls)", 
                type=["xlsx", "xls"], 
                key="sales_profit_upload"
            )
        with col_o2:
            st.markdown("**4. รายงานเลขสต๊อกรถยนต์ (Stock Number Report)**")
            st.caption("เช่น `StockNumber2026.xlsx` (ช่วยแมปเลขสต๊อกและ VIN เพิ่มเติม)")
            stock_file = st.file_uploader(
                "เลือกไฟล์เลขสต๊อกรถยนต์ (.xlsx / .xls)", 
                type=["xlsx", "xls"], 
                key="sales_stock_upload"
            )
        
    # -------------------------------------------------------------------------
    # 3. ADVANCED CONFIGURATION EXPANDER
    # -------------------------------------------------------------------------
    with st.expander("⚙️ ตั้งค่า Fixed Values สำหรับฝ่ายขาย (Sales Settings)", expanded=False):
        cfg_col1, cfg_col2, cfg_col3 = st.columns(3)
        with cfg_col1:
            sales_subacc_default = st.text_input("Subaccount ลูกค้าทั่วไป (Default)", value="X0003")
            veh_cogs_gl = st.text_input("GL Code ต้นทุนรถ (Vehicle COGS)", value="51111012")
            veh_inv_gl = st.text_input("GL Code สินค้าคงเหลือ (Inventory)", value="11511111")
        with cfg_col2:
            sales_subacc_finance = st.text_input("Subaccount สถาบันการเงิน (Finance Default)", value="ARCODE FINANCE")
            veh_cogs_dept = st.text_input("Department ต้นทุนรถ", value="2002")
            veh_inv_dept = st.text_input("Department สินค้าคงเหลือ", value="0000")
        with cfg_col3:
            sales_dept = st.text_input("รหัสแผนกฝ่ายขาย (Sales Department)", value="2002")
            veh_mfg = st.text_input("MANUFACT (ยี่ห้อ)", value="JEEP")
            veh_model = st.text_input("MODEL (รุ่น)", value="JEEP-G6")
            veh_saletype = st.text_input("SALETYPE (ประเภทการขาย)", value="TRADE")
            
    sales_config_override = {
        "subaccount_default": sales_subacc_default,
        "subaccount_finance": sales_subacc_finance,
        "vehicle_cogs_gl": int(veh_cogs_gl) if str(veh_cogs_gl).isdigit() else 51111012,
        "vehicle_cogs_dept": veh_cogs_dept,
        "vehicle_cogs_manufact": veh_mfg,
        "vehicle_cogs_model": veh_model,
        "vehicle_cogs_saletype": veh_saletype,
        "vehicle_inv_gl": int(veh_inv_gl) if str(veh_inv_gl).isdigit() else 11511111,
        "vehicle_inv_dept": veh_inv_dept,
        "vehicle_inv_manufact": veh_mfg,
        "vehicle_inv_model": veh_model,
    }
    sales_config_override["categories"] = {
        "DEPOSIT": {"subaccount": sales_subacc_default},
        "RED_PLATE": {"subaccount": sales_subacc_default},
        "REGISTRATION": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "ACCESSORIES": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "COMM_FINANCE": {"subaccount": sales_subacc_finance, "rev_dept": sales_dept},
        "CAR_SALE_CASH": {"subaccount": sales_subacc_default, "rev_dept": sales_dept},
        "CAR_SALE_CREDIT": {"subaccount": sales_subacc_finance, "rev_dept": sales_dept}
    }
    
    # -------------------------------------------------------------------------
    # 4. PROCESSING LOGIC
    # -------------------------------------------------------------------------
    if vat_file and gl_file:
        try:
            with st.spinner("กำลังประมวลผลข้อมูลฝ่ายขาย แมปต้นทุน และจัดหมวดหมู่อัตโนมัติ..."):
                df_vat_sales = load_sales_vat_report(vat_file)
                
                # 1. Load General GL descriptions for 0XD / 0XDC bills
                gl_dict_sales = load_gl_descriptions(gl_file)
                
                # 2. Reset stream and load Vehicle GL COGS for 0XWG / 0XWGCN bills from AllCategory
                if hasattr(gl_file, "seek"):
                    gl_file.seek(0)
                veh_gl_dict_sales = load_vehicle_gl_cogs(gl_file)
                
                # Optional Profit file (User upload or local workspace fallback)
                profit_dict_sales = None
                if profit_file:
                    profit_dict_sales = load_vehicle_profit(profit_file)
                else:
                    curr_dir = os.path.dirname(__file__) if "__file__" in locals() else "."
                    cand_profits = [f for f in os.listdir(curr_dir) if "VehicleProfit" in f and f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
                    if cand_profits:
                        try:
                            profit_dict_sales = load_vehicle_profit(os.path.join(curr_dir, cand_profits[0]))
                        except Exception:
                            pass
                
                # Optional Stock file (User upload or local workspace fallback)
                stock_dict_sales, fin_coy_sales, vin_dict_sales = ({}, {}, {})
                if stock_file:
                    stock_dict_sales, fin_coy_sales, vin_dict_sales = load_stock_numbers(stock_file)
                else:
                    curr_dir = os.path.dirname(__file__) if "__file__" in locals() else "."
                    cand_stocks = [f for f in os.listdir(curr_dir) if "StockNumber" in f and f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
                    if cand_stocks:
                        try:
                            stock_dict_sales, fin_coy_sales, vin_dict_sales = load_stock_numbers(os.path.join(curr_dir, cand_stocks[0]))
                        except Exception:
                            pass
                            
                fin_data_sales = DEFAULT_FINANCE_DATA
                
                rows_sales, stats_sales, preview_sales = transform_sales_to_autoline(
                    df_vat_sales,
                    gl_dict_sales,
                    stock_dict=stock_dict_sales,
                    cost_dict={},
                    fin_coy_dict=fin_coy_sales,
                    fin_data=fin_data_sales,
                    vehicle_gl_dict=veh_gl_dict_sales,
                    vehicle_profit_dict=profit_dict_sales,
                    vin_dict=vin_dict_sales,
                    user_config=sales_config_override
                )
                
            st.success("✅ ประมวลผลข้อมูลฝ่ายขายสำเร็จเรียบร้อย!")
            
            # -----------------------------------------------------------------
            # 5. KPI SUMMARY METRICS
            # -----------------------------------------------------------------
            st.markdown("### 📊 สรุปข้อมูลธุรกรรมฝ่ายขาย (Vehicle Sales Summary)")
            skpi1, skpi2, skpi3, skpi4, skpi5 = st.columns(5)
            skpi1.metric("จำนวนบิลที่แปลง (Invoices)", f"{stats_sales['total_invoices']:,} ฉบับ")
            skpi2.metric("จำนวนรายการบัญชี (Lines)", f"{len(rows_sales):,} บรรทัด")
            
            cogs_delta = f"ต้นทุน {stats_sales['total_vehicle_cost']:,.2f} ฿ ({stats_sales['matched_cost_count']} คัน)" if stats_sales['matched_cost_count'] > 0 else None
            skpi3.metric("ยอดรวมเดบิต (Total Debit)", f"{stats_sales['total_debit']:,.2f} ฿", delta=cogs_delta)
            skpi4.metric("ยอดรวมภาษี (Total Tax)", f"{stats_sales['total_tax']:,.2f} ฿")
            
            miss_count = stats_sales.get("missing_count", 0)
            skpi5.metric(
                "เอกสารข้อมูลไม่ครบ", 
                f"{miss_count:,} ฉบับ", 
                delta=None if miss_count == 0 else f"{miss_count} ฉบับที่ต้องติดตาม",
                delta_color="inverse"
            )
            
            # Branch Breakdown Cards
            by_branch = stats_sales.get("by_branch", {})
            if len(by_branch) > 1:
                st.markdown("#### 🏢 จำแนกยอดตามสาขา (Branch Breakdown)")
                branch_cols = st.columns(len(by_branch))
                for idx_b, (b_code, b_data) in enumerate(sorted(by_branch.items())):
                    with branch_cols[idx_b]:
                        b_name = "สำนักงานใหญ่" if b_code == "0001" else ("เกษตรนวมินทร์" if b_code == "0002" else f"สาขา {b_code}")
                        st.info(
                            f"**🏢 สาขา {b_code} ({b_name})**\n"
                            f"- จำนวนบิล: **{b_data['total_invoices']:,}** ฉบับ\n"
                            f"- ยอดเดบิต: **{b_data['total_debit']:,.2f}** ฿\n"
                            f"- ภาษีขาย: **{b_data['total_tax']:,.2f}** ฿\n"
                            f"- ต้นทุนรถ: **{b_data['total_vehicle_cost']:,.2f}** ฿ ({b_data['matched_cost_count']} คัน)"
                        )
            
            # Category Breakdown
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
            
            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # -----------------------------------------------------------------
            # 6. MISSING DATA TRACKER / AUDIT REPORT
            # -----------------------------------------------------------------
            st.markdown("---")
            st.subheader("📋 รายงานตรวจสอบข้อมูลตกหล่น (Missing Data Audit Report)")
            
            if miss_count > 0:
                st.warning(
                    f"⚠️ พบเอกสารที่มีข้อมูลตกหล่น/ไม่ครบถ้วนทั้งหมด **{miss_count} ฉบับ** "
                    f"(ท่านสามารถตรวจสอบเลขที่บิลและรายการที่ตกหล่นด้านล่าง เพื่อนำไปค้นหาหรือปรับแก้ในไฟล์ผลลัพธ์ได้อย่างแม่นยำ)"
                )
                
                # Excel download buffer for missing data report
                missing_excel_buffer = BytesIO()
                with pd.ExcelWriter(missing_excel_buffer, engine="openpyxl") as writer:
                    stats_sales["missing_df"].to_excel(writer, index=False, sheet_name="Missing_Invoices")
                missing_excel_buffer.seek(0)
                
                col_mis1, col_mis2, col_mis3 = st.columns([2, 1, 1])
                with col_mis1:
                    st.download_button(
                        label=f"📥 ดาวน์โหลดรายงานข้อมูลตกหล่น (Missing_Data_Report_{now_str}.xlsx)",
                        data=missing_excel_buffer,
                        file_name=f"Missing_Data_Report_{now_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_missing_data_excel",
                        use_container_width=True
                    )
                with col_mis2:
                    grp_filter = st.selectbox(
                        "กรองกลุ่มเอกสาร", 
                        ["ทั้งหมด (All)", "0XD (เงินจอง/ป้าย/อุปกรณ์)", "0XWG (ขายรถยนต์)"], 
                        key="mis_grp_filter"
                    )
                with col_mis3:
                    f_branch_opts = ["ทุกสาขา (All)"] + [f"สาขา {b}" for b in sorted(by_branch.keys())]
                    b_mis_filter = st.selectbox("กรองสาขา", f_branch_opts, key="mis_branch_filter")
                    
                disp_mis_df = stats_sales["missing_df"].copy()
                if grp_filter.startswith("0XD"):
                    disp_mis_df = disp_mis_df[disp_mis_df["Group"] == "0XD"]
                elif grp_filter.startswith("0XWG"):
                    disp_mis_df = disp_mis_df[disp_mis_df["Group"] == "0XWG"]
                    
                if b_mis_filter != "ทุกสาขา (All)":
                    b_target = b_mis_filter.split()[1]
                    disp_mis_df = disp_mis_df[disp_mis_df["Branch"] == b_target]
                    
                st.dataframe(disp_mis_df, use_container_width=True, height=280)
            else:
                st.success("🎉 สมบูรณ์แบบ! ไม่พบเอกสารที่มีข้อมูลตกหล่น ข้อมูลครบถ้วน 100%")
                
            # -----------------------------------------------------------------
            # 7. DOWNLOAD BUTTONS (BY INVOICE TYPE 0XD / 0XWG, ALL COMBINED, BY BRANCH)
            # -----------------------------------------------------------------
            st.markdown("---")
            st.subheader("📥 ดาวน์โหลดไฟล์สำหรับ Autoline (Sales Import Files)")
            st.caption("เลือกรูปแบบการดาวน์โหลดที่ต้องการ: แยกตามประเภทบิล (0XD vs 0XWG), รวมทั้งหมด, หรือแยกตามรายสาขา (แต่ละไฟล์รันเลข Batch เริ่มจาก 1 ใหม่โดยอัตโนมัติ)")
            
            dl_choice = st.radio(
                "เลือกรูปแบบการดาวน์โหลด (Download Format):",
                [
                    "📑 แยกตามประเภทบิล (0XD / 0XWG)",
                    "📦 รวมทั้งหมด (Combined All Invoices)",
                    "🏢 แยกตามสาขา (By Branch: 0001 / 0002)"
                ],
                horizontal=True,
                key="sales_dl_format_choice"
            )
            
            by_dt = stats_sales.get("by_doc_type", {})
            
            # OPTION 1: SEPARATED BY INVOICE TYPE (0XD / 0XWG)
            if dl_choice == "📑 แยกตามประเภทบิล (0XD / 0XWG)":
                col_dt1, col_dt2 = st.columns(2)
                dt_0xd = by_dt.get("0XD")
                dt_0xwg = by_dt.get("0XWG")
                
                with col_dt1:
                    st.markdown("#### 1. บิลกลุ่ม 0XD (เงินจอง / ป้ายแดง / ค่าจดทะเบียน / อุปกรณ์ / คอมไฟแนนซ์)")
                    if dt_0xd:
                        st.info(
                            f"- จำนวนบิล: **{dt_0xd['total_invoices']:,}** ฉบับ (Batch 1 - {dt_0xd['total_batches']:,})\n"
                            f"- จำนวนแถว: **{len(dt_0xd['rows']):,}** บรรทัด\n"
                            f"- ยอดเดบิตรวม: **{dt_0xd['total_debit']:,.2f}** ฿\n"
                            f"- ภาษีขาย: **{dt_0xd['total_tax']:,.2f}** ฿"
                        )
                        excel_0xd = generate_output_excel(template_path, dt_0xd["rows"])
                        st.download_button(
                            label=f"⬇️ ดาวน์โหลดไฟล์กลุ่ม 0XD ({dt_0xd['total_invoices']:,} ฉบับ)",
                            data=excel_0xd,
                            file_name=f"Autoline_Import_SALES_0XD_{now_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_sales_0xd_btn",
                            use_container_width=True
                        )
                    else:
                        st.warning("ไม่พบเอกสารกลุ่ม 0XD")
                        
                with col_dt2:
                    st.markdown("#### 2. บิลกลุ่ม 0XWG (ขายรถยนต์ / ใบลดหนี้ พร้อมบันทึกต้นทุน COGS)")
                    if dt_0xwg:
                        st.info(
                            f"- จำนวนบิล: **{dt_0xwg['total_invoices']:,}** ฉบับ (Batch 1 - {dt_0xwg['total_batches']:,})\n"
                            f"- จำนวนแถว: **{len(dt_0xwg['rows']):,}** บรรทัด\n"
                            f"- ยอดเดบิตรวม: **{dt_0xwg['total_debit']:,.2f}** ฿\n"
                            f"- ต้นทุนรถยนต์: **{dt_0xwg['total_vehicle_cost']:,.2f}** ฿ ({dt_0xwg['matched_cost_count']} คัน)"
                        )
                        excel_0xwg = generate_output_excel(template_path, dt_0xwg["rows"])
                        st.download_button(
                            label=f"⬇️ ดาวน์โหลดไฟล์กลุ่ม 0XWG ({dt_0xwg['total_invoices']:,} ฉบับ)",
                            data=excel_0xwg,
                            file_name=f"Autoline_Import_SALES_0XWG_{now_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_sales_0xwg_btn",
                            use_container_width=True
                        )
                    else:
                        st.warning("ไม่พบเอกสารกลุ่ม 0XWG")
                        
                if dt_0xd and dt_0xwg:
                    import zipfile
                    zip_dt_buffer = BytesIO()
                    with zipfile.ZipFile(zip_dt_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr(f"Autoline_Import_SALES_0XD_{now_str}.xlsx", excel_0xd.getvalue())
                        zf.writestr(f"Autoline_Import_SALES_0XWG_{now_str}.xlsx", excel_0xwg.getvalue())
                    zip_dt_buffer.seek(0)
                    st.download_button(
                        label=f"📦 ดาวน์โหลดทั้ง 2 ไฟล์ (0XD + 0XWG ใน ZIP เดียว)",
                        data=zip_dt_buffer,
                        file_name=f"Autoline_Import_SALES_BY_DOC_TYPE_{now_str}.zip",
                        mime="application/zip",
                        key="dl_sales_dt_zip_btn"
                    )

            # OPTION 2: COMBINED ALL INVOICES
            elif dl_choice == "📦 รวมทั้งหมด (Combined All Invoices)":
                st.markdown("#### รวมทุกเอกสารทุกสาขาในไฟล์เดียว (Combined All Invoices)")
                st.info(
                    f"- จำนวนบิลทั้งหมด: **{stats_sales['total_invoices']:,}** ฉบับ (Batch 1 - {stats_sales['total_batches']:,})\n"
                    f"- จำนวนแถวทั้งหมด: **{len(rows_sales):,}** บรรทัด\n"
                    f"- ยอดเดบิตรวม: **{stats_sales['total_debit']:,.2f}** ฿\n"
                    f"- ต้นทุนรถยนต์รวม: **{stats_sales['total_vehicle_cost']:,.2f}** ฿ ({stats_sales['matched_cost_count']} คัน)"
                )
                excel_combined = generate_output_excel(template_path, rows_sales)
                st.download_button(
                    label=f"⬇️ ดาวน์โหลด Autoline_Import_SALES_ALL_COMBINED_{now_str}.xlsx ({stats_sales['total_invoices']:,} ฉบับ)",
                    data=excel_combined,
                    file_name=f"Autoline_Import_SALES_ALL_COMBINED_{now_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_sales_all_combined_btn",
                    use_container_width=True
                )

            # OPTION 3: SEPARATED BY BRANCH
            else:
                st.markdown("#### แยกไฟล์ตามรหัสสาขา (Branch 0001 & 0002)")
                if len(by_branch) > 1:
                    col_btns = st.columns(len(by_branch) + 1)
                    branch_excels = {}
                    for idx_b, (b_code, b_data) in enumerate(sorted(by_branch.items())):
                        b_name = "สำนักงานใหญ่" if b_code == "0001" else ("เกษตรนวมินทร์" if b_code == "0002" else f"สาขา {b_code}")
                        b_excel = generate_output_excel(template_path, b_data["rows"])
                        branch_excels[b_code] = b_excel
                        with col_btns[idx_b]:
                            st.download_button(
                                label=f"⬇️ โหลดไฟล์ สาขา {b_code}\n({b_name}: {b_data['total_invoices']:,} ฉบับ)",
                                data=b_excel,
                                file_name=f"Autoline_Import_SALES_BRANCH{b_code}_{now_str}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_branch_{b_code}",
                                use_container_width=True
                            )
                            
                    import zipfile
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for b_code, b_excel in branch_excels.items():
                            zf.writestr(f"Autoline_Import_SALES_BRANCH{b_code}_{now_str}.xlsx", b_excel.getvalue())
                    zip_buffer.seek(0)
                    
                    with col_btns[-1]:
                        st.download_button(
                            label=f"📦 ดาวน์โหลดทุกสาขา (ZIP)\n({len(by_branch)} ไฟล์สาขา)",
                            data=zip_buffer,
                            file_name=f"Autoline_Import_SALES_ALL_BRANCHES_{now_str}.zip",
                            mime="application/zip",
                            key="dl_sales_zip",
                            use_container_width=True
                        )
                else:
                    b_code = list(by_branch.keys())[0] if by_branch else "0001"
                    b_data = by_branch.get(b_code, {"rows": rows_sales, "total_invoices": stats_sales['total_invoices']})
                    b_name = "สำนักงานใหญ่" if b_code == "0001" else ("เกษตรนวมินทร์" if b_code == "0002" else f"สาขา {b_code}")
                    excel_sales_output = generate_output_excel(template_path, b_data["rows"])
                    st.download_button(
                        label=f"⬇️ ดาวน์โหลด Autoline_Import_SALES_BRANCH{b_code}.xlsx ({b_name}: {b_data['total_invoices']:,} ฉบับ)",
                        data=excel_sales_output,
                        file_name=f"Autoline_Import_SALES_BRANCH{b_code}_{now_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_branch_{b_code}_single",
                        use_container_width=True
                    )
                    
            # -----------------------------------------------------------------
            # 8. DATA PREVIEW TABLE WITH SEARCH & FILTERS
            # -----------------------------------------------------------------
            st.markdown("### 🔍 ตรวจสอบรายการเอกสาร (Sales Data Preview)")
            col_p1, col_p2, col_p3 = st.columns([2, 1, 1])
            with col_p1:
                search_sales_kw = st.text_input(
                    "ค้นหาเอกสารฝ่ายขาย (เลขที่บิล, ชื่อลูกค้า, เลขสต๊อก, VIN หรือหมวดหมู่)",
                    placeholder="เช่น 01WG26020001, N001960, 01D26010001, เงินจอง, ธนาคาร, X0003",
                    key="sales_search"
                )
            with col_p2:
                branch_options = ["ทุกสาขา (All Branches)"] + [
                    f"สาขา {b} ({'สำนักงานใหญ่' if b == '0001' else ('เกษตรนวมินทร์' if b == '0002' else b)})"
                    for b in sorted(by_branch.keys())
                ]
                sel_branch = st.selectbox("กรองตามสาขา", branch_options, key="sales_branch_filter")
            with col_p3:
                sel_group = st.selectbox(
                    "กรองตามกลุ่มบิล", 
                    ["ทั้งหมด (All Groups)", "กลุ่ม 0XD (เงินจอง/ป้าย/อุปกรณ์)", "กลุ่ม 0XWG (ขายรถยนต์)"], 
                    key="sales_group_filter"
                )
                
            display_sales_df = preview_sales.copy()
            if sel_branch != "ทุกสาขา (All Branches)":
                b_target = sel_branch.split()[1]
                display_sales_df = display_sales_df[display_sales_df["Branch"] == b_target]
                
            if sel_group.startswith("กลุ่ม 0XD"):
                display_sales_df = display_sales_df[display_sales_df["Invoice"].astype(str).str.contains(r'^\d{2}D', regex=True)]
            elif sel_group.startswith("กลุ่ม 0XWG"):
                display_sales_df = display_sales_df[display_sales_df["Invoice"].astype(str).str.contains(r'^\d{2}WG', regex=True)]
                
            if search_sales_kw.strip():
                skw = search_sales_kw.strip().lower()
                smask = (
                    display_sales_df["Invoice"].astype(str).str.lower().str.contains(skw, na=False) |
                    display_sales_df["Customer"].astype(str).str.lower().str.contains(skw, na=False) |
                    display_sales_df["Category"].astype(str).str.lower().str.contains(skw, na=False) |
                    display_sales_df["Subaccount"].astype(str).str.lower().str.contains(skw, na=False) |
                    display_sales_df["Stock No"].astype(str).str.lower().str.contains(skw, na=False) |
                    display_sales_df["VIN"].astype(str).str.lower().str.contains(skw, na=False)
                )
                display_sales_df = display_sales_df[smask]
                
            st.dataframe(display_sales_df, use_container_width=True, height=450)
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลฝ่ายขาย: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        # Clear feedback for missing mandatory files
        missing_m = []
        if not vat_file:
            missing_m.append("1. รายงานภาษีขาย (Sales VAT Report)")
        if not gl_file:
            missing_m.append("2. บัญชีแยกประเภททุกหมวด (All Category GL Report)")
            
        st.info(
            "💡 กรุณาอัปโหลด **ไฟล์จำเป็น (Mandatory Inputs)** ทั้ง 2 ไฟล์ด้านบนเพื่อเริ่มต้นแปลงข้อมูลฝ่ายขาย:\n" +
            "\n".join([f"- {m}" for m in missing_m])
        )

