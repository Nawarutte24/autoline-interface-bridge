import pandas as pd
import numpy as np

def clean_advisor_name(name):
    if pd.isna(name):
        return ""
    s = str(name).strip()
    parts = s.split()
    return parts[0] if parts else s

def load_and_validate_inputs(header_file, detail_file):
    """
    Reads Header and Detail Excel files, validates structure, and returns cleaned datasets.
    """
    df_header = pd.read_excel(header_file)
    df_detail = pd.read_excel(detail_file)
        
    # Standardize column names
    df_header.columns = [str(c).strip() for c in df_header.columns]
    df_detail.columns = [str(c).strip() for c in df_detail.columns]
    
    # Check key columns
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
    """
    Parses an uploaded Master Data Excel workbook (with sheets Customer_Subaccount, Branch_Mapping, GL_Mapping)
    """
    xl = pd.ExcelFile(master_file)
    master_dict = {
        "customers": {},
        "branches": {},
        "gl_rules": {}
    }
    
    # 1. Customer Subaccounts
    cust_sheet = None
    for s in xl.sheet_names:
        if "customer" in s.lower() or "subaccount" in s.lower():
            cust_sheet = s
            break
    if cust_sheet:
        df_c = pd.read_excel(master_file, sheet_name=cust_sheet)
        for _, r in df_c.iterrows():
            code = str(r.get("customer_code", "")).strip()
            tax_id = str(r.get("tax_ID", "")).strip()
            sub = str(r.get("autoline_subaccount", "")).strip()
            terms = r.get("terms", 30)
            if code and sub:
                master_dict["customers"][code] = {"subaccount": sub, "terms": terms}
            if tax_id and sub:
                master_dict["customers"][tax_id] = {"subaccount": sub, "terms": terms}
                
    # 2. Branch Mapping
    branch_sheet = None
    for s in xl.sheet_names:
        if "branch" in s.lower():
            branch_sheet = s
            break
    if branch_sheet:
        df_b = pd.read_excel(master_file, sheet_name=branch_sheet)
        for _, r in df_b.iterrows():
            prefix = str(r.get("dealer_prefix", "")).strip().upper()
            b_name = str(r.get("branch_name", "")).strip().lower()
            code = str(r.get("autoline_branch_code", "")).strip()
            if prefix and code:
                master_dict["branches"][prefix] = code
            if b_name and code:
                master_dict["branches"][b_name] = code
                
    # 3. GL Mapping
    gl_sheet = None
    for s in xl.sheet_names:
        if "gl" in s.lower():
            gl_sheet = s
            break
    if gl_sheet:
        df_gl = pd.read_excel(master_file, sheet_name=gl_sheet)
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
