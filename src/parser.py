import pandas as pd
import numpy as np

def clean_advisor_name(name):
    if pd.isna(name):
        return ""
    # Extract first name (before space)
    s = str(name).strip()
    parts = s.split()
    return parts[0] if parts else s

def load_and_validate_inputs(header_file, detail_file):
    """
    Reads Header and Detail Excel files, validates structure, and returns joined dataset.
    """
    # Read Header
    if isinstance(header_file, str):
        df_header = pd.read_excel(header_file)
    else:
        df_header = pd.read_excel(header_file)
    
    # Read Detail
    if isinstance(detail_file, str):
        df_detail = pd.read_excel(detail_file)
    else:
        df_detail = pd.read_excel(detail_file)
        
    # Standardize column names
    df_header.columns = [str(c).strip() for c in df_header.columns]
    df_detail.columns = [str(c).strip() for c in df_detail.columns]
    
    # Check key columns
    if "invoice_number" not in df_header.columns:
        raise ValueError("ไฟล์ Header ไม่มีคอลัมน์ 'invoice_number'")
    if "invoice_number" not in df_detail.columns:
        raise ValueError("ไฟล์ Detail ไม่มีคอลัมน์ 'invoice_number'")
        
    # Clean invoice_number
    df_header["invoice_number"] = df_header["invoice_number"].astype(str).str.strip()
    df_detail["invoice_number"] = df_detail["invoice_number"].astype(str).str.strip()
    
    # Drop rows without invoice_number
    df_header = df_header[df_header["invoice_number"] != ""].copy()
    df_detail = df_detail[df_detail["invoice_number"] != ""].copy()
    
    # Filter only relevant categories (P, L, S, N)
    df_detail["category"] = df_detail["category"].fillna("").astype(str).str.strip().str.upper()
    
    # Convert numerical columns
    for col in ["sales", "sales_tax", "nett_price", "discount_amount"]:
        if col in df_detail.columns:
            df_detail[col] = pd.to_numeric(df_detail[col], errors="coerce").fillna(0.0)
            
    for col in ["sales", "sales_tax", "total", "labor_amount", "parts_amount", "other_amount"]:
        if col in df_header.columns:
            df_header[col] = pd.to_numeric(df_header[col], errors="coerce").fillna(0.0)
            
    # Parse dates
    if "invoice_create_date" in df_header.columns:
        df_header["parsed_date"] = pd.to_datetime(df_header["invoice_create_date"], errors="coerce")
    else:
        df_header["parsed_date"] = pd.Timestamp.now()
        
    return df_header, df_detail
