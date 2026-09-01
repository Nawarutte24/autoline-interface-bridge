import pandas as pd
import datetime

def clean_first_name(advisor_name):
    if not advisor_name or pd.isna(advisor_name):
        return ""
    name_str = str(advisor_name).strip()
    parts = name_str.split()
    return parts[0] if parts else name_str

def transform_to_autoline_data(df_header, df_detail, config, master_dict=None):
    """
    Transforms Header and Detail DataFrames into Autoline AR/AP Journal rows.
    Supports Master Data lookups for Customers, Branches, and GL codes.
    """
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
    
    # Group detail by invoice_number
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
        advisor_first = clean_first_name(advisor_full)
        narrative = f"{advisor_first}_{inv_no}" if advisor_first else inv_no
        
        # Accounting figures from Header
        h_nett = float(h_row.get("nett_price", 0.0))
        h_tax = float(h_row.get("sales_tax", 0.0))
        h_total = float(h_row.get("total", 0.0))
        
        # Lookup Subaccount & Terms
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
            
        # Lookup Branch
        dealer_pfx = str(h_row.get("dealer_prefix", "")).strip().upper()
        branch_name_val = str(h_row.get("branch_name", "")).strip().lower()
        src_branch = config.get("src_branch", "0001")
        if dealer_pfx in branch_master:
            src_branch = branch_master[dealer_pfx]
        elif branch_name_val in branch_master:
            src_branch = branch_master[branch_name_val]
        
        # Get details for this invoice
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
                
        # Determine DOCSEQ
        if has_parts and not has_service:
            doc_seq = gl_master.get("P", {}).get("doc_seq") or config["parts"].get("doc_seq", "PINVOICE")
        else:
            doc_seq = gl_master.get("L", {}).get("doc_seq") or config["labor"].get("doc_seq", "SINVOICE")
            
        # -------------------------------------------------------------
        # 1. Document Header Row
        # -------------------------------------------------------------
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
        
        # -------------------------------------------------------------
        # 2. Blank Row between Header and Lines
        # -------------------------------------------------------------
        rows_to_write.append({"type": "BLANK"})
        
        # -------------------------------------------------------------
        # 3. Line 1: DEBIT Line (AR Control Account)
        # -------------------------------------------------------------
        ar_gl = gl_master.get("AR", {}).get("gl_code") or str(config.get("ar_gl_code", "11311001"))
        ar_dept = gl_master.get("AR", {}).get("department") or str(config.get("ar_department", "0000"))
        
        debit_record = {
            "type": "ITEM",
            "B": None,
            "C": None,
            "D": batch_idx,
            "E": None,
            "F": None,
            "G": config.get("currency", "THB"),
            "H": 1,
            "I": src_branch,
            "J": ar_dept,
            "K": int(ar_gl),
            "L": None,
            "M": h_total,
            "N": None,
            "O": narrative,
            "P": None,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "U": None,
            "V": None,
            "W": None,
            "X": None,
            "Y": None,
            "Z": None,
            "AA": None,
            "AB": None,
            "AC": None,
            "AD": None
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
        
        # -------------------------------------------------------------
        # 4. Line 2+: CREDIT Lines
        # -------------------------------------------------------------
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
                "B": None,
                "C": None,
                "D": batch_idx,
                "E": None,
                "F": None,
                "G": config.get("currency", "THB"),
                "H": line_num,
                "I": src_branch,
                "J": dept_val,
                "K": gl_code_val,
                "L": None,
                "M": None,
                "N": amount,
                "O": narrative,
                "P": None,
                "Q": tax_code_val,
                "R": None,
                "S": None,
                "T": None,
                "U": None,
                "V": None,
                "W": None,
                "X": aftstype_val,
                "Y": None,
                "Z": partfran,
                "AA": servprod,
                "AB": partprod,
                "AC": servicefranc,
                "AD": None
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
