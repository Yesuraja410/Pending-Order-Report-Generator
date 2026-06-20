# -*- coding: utf-8 -*-
# VERSION: v5 - Google Sheets & Excel Styling Update
import pandas as pd
import numpy as np
import io
import re
import urllib.request
import urllib.error
from datetime import datetime

def get_google_sheet_download_url(url):
    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        spreadsheet_id = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    return None

def download_google_sheet(url):
    download_url = get_google_sheet_download_url(url)
    if not download_url:
        download_url = url
    try:
        req = urllib.request.Request(
            download_url,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req) as response:
            data = response.read()
        return io.BytesIO(data)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise ValueError(
                "Failed to download Google Sheet: Access Denied (HTTP 401/403 Unauthorized). "
                "Please ensure that the sharing settings of your Google Sheet are set to "
                "'Anyone with the link can view' or 'Anyone with the link can edit' (restricted sheets require auth, which is not supported)."
            )
        raise ValueError(f"Failed to download Google Sheet: HTTP Error {e.code} - {e.reason}")
    except Exception as e:
        raise ValueError(f"Failed to download Google Sheet from URL: {str(e)}")

def compute_sla_status(sla_date, ref_date):
    if _is_blank(sla_date) or _is_blank(ref_date):
        return ""
    s_dt = extract_date(sla_date)
    r_dt = extract_date(ref_date)
    if not (len(s_dt) == 10 and len(r_dt) == 10):
        return ""
    if s_dt < r_dt:
        return "Breached"
    elif s_dt == r_dt:
        return "Today"
    else:
        return "Future"


def _clean_str(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()

def _clean_order_id(val):
    """
    Standardizes the order number as a string. Handles potential float scientific 
    notations introduced by Excel for numbers larger than 10 digits.
    """
    s = _clean_str(val)
    if not s or s.lower() in ("nan", "none", "nat"):
        return ""
    
    # Check if the string was converted to scientific notation (e.g. 2.3E+14)
    if "e" in s.lower() and "+" in s:
        try:
            f_val = float(s)
            s = str(int(f_val))
        except Exception:
            pass
            
    # Remove float decimal point representation
    if s.endswith(".0"):
        s = s[:-2]
        
    return s

def parse_country_and_channel(nickname):
    nick = str(nickname).strip().upper()
    # Extract Country
    if nick.endswith("SG") or "-SG" in nick or "_SG" in nick:
        country = "SG"
    elif nick.endswith("MY") or "-MY" in nick or "_MY" in nick:
        country = "MY"
    elif nick.endswith("PH") or "-PH" in nick or "_PH" in nick:
        country = "PH"
    else:
        # Default fallback checks
        if "SG" in nick:
            country = "SG"
        elif "MY" in nick:
            country = "MY"
        elif "PH" in nick:
            country = "PH"
        else:
            country = "UNKNOWN"
            
    # Extract Channel
    nick_lower = nick.lower()
    if "shopee" in nick_lower:
        channel = "Shopee"
    elif "lazada" in nick_lower:
        channel = "Lazada"
    elif "zalora" in nick_lower:
        channel = "Zalora"
    elif "tiktok" in nick_lower:
        channel = "TikTok"
    else:
        channel = "Other"
        
    return country, channel

def _find_column(df, candidates):
    """Find a column in df that matches any of the candidate names case-insensitively and ignoring underscores/spaces."""
    cols = list(df.columns)
    for cand in candidates:
        cand_norm = cand.lower().replace(" ", "").replace("_", "").replace("-", "")
        for col in cols:
            col_norm = str(col).lower().replace(" ", "").replace("_", "").replace("-", "")
            if col_norm == cand_norm:
                return col
    return None

def _is_blank(val):
    """Check if a value is null, empty string, or nan/nat strings produced by Excel loading."""
    s = str(val).strip().replace('\r', '').replace('\n', '')
    return not s or s.lower() in ("nan", "none", "nat", "null", "undefined", "nat")

def extract_date(val):
    """Extract a YYYY-MM-DD date from various timestamp formats."""
    s = str(val).strip()
    if len(s) >= 10:
        if '-' in s:
            parts = s.split(' ')[0].split('-')
            if len(parts) == 3:
                if len(parts[0]) == 4:
                    return f"{parts[0]}-{parts[1]}-{parts[2]}"
                else:
                    return f"{parts[2]}-{parts[1]}-{parts[0]}"
        elif '/' in s:
            parts = s.split(' ')[0].split('/')
            if len(parts) == 3:
                if len(parts[2]) == 4:
                    return f"{parts[2]}-{parts[1]}-{parts[0]}"
                else:
                    return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return s

def load_file_safely(file):
    """Load uploaded file object (CSV or Excel) into a DataFrame. Scans sheets if Excel."""
    if file is None:
        return pd.DataFrame()
    
    try:
        file.seek(0)
    except Exception:
        pass
        
    name = getattr(file, "name", "google_sheet.xlsx").lower()
    
    try:
        if name.endswith(".csv"):
            try:
                file.seek(0)
                df = pd.read_csv(file, dtype=str)
                if not df.empty:
                    return df.dropna(how="all").reset_index(drop=True)
            except Exception:
                pass
            
            # Fallback to reading bytes
            file.seek(0)
            raw = file.read()
            df = pd.read_csv(io.BytesIO(raw), dtype=str)
            return df.dropna(how="all").reset_index(drop=True)
        else:
            # Excel - Scan all sheets to find the first non-empty one
            try:
                file.seek(0)
                xl = pd.ExcelFile(file)
            except Exception:
                file.seek(0)
                raw = file.read()
                xl = pd.ExcelFile(io.BytesIO(raw))
                
            for sheet in xl.sheet_names:
                try:
                    df = xl.parse(sheet, dtype=str)
                    if df is not None and not df.empty:
                        df_clean = df.dropna(how="all").reset_index(drop=True)
                        if not df_clean.empty:
                            return df_clean
                except Exception:
                    continue
            
            # Fallback: if all sheets are empty, return the first sheet
            if xl.sheet_names:
                return xl.parse(xl.sheet_names[0], dtype=str)
            return pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Failed to read file {getattr(file, 'name', 'Google Sheet')}: {str(e)}")

def process_and_validate_orders(pending_file, tc_file, oms_file, contacts_file=None):
    """
    Process Pending Order Report (SLA file), TC Report (All file), and OMS Report (Sales Order file).
    """
    # == 1. Load DataFrames ====================================================
    if isinstance(pending_file, str) and (pending_file.startswith("http://") or pending_file.startswith("https://")):
        pending_file = download_google_sheet(pending_file)
    df_pending = load_file_safely(pending_file)
    df_tc = load_file_safely(tc_file)
    df_oms = load_file_safely(oms_file)
    df_contacts = load_file_safely(contacts_file) if contacts_file is not None else pd.DataFrame()

    if df_pending.empty:
        raise ValueError("Pending Order Report (SLA Report) is empty or could not be loaded.")
    if df_tc.empty:
        raise ValueError("TC Report (All file) is empty or could not be loaded.")
    if df_oms.empty:
        raise ValueError("OMS Report (Sales Order file) is empty or could not be loaded.")

    # Find store column in SLA Report first to ignore TikTok PH
    pend_store_col = _find_column(df_pending, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    if pend_store_col and pend_store_col in df_pending.columns:
        # Filter out TikTok PH orders
        df_pending = df_pending[df_pending[pend_store_col].astype(str).str.strip().str.lower() != 'tiktok-ph'].copy()

    if df_pending.empty:
        raise ValueError("Pending Order Report (SLA Report) has no rows after filtering out TikTok PH.")

    # == 2. Standardize Columns ===============================================
    # Pending Order columns
    pend_id_col = _find_column(df_pending, ["order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    pend_sla_col = _find_column(df_pending, ["mp_sla_date", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date", "mp_sla_date_updated"])
    
    # TC Report columns (All file)
    tc_id_col = _find_column(df_tc, ["order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    tc_num_col = _find_column(df_tc, ["order_number", "order_id"]) # specifically find order_number column
    tc_status_col = _find_column(df_tc, ["order_status", "TC Status", "Order Status", "Status", "TC_Status"])
    
    # SLA Lookup targets time_to_ship_dead_line as primary, then fallback
    tc_sla_col = _find_column(df_tc, ["time_to_ship_dead_line", "order_sla", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
    
    tc_pay_status_col = _find_column(df_tc, ["payment_status", "Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    tc_pay_method_col = _find_column(df_tc, ["payment_methods", "Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])
    
    # OMS Report columns (Sales Order file)
    oms_id_col = _find_column(df_oms, ["order_no", "order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    oms_status_col = _find_column(df_oms, ["order_status", "OMS Status", "Order Status", "Status", "OMS_Status"])
    oms_pay_status_col = _find_column(df_oms, ["Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    oms_pay_method_col = _find_column(df_oms, ["Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])

    # Raise error if critical columns are missing
    if not pend_id_col:
        raise KeyError(f"Could not find 'Order ID' column in Pending Order Report. Available: {list(df_pending.columns)}")
    if not tc_id_col:
        raise KeyError(f"Could not find 'Order ID' column in TC Report (All file). Available: {list(df_tc.columns)}")
    if not oms_id_col:
        raise KeyError(f"Could not find 'Order ID' column in OMS Report (Sales Order file). Available: {list(df_oms.columns)}")

    # Clean Order IDs to ensure matches (retaining large string values correctly)
    df_pending[pend_id_col] = df_pending[pend_id_col].apply(_clean_order_id)
    df_tc[tc_id_col] = df_tc[tc_id_col].apply(_clean_order_id)
    df_oms[oms_id_col] = df_oms[oms_id_col].apply(_clean_order_id)
    if tc_num_col:
        df_tc[tc_num_col] = df_tc[tc_num_col].apply(_clean_order_id)

    # == 3. SLA Enrichment & Pushed Status ====================================
    # Build TC mappings
    tc_sla_map = {}
    if tc_sla_col:
        tc_sla_map = df_tc.set_index(tc_id_col)[tc_sla_col].dropna().to_dict()

    tc_payment_status = {}
    if tc_pay_status_col:
        tc_payment_status = df_tc.set_index(tc_id_col)[tc_pay_status_col].dropna().to_dict()

    tc_payment_method = {}
    if tc_pay_method_col:
        tc_payment_method = df_tc.set_index(tc_id_col)[tc_pay_method_col].dropna().to_dict()

    # Build bidirectional ID to package-number mappings from TC report (crucial for Zalora package lookup)
    tc_id_to_num = {}
    tc_num_to_id = {}
    if tc_id_col and tc_num_col and tc_id_col != tc_num_col:
        for _, row in df_tc.iterrows():
            oid = row[tc_id_col]
            onum = row[tc_num_col]
            if oid and onum:
                tc_id_to_num[oid] = onum
                tc_num_to_id[onum] = oid

    # Build OMS status lookup maps
    oms_status_map = df_oms.set_index(oms_id_col)[oms_status_col].dropna().to_dict() if oms_status_col else {}
    oms_pay_status_map = df_oms.set_index(oms_id_col)[oms_pay_status_col].dropna().to_dict() if oms_pay_status_col else {}
    oms_pay_method_map = df_oms.set_index(oms_id_col)[oms_pay_method_col].dropna().to_dict() if oms_pay_method_col else {}

    # Ensure SLA column exists in pending
    target_sla_col = pend_sla_col if pend_sla_col else "SLA"
    if target_sla_col not in df_pending.columns:
        df_pending[target_sla_col] = ""

    # Ensure Store Name exists in pending
    target_store_col = pend_store_col if pend_store_col else "Store Name"
    if target_store_col not in df_pending.columns:
        df_pending[target_store_col] = "Default Store"

    # Add output columns to Pending Order Report
    df_pending["Correct Order Number"] = df_pending[pend_id_col]
    df_pending["SLA Source"] = "Pending Report"
    df_pending["OMS Order Status"] = ""
    df_pending["Final Remarks"] = ""

    # Determine reference date from "Today" rows or data
    ref_date = None
    if "sla_status" in df_pending.columns:
        today_rows = df_pending[df_pending["sla_status"].astype(str).str.strip().str.lower() == "today"]
        if not today_rows.empty:
            for val in today_rows[target_sla_col]:
                if not _is_blank(val):
                    p_dt = extract_date(val)
                    if len(p_dt) == 10:
                        ref_date = p_dt
                        break
    if not ref_date:
        if target_sla_col in df_pending.columns:
            for val in df_pending[target_sla_col]:
                if not _is_blank(val):
                    p_dt = extract_date(val)
                    if len(p_dt) == 10:
                        ref_date = p_dt
                        break
    if not ref_date:
        ref_date = datetime.today().strftime('%Y-%m-%d')

    # Ensure sla_status column exists in df_pending
    if "sla_status" not in df_pending.columns:
        df_pending["sla_status"] = ""

    enriched_sla_count = 0
    blank_sla_not_found = 0
    pushed_count = 0
    not_pushed_count = 0
    unpaid_count = 0

    for idx, row in df_pending.iterrows():
        order_id = row[pend_id_col]
        sla_val = row[target_sla_col]
        
        # ── SLA Check ──
        if _is_blank(sla_val): # Blank SLA (recognizing "NaT", "nan" loaded via dtype=str)
            tc_sla_val = _clean_str(tc_sla_map.get(order_id, ""))
            if not _is_blank(tc_sla_val):
                df_pending.at[idx, target_sla_col] = tc_sla_val
                df_pending.at[idx, "SLA Source"] = "Enriched from TC"
                enriched_sla_count += 1
            else:
                df_pending.at[idx, "SLA Source"] = "Missing (Not in TC)"
                blank_sla_not_found += 1
        else:
            df_pending.at[idx, "SLA Source"] = "Pending Report"

        # Update/Compute sla_status for blanks/enriched or if not present
        sla_status_val = row.get("sla_status", "")
        if _is_blank(sla_status_val):
            curr_sla = df_pending.at[idx, target_sla_col]
            calculated_status = compute_sla_status(curr_sla, ref_date)
            if calculated_status:
                df_pending.at[idx, "sla_status"] = calculated_status

        # ── OMS Status & Final Remarks Check (Checking both SLA ID and Mapped TC Order Number) ──
        tc_mapped_num = tc_id_to_num.get(order_id, "")
        
        is_in_oms = False
        oms_stat = ""
        
        if order_id in oms_status_map:
            is_in_oms = True
            oms_stat = oms_status_map[order_id]
        elif tc_mapped_num and tc_mapped_num in oms_status_map:
            is_in_oms = True
            oms_stat = oms_status_map[tc_mapped_num]
            
        if is_in_oms:
            df_pending.at[idx, "OMS Order Status"] = oms_stat
            df_pending.at[idx, "Final Remarks"] = "Successfully Pushed to OMS"
            pushed_count += 1
        else:
            df_pending.at[idx, "OMS Order Status"] = "Not in OMS"
            
            # Retrieve payment status & method from TC (All file) as fallback
            pay_status = _clean_str(tc_payment_status.get(order_id, ""))
            pay_method = _clean_str(tc_payment_method.get(order_id, ""))
            
            is_cod = any(term in pay_method.lower() for term in ["cod", "cash on delivery", "cashondelivery"])
            is_pending = (pay_status.lower() in ("pending", "unpaid", "awaiting"))
            is_completed = (pay_status.lower() in ("completed", "paid", "success", "complete", "fully_paid"))
            
            if is_pending and not is_cod:
                df_pending.at[idx, "Final Remarks"] = "Unpaid Orders"
                unpaid_count += 1
            elif (is_pending and is_cod) or is_completed or not pay_status:
                df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS"
                not_pushed_count += 1
            else:
                df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS"
                not_pushed_count += 1

    # == 4. Status Discrepancy Validations ====================================
    tc_lookup = {}
    for _, row in df_tc.iterrows():
        oid = row[tc_id_col]
        tc_lookup[oid] = {
            "Status": _clean_str(row[tc_status_col]) if tc_status_col else "",
            "Row": row.to_dict()
        }

    oms_lookup = {}
    for _, row in df_oms.iterrows():
        oid_raw = row[oms_id_col]
        # Map OMS package number back to TC order_id to align keys correctly for Zalora
        oid = tc_num_to_id.get(oid_raw, oid_raw)
        oms_lookup[oid] = {
            "Status": _clean_str(row[oms_status_col]) if oms_status_col else "",
            "Row": row.to_dict()
        }

    # Compare status logic
    all_order_ids = set(tc_lookup.keys()) & set(oms_lookup.keys()) # intersect keys
    discrepancies = []

    for oid in all_order_ids:
        tc_status = tc_lookup[oid]["Status"]
        oms_status = oms_lookup[oid]["Status"]

        # Rule 1: Cancelled status check
        is_tc_cancelled = tc_status.lower() in ("cancelled", "canceled")
        is_oms_cancelled = oms_status.lower() in ("cancelled", "canceled")
        
        # Exception: TC status is Cancelled and OMS is Returned/Return -> Ignore mismatch!
        is_tc_cancelled_and_oms_returned = is_tc_cancelled and (oms_status.lower() in ("returned", "return"))
        
        if is_tc_cancelled != is_oms_cancelled and not is_tc_cancelled_and_oms_returned:
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "Cancelled Sync Mismatch",
                "TC Status (All file)": tc_status,
                "OMS Status (Sales Order file)": oms_status,
                "Details": f"Cancellation mismatch: TC is {tc_status}, OMS is {oms_status}."
            })

        # Rule 2: Packed status check
        if oms_status.lower() == "packed" and tc_status.lower() == "new":
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "OMS Packed but TC New",
                "TC Status (All file)": tc_status,
                "OMS Status (Sales Order file)": oms_status,
                "Details": "Discrepancy: OMS status is Packed, but TC status is New (must be Accepted, Picked, or Ready to Ship)."
            })

        # Rule 3: Shipped status check
        if oms_status.lower() == "shipped" and tc_status.lower() != "shipped":
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "OMS Shipped but TC not Shipped",
                "TC Status (All file)": tc_status,
                "OMS Status (Sales Order file)": oms_status,
                "Details": f"Discrepancy: OMS status is Shipped, but TC status is '{tc_status}' (must be Shipped)."
            })

    df_discrepancies = pd.DataFrame(discrepancies) if discrepancies else pd.DataFrame(columns=[
        "Order ID", "Validation Result", "TC Status (All file)", "OMS Status (Sales Order file)", "Details"
    ])

    # == 5. Seller Contact Map & Grouping =====================================
    email_map = {}
    if not df_contacts.empty:
        c_store = _find_column(df_contacts, ["Store Name", "Store", "Seller", "Shop Name", "Shop"])
        c_email = _find_column(df_contacts, ["Seller Email", "Email", "SellerEmail", "Email Address"])
        if c_store and c_email:
            for _, row in df_contacts.iterrows():
                store_key = _clean_str(row[c_store]).lower()
                email_val = _clean_str(row[c_email])
                if store_key and email_val:
                    email_map[store_key] = email_val

    # Group enriched pending orders by Seller
    seller_groups = {}
    stores = df_pending[target_store_col].unique()

    for store in stores:
        store_clean = _clean_str(store)
        store_key = store_clean.lower()
        
        store_df = df_pending[df_pending[target_store_col] == store].copy()
        
        mapped_email = email_map.get(store_key, "")
        store_df["Seller Email"] = mapped_email
        
        seller_groups[store_clean] = {
            "df": store_df,
            "email": mapped_email
        }

    # == 6. Country-specific datasets & Pivot Tables ==========================
    country_reports = {}
    
    # Pre-calculate clean SLA Date for columns
    date_col = target_sla_col
    if date_col and date_col in df_pending.columns:
        df_pending["Order Date"] = df_pending[date_col].apply(extract_date)
    else:
        df_pending["Order Date"] = "Unknown"
    
    for country in ["SG", "MY", "PH"]:
        c_rows = []
        for idx, row in df_pending.iterrows():
            store_val = _clean_str(row[target_store_col])
            c_code, chan = parse_country_and_channel(store_val)
            if c_code == country:
                # Apply channel filter requirements:
                if country == "SG" and chan in ["Lazada", "Shopee", "Zalora"]:
                    row_dict = row.to_dict()
                    row_dict["Country"] = country
                    row_dict["Channel"] = f"{chan} {country}"
                    c_rows.append(row_dict)
                elif country == "MY" and chan in ["Lazada", "Shopee", "Zalora", "TikTok"]:
                    row_dict = row.to_dict()
                    row_dict["Country"] = country
                    row_dict["Channel"] = f"{chan} {country}"
                    c_rows.append(row_dict)
                elif country == "PH" and chan in ["Lazada", "Shopee", "Zalora"]:
                    row_dict = row.to_dict()
                    row_dict["Country"] = country
                    row_dict["Channel"] = f"{chan} {country}"
                    c_rows.append(row_dict)
                    
        country_df = pd.DataFrame(c_rows)
        if not country_df.empty:
            # Pivot table: Channel & OMS status in Rows, Order date in Columns, Order number (count) in Values
            pivot_df = country_df.pivot_table(
                index=["Channel", "OMS Order Status"],
                columns="Order Date",
                values="Correct Order Number",
                aggfunc="count",
                fill_value=0
            )
            
            # Format columns of Pivot Table as DD-MM-YYYY
            new_cols = []
            for col in pivot_df.columns:
                try:
                    dt = pd.to_datetime(col)
                    new_cols.append(dt.strftime('%d-%m-%Y'))
                except Exception:
                    new_cols.append(col)
            pivot_df.columns = new_cols
            
            # Add Grand Total column
            pivot_df["Grand Total"] = pivot_df.sum(axis=1)
            # Add Grand Total row
            pivot_df.loc[("Grand Total", ""), :] = pivot_df.sum(axis=0)
            pivot_df = pivot_df.reset_index()
            
            # Highlight Summary metrics
            summary_metrics = [
                {"Metric": "Overdue (SLA breached)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "breached").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Handover today (Today SLA)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "today").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Order Status at New", "Count": int((country_df["OMS Order Status"].astype(str).str.strip().str.lower() == "new").sum())},
                {"Metric": "Within SLA (Future)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "future").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Not reflecting in OM", "Count": int((country_df["OMS Order Status"] == "Not in OMS").sum())}
            ]
            summary_df = pd.DataFrame(summary_metrics)
            
            country_reports[country] = {
                "raw_df": country_df,
                "pivot_df": pivot_df,
                "summary_df": summary_df
            }
        else:
            country_reports[country] = {
                "raw_df": pd.DataFrame(),
                "pivot_df": pd.DataFrame(),
                "summary_df": pd.DataFrame()
            }

    # Summary metrics
    summary = {
        "total_pending_orders": len(df_pending),
        "enriched_sla_count": enriched_sla_count,
        "blank_sla_not_found": blank_sla_not_found,
        "total_discrepancies": len(df_discrepancies),
        "cancelled_mismatches": int((df_discrepancies["Validation Result"] == "Cancelled Sync Mismatch").sum()),
        "packed_mismatches": int((df_discrepancies["Validation Result"] == "OMS Packed but TC New").sum()),
        "pushed_count": pushed_count,
        "not_pushed_count": not_pushed_count,
        "unpaid_count": unpaid_count,
        "total_sellers": len(seller_groups)
    }

    ref_date_dmy = ""
    try:
        ref_dt = datetime.strptime(ref_date, "%Y-%m-%d")
        ref_date_dmy = ref_dt.strftime("%d-%m-%Y")
    except Exception:
        ref_date_dmy = ref_date

    return {
        "enriched_pending_df": df_pending,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": target_sla_col,
        "country_reports": country_reports,
        "ref_date_dmy": ref_date_dmy
    }
