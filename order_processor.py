# -*- coding: utf-8 -*-
# VERSION: v1 - Order Processor & SLA Validator
import pandas as pd
import numpy as np
import io

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
    # Strip any decimal part like .0 from order IDs
    s = _clean_str(val)
    if s.endswith(".0"):
        s = s[:-2]
    return s

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

def load_file_safely(file):
    """Load uploaded file object (CSV or Excel) into a DataFrame."""
    if file is None:
        return pd.DataFrame()
    name = file.name.lower()
    raw = file.read()
    file.seek(0)
    try:
        if name.endswith(".csv"):
            return pd.read_csv(io.BytesIO(raw), dtype=str)
        try:
            import python_calamine
            return pd.read_excel(io.BytesIO(raw), dtype=str, engine="calamine")
        except ImportError:
            return pd.read_excel(io.BytesIO(raw), dtype=str)
    except Exception as e:
        raise ValueError(f"Failed to read file {file.name}: {str(e)}")

def process_and_validate_orders(pending_file, tc_file, oms_file, contacts_file=None):
    """
    Process Pending Order Report, TC Report, and OMS Report.
    Returns:
      dict: {
         "enriched_pending_df": DataFrame,
         "discrepancies_df": DataFrame,
         "summary": dict,
         "seller_groups": dict { seller_name: DataFrame }
      }
    """
    # == 1. Load DataFrames ====================================================
    df_pending = load_file_safely(pending_file)
    df_tc = load_file_safely(tc_file)
    df_oms = load_file_safely(oms_file)
    df_contacts = load_file_safely(contacts_file) if contacts_file is not None else pd.DataFrame()

    if df_pending.empty:
        raise ValueError("Pending Order Report is empty or could not be loaded.")
    if df_tc.empty:
        raise ValueError("TC Report is empty or could not be loaded.")
    if df_oms.empty:
        raise ValueError("OMS Report is empty or could not be loaded.")

    # == 2. Standardize Columns ==============================================-
    # Pending Order columns
    pend_id_col = _find_column(df_pending, ["Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    pend_sla_col = _find_column(df_pending, ["SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
    pend_store_col = _find_column(df_pending, ["Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    
    # TC Report columns
    tc_id_col = _find_column(df_tc, ["Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    tc_status_col = _find_column(df_tc, ["TC Status", "Order Status", "Status", "TC_Status"])
    tc_sla_col = _find_column(df_tc, ["SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
    
    # OMS Report columns
    oms_id_col = _find_column(df_oms, ["Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    oms_status_col = _find_column(df_oms, ["OMS Status", "Order Status", "Status", "OMS_Status"])
    oms_pay_status_col = _find_column(df_oms, ["Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    oms_pay_method_col = _find_column(df_oms, ["Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])

    # Raise error if critical columns are missing
    if not pend_id_col:
        raise KeyError(f"Could not find 'Order ID' column in Pending Order Report. Available: {list(df_pending.columns)}")
    if not tc_id_col:
        raise KeyError(f"Could not find 'Order ID' column in TC Report. Available: {list(df_tc.columns)}")
    if not oms_id_col:
        raise KeyError(f"Could not find 'Order ID' column in OMS Report. Available: {list(df_oms.columns)}")

    # Clean Order IDs to ensure matches
    df_pending[pend_id_col] = df_pending[pend_id_col].apply(_clean_order_id)
    df_tc[tc_id_col] = df_tc[tc_id_col].apply(_clean_order_id)
    df_oms[oms_id_col] = df_oms[oms_id_col].apply(_clean_order_id)

    # == 3. SLA Enrichment ====================================================
    # Build TC SLA lookup map
    tc_sla_map = {}
    if tc_sla_col:
        tc_sla_map = df_tc.set_index(tc_id_col)[tc_sla_col].dropna().to_dict()

    enriched_rows = []
    enriched_sla_count = 0
    blank_sla_not_found = 0

    # Ensure SLA column exists in pending
    target_sla_col = pend_sla_col if pend_sla_col else "SLA"
    if target_sla_col not in df_pending.columns:
        df_pending[target_sla_col] = ""

    # Ensure Store Name exists in pending
    target_store_col = pend_store_col if pend_store_col else "Store Name"
    if target_store_col not in df_pending.columns:
        df_pending[target_store_col] = "Default Store"

    df_pending["SLA Source"] = "Pending Report"

    for idx, row in df_pending.iterrows():
        order_id = row[pend_id_col]
        sla_val = _clean_str(row[target_sla_col])
        
        if not sla_val: # Blank SLA
            tc_sla_val = _clean_str(tc_sla_map.get(order_id, ""))
            if tc_sla_val:
                df_pending.at[idx, target_sla_col] = tc_sla_val
                df_pending.at[idx, "SLA Source"] = "Enriched from TC"
                enriched_sla_count += 1
            else:
                df_pending.at[idx, "SLA Source"] = "Missing (Not in TC)"
                blank_sla_not_found += 1

    # == 4. OMS vs TC Validations ============================================-
    # Create indexed lookups for validations
    tc_lookup = {}
    for _, row in df_tc.iterrows():
        oid = row[tc_id_col]
        tc_lookup[oid] = {
            "Status": _clean_str(row[tc_status_col]) if tc_status_col else "",
            "Row": row.to_dict()
        }

    oms_lookup = {}
    for _, row in df_oms.iterrows():
        oid = row[oms_id_col]
        oms_lookup[oid] = {
            "Status": _clean_str(row[oms_status_col]) if oms_status_col else "",
            "Payment Status": _clean_str(row[oms_pay_status_col]) if oms_pay_status_col else "",
            "Payment Method": _clean_str(row[oms_pay_method_col]) if oms_pay_method_col else "",
            "Row": row.to_dict()
        }

    # All unique order IDs across TC and OMS
    all_order_ids = set(tc_lookup.keys()) | set(oms_lookup.keys())
    discrepancies = []

    for oid in all_order_ids:
        in_tc = oid in tc_lookup
        in_oms = oid in oms_lookup
        
        tc_status = tc_lookup[oid]["Status"] if in_tc else ""
        oms_status = oms_lookup[oid]["Status"] if in_oms else ""
        oms_pay_status = oms_lookup[oid]["Payment Status"] if in_oms else ""
        oms_pay_method = oms_lookup[oid]["Payment Method"] if in_oms else ""

        # Check Payment method: is it COD?
        is_cod = any(term in oms_pay_method.lower() for term in ["cod", "cash on delivery", "cashondelivery"]) if in_oms else False
        is_pay_pending = False

        # Let's also look for payment status and payment method in TC if not in OMS
        tc_pay_status_col = _find_column(df_tc, ["Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
        tc_pay_method_col = _find_column(df_tc, ["Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])
        
        if in_tc:
            tc_pay_status = _clean_str(tc_lookup[oid]["Row"].get(tc_pay_status_col, "")) if tc_pay_status_col else ""
            tc_pay_method = _clean_str(tc_lookup[oid]["Row"].get(tc_pay_method_col, "")) if tc_pay_method_col else ""
            
            # If we don't have it from OMS, use TC's values
            if not oms_pay_status:
                oms_pay_status = tc_pay_status
            if not oms_pay_method:
                oms_pay_method = tc_pay_method
                is_cod = any(term in oms_pay_method.lower() for term in ["cod", "cash on delivery", "cashondelivery"])

        # == Rule 3: Payment Pending check ==============================
        if in_tc and tc_status.lower() == "new" and oms_pay_status.lower() == "pending" and not is_cod:
            is_pay_pending = True
            
            if not in_oms:
                discrepancies.append({
                    "Order ID": oid,
                    "Validation Result": "Payment Pending (Hold)",
                    "TC Status": tc_status,
                    "OMS Status": "Not in OMS",
                    "Payment Status": oms_pay_status,
                    "Payment Method": oms_pay_method if oms_pay_method else "Non-COD",
                    "Details": "Expected hold: TC status is New, Payment Status is Pending, and Payment Method is not COD."
                })
                continue
            else:
                discrepancies.append({
                    "Order ID": oid,
                    "Validation Result": "Payment Pending Push Discrepancy",
                    "TC Status": tc_status,
                    "OMS Status": oms_status,
                    "Payment Status": oms_pay_status,
                    "Payment Method": oms_pay_method,
                    "Details": "Order pushed to OMS despite TC being New, Payment Status Pending, and Method not COD."
                })

        # == Pushed Status Validation (Missing order check) ==============
        if in_tc and not in_oms and not is_pay_pending:
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "Missing in OMS",
                "TC Status": tc_status,
                "OMS Status": "Missing",
                "Payment Status": oms_pay_status,
                "Payment Method": oms_pay_method,
                "Details": "Order exists in TC but was not successfully pushed to OMS."
            })
            continue

        if in_oms and not in_tc:
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "Missing in TC",
                "TC Status": "Missing",
                "OMS Status": oms_status,
                "Payment Status": oms_pay_status,
                "Payment Method": oms_pay_method,
                "Details": "Order exists in OMS but does not exist in TC."
            })
            continue

        # == Rule 1: Cancelled status check ==============================
        is_tc_cancelled = (tc_status.lower() == "cancelled")
        is_oms_cancelled = (oms_status.lower() == "cancelled")
        
        if is_tc_cancelled != is_oms_cancelled:
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "Cancelled Sync Mismatch",
                "TC Status": tc_status,
                "OMS Status": oms_status,
                "Payment Status": oms_pay_status,
                "Payment Method": oms_pay_method,
                "Details": f"Cancellation mismatch: TC is {tc_status}, OMS is {oms_status}."
            })

        # == Rule 2: Packed status check ================================-
        if oms_status.lower() == "packed" and tc_status.lower() == "new":
            discrepancies.append({
                "Order ID": oid,
                "Validation Result": "OMS Packed but TC New",
                "TC Status": tc_status,
                "OMS Status": oms_status,
                "Payment Status": oms_pay_status,
                "Payment Method": oms_pay_method,
                "Details": "Discrepancy: OMS status is Packed, but TC status is New (must be Accepted, Picked, or Ready to Ship)."
            })

    df_discrepancies = pd.DataFrame(discrepancies) if discrepancies else pd.DataFrame(columns=[
        "Order ID", "Validation Result", "TC Status", "OMS Status", "Payment Status", "Payment Method", "Details"
    ])

    # == 5. Seller Contact Map & Grouping ====================================-
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

    # Summary metrics
    summary = {
        "total_pending_orders": len(df_pending),
        "enriched_sla_count": enriched_sla_count,
        "blank_sla_not_found": blank_sla_not_found,
        "total_discrepancies": len(df_discrepancies),
        "cancelled_mismatches": int((df_discrepancies["Validation Result"] == "Cancelled Sync Mismatch").sum()),
        "packed_mismatches": int((df_discrepancies["Validation Result"] == "OMS Packed but TC New").sum()),
        "payment_pending_holds": int((df_discrepancies["Validation Result"] == "Payment Pending (Hold)").sum()),
        "missing_in_oms": int((df_discrepancies["Validation Result"] == "Missing in OMS").sum()),
        "total_sellers": len(seller_groups)
    }

    return {
        "enriched_pending_df": df_pending,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": target_sla_col
    }
