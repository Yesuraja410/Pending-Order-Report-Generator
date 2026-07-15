# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd

project_dir = r"C:\Users\Yesuraja\.gemini\antigravity\scratch\Pending order report"
sys.path.append(project_dir)

from order_processor import process_and_validate_orders

def test_mode_1_gsheet_oms():
    print("Testing Mode 1: GSheet + OMS Report Alone...")
    # ORD-1: Pushed (no status mismatch)
    # ORD-2: Pushed (status mismatch)
    # ORD-3: Not Pushed to OMS
    df_pending = pd.DataFrame({
        "Order ID": ["ORD-1", "ORD-2", "ORD-3"],
        "Store Name": ["shopee-MY", "lazada-SG", "zalora-PH"],
        "SLA": ["2026-07-20", "2026-07-21", "2026-07-22"],
        "order_status": ["READY_TO_SHIP", "READY_TO_SHIP", "NEW"],
        "sla_status": ["future", "future", "breached"]
    })
    
    df_oms = pd.DataFrame({
        "Order ID": ["ORD-1", "ORD-2"],
        "order_status": ["READY_TO_SHIP", "SHIPPED"], # ORD-2 has mismatch
        "line_status": ["READY_TO_SHIP", "SHIPPED"]
    })

    res = process_and_validate_orders(
        pending_file=df_pending,
        tc_file=None,
        marketplace_file=None,
        oms_file=df_oms
    )

    enriched = res["enriched_pending_df"]
    disc = res["discrepancies_df"]
    summary = res["summary"]

    print("  Summary:", summary)
    print("  Enriched:\n", enriched[["Order ID", "OMS Order Status", "Final Remarks"]].to_string())
    print("  Discrepancies:\n", disc[["Order ID", "Validation Result", "Details"]].to_string())

    assert summary["total_pending_orders"] == 3
    assert summary["pushed_count"] == 2
    assert summary["not_pushed_count"] == 1
    assert summary["total_discrepancies"] == 2 # 1 mismatch + 1 not pushed
    print("Mode 1 passed!\n")

def test_mode_2_tc_marketplace():
    print("Testing Mode 2: TC + Marketplace Reconciliation...")
    # MP-1: Reflected
    # MP-2: Reflected
    # MP-3: Missing
    df_tc = pd.DataFrame({
        "Order ID": ["MP-1", "MP-2"]
    })
    
    df_marketplace = pd.DataFrame({
        "Order ID": ["MP-1", "MP-2", "MP-3"],
        "Store Name": ["shopee-MY", "lazada-SG", "zalora-PH"],
        "custom_sku": ["SKU-1", "SKU-2", "SKU-3"],
        "date": ["2026-07-15", "2026-07-15", "2026-07-15"]
    })

    res = process_and_validate_orders(
        pending_file=None,
        tc_file=df_tc,
        marketplace_file=df_marketplace,
        oms_file=None
    )

    enriched = res["enriched_pending_df"]
    disc = res["discrepancies_df"]
    summary = res["summary"]

    print("  Summary:", summary)
    print("  Enriched:\n", enriched[["Order ID", "Final Remarks"]].to_string())
    print("  Discrepancies:\n", disc[["Order ID", "Validation Result", "Details"]].to_string())

    assert summary["total_pending_orders"] == 3
    assert summary["pushed_count"] == 2 # reflected count
    assert summary["not_pushed_count"] == 1 # missing count
    assert summary["all_imported_to_tc"] is False
    assert len(disc) == 1
    print("Mode 2 passed!\n")

def test_mode_3_tc_oms():
    print("Testing Mode 3: TC + OMS Validation (Pending Extracted)...")
    # TC contains ORD-1 (READY TO SHIP - pending), ORD-2 (SHIPPED - not pending), ORD-3 (READY TO SHIP - pending)
    df_tc = pd.DataFrame({
        "Order ID": ["ORD-1", "ORD-2", "ORD-3"],
        "Store Name": ["shopee-MY", "lazada-SG", "zalora-PH"],
        "order_status": ["READY_TO_SHIP", "SHIPPED", "READY_TO_SHIP"],
        "Payment Status": ["completed", "completed", "completed"],
        "Payment Method": ["creditcard", "creditcard", "creditcard"],
        "time_to_ship_dead_line": ["2026-07-20", "2026-07-21", "2026-07-22"],
        "Custom SKU": ["SKU-1", "SKU-2", "SKU-3"]
    })

    # OMS contains ORD-1 (READY TO SHIP - matches), ORD-2 (SHIPPED - matches), ORD-3 (missing - not pushed)
    df_oms = pd.DataFrame({
        "Order ID": ["ORD-1", "ORD-2"],
        "order_status": ["READY_TO_SHIP", "SHIPPED"],
        "line_status": ["READY_TO_SHIP", "SHIPPED"],
        "ean": ["SKU-1", "SKU-2"]
    })

    res = process_and_validate_orders(
        pending_file=None,
        tc_file=df_tc,
        marketplace_file=None,
        oms_file=df_oms
    )

    enriched = res["enriched_pending_df"]
    disc = res["discrepancies_df"]
    summary = res["summary"]

    print("  Summary:", summary)
    print("  Enriched Pending list (extracted from TC):\n", enriched[["Order ID", "OMS Order Status", "Final Remarks"]].to_string())
    print("  Discrepancies:\n", disc[["Order ID", "Validation Result", "Details"]].to_string())

    assert len(enriched) == 2 # Only ORD-1 and ORD-3 are pending
    assert summary["pushed_count"] == 1 # ORD-1
    assert summary["not_pushed_count"] == 1 # ORD-3
    print("Mode 3 passed!\n")

if __name__ == "__main__":
    test_mode_1_gsheet_oms()
    test_mode_2_tc_marketplace()
    test_mode_3_tc_oms()
