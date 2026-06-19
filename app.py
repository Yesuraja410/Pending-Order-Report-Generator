# -*- coding: utf-8 -*-
import io
import os
import gc
import tempfile
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="Status Validation Analyzer",
    page_icon="chart_with_upwards_trend",
    layout="wide",
)

from file_loaders import load_all_files
from validators import run_sku_validation, run_pid_validation
from report_generator import generate_status_report
from styles import inject_css
from order_processor import process_and_validate_orders
from email_sender import test_smtp_connection, send_seller_report_email

inject_css()

# == Persistent report directory ==============================================-
REPORT_DIR = os.path.join(tempfile.gettempdir(), "svr_reports")
os.makedirs(REPORT_DIR, exist_ok=True)


def _make_filename(data, country):
    channel_map = {
        "lazada": "Lazada",
        "shopee": "Shopee",
        "zalora": "Zalora",
        "tiktok": "TikTok",
    }
    channels = []
    for key, label in channel_map.items():
        df = data.get(key, pd.DataFrame())
        if df is not None and not df.empty:
            channels.append(label)
    today = datetime.today().strftime("%Y-%m-%d")
    if channels:
        return "_".join(channels) + "_" + country + "_Status_Validation_Report_" + today + ".xlsx"
    return "Status_Validation_Report_" + country + "_" + today + ".xlsx"


def _write_report(sheets, fname):
    """Write report to persistent directory. Returns file path."""
    fpath = os.path.join(REPORT_DIR, fname)
    with pd.ExcelWriter(fpath, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    return fpath


def _write_qc_report(sheets, fname):
    """Write QC audit report to persistent directory. Returns file path."""
    fpath = os.path.join(REPORT_DIR, fname)
    with pd.ExcelWriter(fpath, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return fpath


def _list_saved_reports():
    """List all saved report files sorted newest first."""
    files = []
    for f in os.listdir(REPORT_DIR):
        if f.endswith(".xlsx"):
            fpath = os.path.join(REPORT_DIR, f)
            mtime = os.path.getmtime(fpath)
            size_mb = round(os.path.getsize(fpath) / (1024 * 1024), 2)
            files.append((f, fpath, mtime, size_mb))
    files.sort(key=lambda x: x[2], reverse=True)
    return files


def show_df(df, max_rows=500):
    if df is None or df.empty:
        st.warning("No data to display.")
        return
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Rows", len(df))
    if "Final Status" in df.columns:
        c2.metric("Active",   int((df["Final Status"] == "Active").sum()))
        c3.metric("Inactive", int((df["Final Status"] == "Inactive").sum()))
    if "Final Check" in df.columns:
        c4.metric("True Checks", int((df["Final Check"] == "True").sum()))
    preview = df.head(max_rows)
    if len(df) > max_rows:
        st.caption(
            "Showing first " + str(max_rows) + " of " + str(len(df)) +
            " rows. Download the full report for all data."
        )
    st.dataframe(preview, use_container_width=True, height=450)


# == Sidebar ==================================================================-
with st.sidebar:
    st.markdown("## Configuration")
    country = st.selectbox("Select Country", ["SG", "MY", "PH"])
    st.markdown("---")

    with st.expander("Lazada " + country):
        laz = st.file_uploader("Lazada File", type=["xlsx","xls","csv"], key="laz")

    with st.expander("Shopee " + country):
        sh_stk = st.file_uploader("Shopee Stock (ZIP)", type=["xlsx","xls","csv","zip"], key="sh_stk")
        sh_sts = st.file_uploader("Shopee Status", type=["xlsx","xls","csv","zip"], key="sh_sts")

    with st.expander("Zalora " + country):
        zal_stk = st.file_uploader("Zalora Stock", type=["xlsx","xls","csv"], key="zal_stk")
        zal_sts = st.file_uploader("Zalora Status", type=["xlsx","xls","csv"], key="zal_sts")

    tt_act = None
    tt_ina = None
    if country == "MY":
        with st.expander("TikTok MY"):
            tt_act = st.file_uploader("TikTok Active",   type=["xlsx","xls","csv"], key="tt_act")
            tt_ina = st.file_uploader("TikTok Inactive", type=["xlsx","xls","csv"], key="tt_ina")

    with st.expander("Reference Files"):
        cnt  = st.file_uploader("Content File",      type=["xlsx","xls","csv"], key="cnt")
        tc   = st.file_uploader("TC Inventory",      type=["xlsx","xls","csv"], key="tc")
        zec  = st.file_uploader("zEcom File",        type=["xlsx","xls","csv"], key="zec")
        alf  = st.file_uploader("ALL File",          type=["xlsx","xls","csv"], key="alf")
        excl = st.file_uploader("Exclusion List",    type=["xlsx","xls","csv"], key="excl")

    st.markdown("## Order & OMS Validation")
    with st.expander("Upload Order Reports"):
        order_pending = st.file_uploader("Pending Order Report", type=["xlsx","xls","csv"], key="order_pending")
        order_tc = st.file_uploader("TC Order Report", type=["xlsx","xls","csv"], key="order_tc")
        order_oms = st.file_uploader("OMS Order Report", type=["xlsx","xls","csv"], key="order_oms")
        seller_contacts = st.file_uploader("Seller Contact List (Optional)", type=["xlsx","xls","csv"], key="seller_contacts")

    st.markdown("---")
    run_btn = st.button("Run Validation", use_container_width=True, type="primary")


# == Main ======================================================================
st.title("Status Validation Analyzer")
st.write("Country: " + country + "  |  Upload files in the sidebar then click Run Validation.")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Status Report",
    "SKU Validation",
    "PID Validation",
    "Downloads",
    "Saved Reports",
    "QC Cross-Check",
    "Order & OMS Validation",
])

# == Run validation ============================================================-
if run_btn:
    with st.spinner("Loading files..."):
        try:
            data = load_all_files(
                country=country,
                lazada_file=laz,
                shopee_stock_file=sh_stk,
                shopee_status_file=sh_sts,
                zalora_stock_file=zal_stk,
                zalora_status_file=zal_sts,
                tiktok_active_file=tt_act,
                tiktok_inactive_file=tt_ina,
                content_file=cnt,
                tc_inv_file=tc,
                zecom_file=zec,
                all_file=alf,
                exclusion_file=excl,
            )

            parts = []
            for k, v in data.items():
                if isinstance(v, pd.DataFrame) and not v.empty:
                    parts.append(k + ":" + str(len(v)))
            st.success("Loaded: " + "  |  ".join(parts))

            with st.expander("Column names per file"):
                for k, v in data.items():
                    if isinstance(v, pd.DataFrame) and not v.empty:
                        st.write(k + " -> " + str(list(v.columns)))

            # Exclusion check
            from validators import _build_article_map, _build_excl_map
            excl_df     = data.get("exclusion", pd.DataFrame())
            content_df  = data.get("content", pd.DataFrame())
            if excl_df is not None and not excl_df.empty:
                excl_map   = _build_excl_map(excl_df)
                art_map    = _build_article_map(content_df)
                if not art_map:
                    st.error(
                        "EXCLUSION WILL NOT APPLY: " + str(len(excl_map)) +
                        " exclusion entries loaded but NO Content File "
                        "SKU->Article No mapping found. Upload the Content "
                        "File with 'SKU' and 'Article No' columns."
                    )
                else:
                    matched = set(art_map.values()) & set(excl_map.keys())
                    if matched:
                        st.success("Exclusion OK: " + str(len(matched)) + " Article Nos matched.")
                    else:
                        st.warning(
                            "EXCLUSION WILL NOT APPLY: 0 matches between "
                            "Content Article Nos and Exclusion list. "
                            "Check Article No format consistency."
                        )

        except Exception as e:
            st.error("Load error: " + str(e))
            st.exception(e)
            st.stop()

    with st.spinner("Running validations..."):
        try:
            sk = run_sku_validation(data, country)
            pi = run_pid_validation(data, country)

            # Build Status Report preview from sk and pi to save time
            cols = [
                "Marketplace", "Seller SKU", "TC SKU", "Article No", "MP Status",
                "TC Status", "e-com (Yes/No)", "Launch Date", "Exclusion", "ECOM Status",
                "MP Stock", "TC Stock", "Reserved Stock", "Max 0"
            ]
            sr_parts = []
            if not sk.empty:
                sr_parts.append(sk[cols])
            if not pi.empty:
                # pi uses SellerSku instead of Seller SKU, so rename it
                pi_cols = [c if c != "Seller SKU" else "SellerSku" for c in cols]
                temp_pi = pi[pi_cols].rename(columns={"SellerSku": "Seller SKU"})
                sr_parts.append(temp_pi)
            sr = pd.concat(sr_parts, ignore_index=True) if sr_parts else pd.DataFrame()
        except Exception as e:
            st.error("Validation error: " + str(e))
            st.exception(e)
            st.stop()

    with st.spinner("Saving report to disk..."):
        try:
            fname  = _make_filename(data, country)
            sheets = {}
            if not sk.empty: sheets["SKU Level Validation"] = sk
            if not pi.empty: sheets["PID Level Validation"] = pi

            if sheets:
                fpath = _write_report(sheets, fname)
                st.session_state["report_path"]  = fpath
                st.session_state["report_fname"] = fname
                st.session_state["sr_preview"]   = sr.head(500) if not sr.empty else pd.DataFrame()
                st.session_state["sk_preview"]   = sk.head(500)
                st.session_state["pi_preview"]   = pi.head(500)
                st.session_state["sr_len"]       = len(sr)
                st.session_state["sk_len"]       = len(sk)
                st.session_state["pi_len"]       = len(pi)
                st.session_state["country"]      = country
                # Free large DataFrames from memory immediately
                del sr, sk, pi, data
                gc.collect()
                st.success("Validation complete! Report saved.")
            else:
                st.warning("No data generated. Check your input files.")
        except Exception as e:
            st.error("Report save error: " + str(e))
            st.exception(e)


# == Tab 1 – Status Report ====================================================
with tab1:
    if "sr_preview" in st.session_state:
        rc  = st.session_state.get("country", country)
        df  = st.session_state["sr_preview"].copy()
        tot = st.session_state.get("sr_len", len(df))
        st.markdown("### Status Report - " + rc)
        if not df.empty and "Marketplace" in df.columns:
            opts = sorted(df["Marketplace"].unique())
            sel  = st.multiselect("Filter by Marketplace", opts, default=opts, key="f1")
            df   = df[df["Marketplace"].isin(sel)]
        show_df(df)
        if tot > 500:
            st.info("Preview shows first 500 rows. Download full report for all " + str(tot) + " rows.")
    else:
        st.info("Run validation to see results.")


# == Tab 2 – SKU Validation ==================================================-
with tab2:
    if "sk_preview" in st.session_state:
        rc  = st.session_state.get("country", country)
        df  = st.session_state["sk_preview"].copy()
        tot = st.session_state.get("sk_len", len(df))
        st.markdown("### SKU Validation - " + rc)
        if not df.empty:
            c1, c2 = st.columns(2)
            if "Final Check" in df.columns:
                opts = sorted(df["Final Check"].unique())
                sel  = c1.multiselect("Final Check", opts, default=opts, key="f2")
                df   = df[df["Final Check"].isin(sel)]
            if "Marketplace" in df.columns:
                opts2 = sorted(df["Marketplace"].unique())
                sel2  = c2.multiselect("Marketplace", opts2, default=opts2, key="f2b")
                df    = df[df["Marketplace"].isin(sel2)]
        show_df(df)
        if tot > 500:
            st.info("Preview shows first 500 rows. Download full report for all " + str(tot) + " rows.")
    else:
        st.info("Run validation to see results.")


# == Tab 3 – PID Validation ==================================================-
with tab3:
    if "pi_preview" in st.session_state:
        rc  = st.session_state.get("country", country)
        df  = st.session_state["pi_preview"].copy()
        tot = st.session_state.get("pi_len", len(df))
        st.markdown("### PID Validation - " + rc)
        if not df.empty:
            c1, c2 = st.columns(2)
            if "Final Check" in df.columns:
                opts = sorted(df["Final Check"].unique())
                sel  = c1.multiselect("Final Check", opts, default=opts, key="f3")
                df   = df[df["Final Check"].isin(sel)]
            if "Dual Status" in df.columns:
                opts2 = sorted(df["Dual Status"].unique())
                sel2  = c2.multiselect("Dual Status", opts2, default=opts2, key="f3b")
                df    = df[df["Dual Status"].isin(sel2)]
        show_df(df)
        if tot > 500:
            st.info("Preview shows first 500 rows. Download full report for all " + str(tot) + " rows.")
    else:
        st.info("Run validation to see results.")


# == Tab 4 – Downloads (current session) ======================================
with tab4:
    st.markdown("### Download Current Report")
    report_path  = st.session_state.get("report_path")
    report_fname = st.session_state.get("report_fname")

    if report_path and os.path.exists(report_path):
        size_mb = round(os.path.getsize(report_path) / (1024 * 1024), 2)
        st.info("File: **" + report_fname + "**  (" + str(size_mb) + " MB)")
        with open(report_path, "rb") as f:
            st.download_button(
                "Download Excel Report",
                data=f.read(),
                file_name=report_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_current",
            )
    else:
        st.info("Run validation first to generate a report.")


# == Tab 5 – Saved Reports (persistent) ======================================-
with tab5:
    st.markdown("### All Saved Reports")
    st.caption(
        "Reports are saved to the server and remain available until "
        "the app is restarted. You can download any previous report here."
    )

    saved = _list_saved_reports()
    if not saved:
        st.info("No saved reports yet. Run validation to generate one.")
    else:
        for fname, fpath, mtime, size_mb in saved:
            col1, col2, col3 = st.columns([5, 2, 2])
            col1.write("**" + fname + "**")
            col2.write(str(size_mb) + " MB")
            ts = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            col3.write(ts)
            with open(fpath, "rb") as f:
                st.download_button(
                    "Download " + fname,
                    data=f.read(),
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_" + fname,
                )
            st.markdown("---")

    if saved:
        if st.button("Clear all saved reports", type="secondary"):
            for _, fpath, _, _ in saved:
                try:
                    os.remove(fpath)
                except Exception:
                    pass
            st.success("All saved reports cleared.")
            st.rerun()


# == Tab 6 – QC Cross-Check ==================================================-
with tab6:
    st.markdown("### Stock Validation QC Cross-Check")
    st.caption(
        "Upload a completed Status/Stock validation Excel report (working sheet) "
        "shared by team members. The tool will run the correct matching logic on "
        "the raw reference files uploaded in the sidebar and highlight any mismatches."
    )
    
    working_file = st.file_uploader(
        "Upload Team Working Sheet (.xlsx)", 
        type=["xlsx"], 
        key="qc_working_file"
    )
    
    run_qc_btn = st.button("Run QC Cross-Check", type="primary", use_container_width=True)
    
    # == Persistent QC Report Download & Results display ======================
    qc_report_path = st.session_state.get("qc_report_path")
    qc_report_fname = st.session_state.get("qc_report_fname")
    if qc_report_path and os.path.exists(qc_report_path):
        size_mb = round(os.path.getsize(qc_report_path) / (1024 * 1024), 2)
        st.success(f"QC Audit complete! Report: **{qc_report_fname}** ({size_mb} MB)")
        with open(qc_report_path, "rb") as f:
            st.download_button(
                "Download QC Audit Report (Excel)",
                data=f.read(),
                file_name=qc_report_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_qc_report"
            )
        st.markdown("---")
        
        # Render SKU results from session state
        if st.session_state.get("qc_has_sku"):
            st.markdown("#### SKU Validation Audit Results (Lazada / Zalora)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Rows Checked", st.session_state["qc_sku_total_checked"])
            m_count = st.session_state["qc_sku_m_count"]
            c2.metric("Mismatched Rows", m_count, delta=-m_count if m_count > 0 else 0, delta_color="inverse")
            c3.metric("Missing Rows", st.session_state["qc_sku_missing_count"])
            c4.metric("Extra Rows", st.session_state["qc_sku_extra_count"])
            
            m_df = st.session_state["qc_sku_m_df"]
            if m_count > 0 and not m_df.empty:
                st.warning(f"Found {m_count} rows with value mismatches!")
                st.dataframe(m_df, use_container_width=True, hide_index=True)
            else:
                st.success("All SKU Validation values match perfectly!")
                
            missing_df = st.session_state["qc_sku_missing_df"]
            if not missing_df.empty:
                with st.expander("Show Missing Rows (Present in computed but not in working sheet)"):
                    st.dataframe(missing_df, use_container_width=True, hide_index=True)
                    
            extra_df = st.session_state["qc_sku_extra_df"]
            if not extra_df.empty:
                with st.expander("Show Extra Rows (Present in working sheet but not in computed results)"):
                    st.dataframe(extra_df, use_container_width=True, hide_index=True)
                    
        # Render PID results from session state
        if st.session_state.get("qc_has_pid"):
            if st.session_state.get("qc_has_sku"):
                st.markdown("---")
            st.markdown("#### PID Validation Audit Results (Shopee / TikTok)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Rows Checked", st.session_state["qc_pid_total_checked"])
            m_count = st.session_state["qc_pid_m_count"]
            c2.metric("Mismatched Rows", m_count, delta=-m_count if m_count > 0 else 0, delta_color="inverse")
            c3.metric("Missing Rows", st.session_state["qc_pid_missing_count"])
            c4.metric("Extra Rows", st.session_state["qc_pid_extra_count"])
            
            m_df = st.session_state["qc_pid_m_df"]
            if m_count > 0 and not m_df.empty:
                st.warning(f"Found {m_count} rows with value mismatches!")
                st.dataframe(m_df, use_container_width=True, hide_index=True)
            else:
                st.success("All PID Validation values match perfectly!")
                
            missing_df = st.session_state["qc_pid_missing_df"]
            if not missing_df.empty:
                with st.expander("Show Missing Rows (Present in computed but not in working sheet)"):
                    st.dataframe(missing_df, use_container_width=True, hide_index=True)
                    
            extra_df = st.session_state["qc_pid_extra_df"]
            if not extra_df.empty:
                with st.expander("Show Extra Rows (Present in working sheet but not in computed results)"):
                    st.dataframe(extra_df, use_container_width=True, hide_index=True)

    # == Calculation Trigger ==================================================
    if run_qc_btn:
        if not working_file:
            st.error("Please upload the team's working sheet first.")
        else:
            # Clear old QC results from session state
            for k in list(st.session_state.keys()):
                if k.startswith("qc_"):
                    del st.session_state[k]
                    
            # 1. Load raw files from sidebar
            with st.spinner("Loading raw reference files from sidebar..."):
                try:
                    data = load_all_files(
                        country=country,
                        lazada_file=laz,
                        shopee_stock_file=sh_stk,
                        shopee_status_file=sh_sts,
                        zalora_stock_file=zal_stk,
                        zalora_status_file=zal_sts,
                        tiktok_active_file=tt_act,
                        tiktok_inactive_file=tt_ina,
                        content_file=cnt,
                        tc_inv_file=tc,
                        zecom_file=zec,
                        all_file=alf,
                        exclusion_file=excl,
                    )
                except Exception as e:
                    st.error("Failed to load raw reference files: " + str(e))
                    st.stop()
            
            # 2. Open the uploaded working sheet to inspect sheets
            with st.spinner("Reading uploaded working sheet..."):
                try:
                    xls = pd.ExcelFile(working_file)
                    sheet_names = xls.sheet_names
                except Exception as e:
                    st.error("Failed to read working sheet Excel structure: " + str(e))
                    st.stop()
            
            # 3. Perform comparison if sheets are found
            sku_sheet_name = None
            pid_sheet_name = None
            for name in sheet_names:
                name_lower = name.lower().strip()
                if "sku" in name_lower or "lazada" in name_lower or "zalora" in name_lower:
                    sku_sheet_name = name
                elif "pid" in name_lower or "shopee" in name_lower or "tiktok" in name_lower:
                    pid_sheet_name = name
            
            has_sku = sku_sheet_name is not None
            has_pid = pid_sheet_name is not None
            
            if not has_sku and not has_pid:
                st.error("The uploaded file does not contain 'SKU Level Validation' (or SKU Validation) or 'PID Level Validation' (or PID Validation) sheets. Found sheets: " + ", ".join(sheet_names))
                st.stop()
                
            # Helper functions for standardizing working sheet format
            def standardize_columns(df):
                mapping = {
                    "sellersku": "Seller SKU",
                    "seller sku": "Seller SKU",
                    "ecom (yes/no)": "e-com (Yes/No)",
                    "e-com (yes/no)": "e-com (Yes/No)",
                    "ecom(yes/no)": "e-com (Yes/No)",
                    "graas status": "TC Status",
                    "tc status": "TC Status",
                    "final status (t/f)": "Final Check",
                    "final check": "Final Check",
                    "final status": "Final Status",
                    "stock check": "Stock Check",
                    "max setup": "Max Setup",
                    "update 0": "Update 0",
                }
                new_cols = {}
                for col in df.columns:
                    norm_col = str(col).strip().lower()
                    if norm_col in mapping:
                        new_cols[col] = mapping[norm_col]
                    else:
                        cleaned_col = norm_col.replace(" ", "").replace("-", "").replace("_", "")
                        matched = False
                        for k, v in mapping.items():
                            cleaned_k = k.replace(" ", "").replace("-", "").replace("_", "")
                            if cleaned_col == cleaned_k:
                                new_cols[col] = v
                                matched = True
                                break
                        if not matched:
                            new_cols[col] = col
                return df.rename(columns=new_cols)

            def prepare_working_df(working_df, correct_df, sheet_name, country):
                working_df = standardize_columns(working_df)
                
                # If Marketplace column is missing, populate it using correct_df mapping
                if "Marketplace" not in working_df.columns:
                    if "Seller SKU" in working_df.columns and not correct_df.empty and "Seller SKU" in correct_df.columns:
                        sku_to_mp = dict(zip(correct_df["Seller SKU"], correct_df["Marketplace"]))
                        working_df["Marketplace"] = working_df["Seller SKU"].map(sku_to_mp)
                    else:
                        working_df["Marketplace"] = None
                    
                    fallback_mp = "Unknown"
                    sheet_lower = sheet_name.lower()
                    if "lazada" in sheet_lower:
                        fallback_mp = f"Lazada {country}"
                    elif "zalora" in sheet_lower:
                        fallback_mp = f"Zalora {country}"
                    elif "shopee" in sheet_lower:
                        fallback_mp = f"Shopee {country}"
                    elif "tiktok" in sheet_lower:
                        fallback_mp = "TikTok MY"
                    
                    working_df["Marketplace"] = working_df["Marketplace"].fillna(fallback_mp)
                return working_df

            # Helper function for values matching
            def values_match_fast(val1, val2):
                if val1 is val2:
                    return True
                
                is_nan1 = (val1 is None or val1 != val1 or val1 == "nan" or val1 == "NaN")
                is_nan2 = (val2 is None or val2 != val2 or val2 == "nan" or val2 == "NaN")
                
                if is_nan1 and is_nan2:
                    return True
                if is_nan1 or is_nan2:
                    return False
                    
                s1 = str(val1).strip()
                s2 = str(val2).strip()
                if s1 == s2:
                    return True
                    
                s1_lower = s1.lower()
                s2_lower = s2.lower()
                if s1_lower == s2_lower:
                    return True
                    
                try:
                    f1 = float(s1.replace(",", ""))
                    f2 = float(s2.replace(",", ""))
                    if abs(f1 - f2) < 1e-5:
                        return True
                except ValueError:
                    pass
                return False

            def perform_qc_comparison(working_df, correct_df, key_cols, compare_cols):
                working_df = working_df.copy()
                correct_df = correct_df.copy()
                
                # Verify columns exist
                for col in key_cols:
                    if col not in working_df.columns:
                        return None, f"Missing key column '{col}' in working sheet."
                    if col not in correct_df.columns:
                        return None, f"Missing key column '{col}' in computed sheet."
                        
                # Standardize keys
                for col in key_cols:
                    working_df[col] = working_df[col].astype(str).str.strip()
                    correct_df[col] = correct_df[col].astype(str).str.strip()
                    
                working_df["_merge_key"] = working_df[key_cols].agg("||".join, axis=1)
                correct_df["_merge_key"] = correct_df[key_cols].agg("||".join, axis=1)
                
                # Deduplicate merge keys
                working_df = working_df.drop_duplicates(subset=["_merge_key"])
                correct_df = correct_df.drop_duplicates(subset=["_merge_key"])
                
                working_keys = set(working_df["_merge_key"])
                correct_keys = set(correct_df["_merge_key"])
                
                missing_in_working = correct_df[~correct_df["_merge_key"].isin(working_keys)].copy()
                extra_in_working = working_df[~working_df["_merge_key"].isin(correct_keys)].copy()
                
                common_keys = working_keys.intersection(correct_keys)
                
                w_indexed = working_df.set_index("_merge_key")
                c_indexed = correct_df.set_index("_merge_key")
                
                # Use fast dict mapping to bypass pandas .loc indexing overhead in python loop
                w_dict = w_indexed.to_dict(orient="index")
                c_dict = c_indexed.to_dict(orient="index")
                
                mismatches = []
                cols_to_compare = [c for c in compare_cols if c in working_df.columns and c in correct_df.columns]
                
                for key in common_keys:
                    w_row = w_dict[key]
                    c_row = c_dict[key]
                    
                    row_diff = {}
                    for col in cols_to_compare:
                        w_val = w_row[col]
                        c_val = c_row[col]
                        if not values_match_fast(w_val, c_val):
                            row_diff[col] = (w_val, c_val)
                            
                    if row_diff:
                        disp_vals = {col: w_row[col] for col in key_cols}
                        mismatches.append({
                            **disp_vals,
                            "diffs": row_diff
                        })
                        
                return {
                    "mismatches": mismatches,
                    "missing_in_working": missing_in_working,
                    "extra_in_working": extra_in_working,
                    "total_checked": len(common_keys)
                }, None

            qc_sheets = {}
            st.session_state["qc_has_sku"] = has_sku
            st.session_state["qc_has_pid"] = has_pid

            # SKU Validation Audit
            if has_sku:
                with st.spinner("Running SKU Validation correct logic..."):
                    try:
                        correct_sku = run_sku_validation(data, country)
                        working_sku = pd.read_excel(working_file, sheet_name=sku_sheet_name)
                        working_sku = prepare_working_df(working_sku, correct_sku, sku_sheet_name, country)
                    except Exception as e:
                        st.error("SKU Validation Load Error: " + str(e))
                        st.stop()
                
                if correct_sku.empty:
                    st.warning("No calculated results for SKU Validation (Lazada/Zalora raw files are empty).")
                elif working_sku.empty:
                    st.warning("The uploaded SKU Validation sheet is empty.")
                else:
                    if "SellerSku" in correct_sku.columns:
                        correct_sku = correct_sku.rename(columns={"SellerSku": "Seller SKU"})
                        
                    sku_cols = ["TC SKU", "Article No", "MP Status", "TC Status", "e-com (Yes/No)", "Launch Date", "Exclusion", "ECOM Status", "MP Stock", "TC Stock", "Reserved Stock", "Max 0", "Final Status", "Comments", "Final Check", "Stock Check", "Remarks", "Max Setup", "Update 0"]
                    sku_res, err = perform_qc_comparison(
                        working_df=working_sku,
                        correct_df=correct_sku,
                        key_cols=["Marketplace", "Seller SKU"],
                        compare_cols=sku_cols
                    )
                    
                    if err:
                        st.error(f"SKU Comparison Error: {err}")
                    else:
                        m_rows = []
                        for m in sku_res["mismatches"]:
                            for col, (w_val, c_val) in m["diffs"].items():
                                m_rows.append({
                                    "Marketplace": m["Marketplace"],
                                    "Seller SKU": m["Seller SKU"],
                                    "Field Name": col,
                                    "Team Value (Working)": w_val,
                                    "Expected Value (Computed)": c_val
                                })
                        sku_mismatch_df = pd.DataFrame(m_rows) if m_rows else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Field Name", "Team Value (Working)", "Expected Value (Computed)"])
                        sku_missing_df = sku_res["missing_in_working"][["Marketplace", "Seller SKU", "Final Status"]].copy() if not sku_res["missing_in_working"].empty else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Final Status"])
                        sku_extra_df = sku_res["extra_in_working"][["Marketplace", "Seller SKU", "Final Status"]].copy() if not sku_res["extra_in_working"].empty else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Final Status"])
                        
                        qc_sheets["SKU Value Mismatches"] = sku_mismatch_df
                        qc_sheets["SKU Missing Rows"] = sku_missing_df
                        qc_sheets["SKU Extra Rows"] = sku_extra_df

                        # Save values to session state
                        st.session_state["qc_sku_total_checked"] = sku_res["total_checked"]
                        st.session_state["qc_sku_m_count"] = len(sku_res["mismatches"])
                        st.session_state["qc_sku_missing_count"] = len(sku_missing_df)
                        st.session_state["qc_sku_extra_count"] = len(sku_extra_df)
                        st.session_state["qc_sku_m_df"] = sku_mismatch_df
                        st.session_state["qc_sku_missing_df"] = sku_missing_df
                        st.session_state["qc_sku_extra_df"] = sku_extra_df

            # PID Validation Audit
            if has_pid:
                with st.spinner("Running PID Validation correct logic..."):
                    try:
                        correct_pid = run_pid_validation(data, country)
                        working_pid = pd.read_excel(working_file, sheet_name=pid_sheet_name)
                        working_pid = prepare_working_df(working_pid, correct_pid, pid_sheet_name, country)
                    except Exception as e:
                        st.error("PID Validation Load Error: " + str(e))
                        st.stop()
                        
                if correct_pid.empty:
                    st.warning("No calculated results for PID Validation (Shopee/TikTok raw files are empty).")
                elif working_pid.empty:
                    st.warning("The uploaded PID Validation sheet is empty.")
                else:
                    if "SellerSku" in correct_pid.columns:
                        correct_pid = correct_pid.rename(columns={"SellerSku": "Seller SKU"})
                        
                    pid_cols = ["TC SKU", "Product ID", "Article No", "MP Status", "TC Status", "e-com (Yes/No)", "Launch Date", "Exclusion", "ECOM Status", "Final Status", "Comments", "Final Check", "Dual Status", "Consolidated SUM QTY", "MP Stock", "TC Stock", "Reserved Stock", "Max 0", "Stock Check", "Remarks", "Max Setup", "Update 0"]
                    pid_res, err = perform_qc_comparison(
                        working_df=working_pid,
                        correct_df=correct_pid,
                        key_cols=["Marketplace", "Seller SKU"],
                        compare_cols=pid_cols
                    )
                    
                    if err:
                        st.error(f"PID Comparison Error: {err}")
                    else:
                        m_rows = []
                        for m in pid_res["mismatches"]:
                            for col, (w_val, c_val) in m["diffs"].items():
                                m_rows.append({
                                    "Marketplace": m["Marketplace"],
                                    "Seller SKU": m["Seller SKU"],
                                    "Field Name": col,
                                    "Team Value (Working)": w_val,
                                    "Expected Value (Computed)": c_val
                                })
                        pid_mismatch_df = pd.DataFrame(m_rows) if m_rows else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Field Name", "Team Value (Working)", "Expected Value (Computed)"])
                        pid_missing_df = pid_res["missing_in_working"][["Marketplace", "Seller SKU", "Final Status"]].copy() if not pid_res["missing_in_working"].empty else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Final Status"])
                        pid_extra_df = pid_res["extra_in_working"][["Marketplace", "Seller SKU", "Final Status"]].copy() if not pid_res["extra_in_working"].empty else pd.DataFrame(columns=["Marketplace", "Seller SKU", "Final Status"])
                        
                        qc_sheets["PID Value Mismatches"] = pid_mismatch_df
                        qc_sheets["PID Missing Rows"] = pid_missing_df
                        qc_sheets["PID Extra Rows"] = pid_extra_df

                        # Save values to session state
                        st.session_state["qc_pid_total_checked"] = pid_res["total_checked"]
                        st.session_state["qc_pid_m_count"] = len(pid_res["mismatches"])
                        st.session_state["qc_pid_missing_count"] = len(pid_missing_df)
                        st.session_state["qc_pid_extra_count"] = len(pid_extra_df)
                        st.session_state["qc_pid_m_df"] = pid_mismatch_df
                        st.session_state["qc_pid_missing_df"] = pid_missing_df
                        st.session_state["qc_pid_extra_df"] = pid_extra_df

            # Save the compiled QC report Excel file to disk
            if qc_sheets:
                with st.spinner("Generating and saving QC Audit report..."):
                    try:
                        today = datetime.today().strftime("%Y-%m-%d")
                        qc_fname = f"QC_Audit_Report_{country}_{today}.xlsx"
                        qc_fpath = _write_qc_report(qc_sheets, qc_fname)
                        st.session_state["qc_report_path"] = qc_fpath
                        st.session_state["qc_report_fname"] = qc_fname
                    except Exception as e:
                        st.error("Failed to generate QC Excel report: " + str(e))

            st.rerun()


# == Tab 7 – Order & OMS Validation ==========================================-
with tab7:
    st.markdown("### Order SLA Enrichment & OMS Status Validation")
    st.caption("Upload the Pending Order Report, TC Order Report, and OMS Order Report in the sidebar to run validations and email reports directly to the sellers.")
    
    # Check if files are uploaded
    if not (order_pending and order_tc and order_oms):
        st.info("Please upload the Pending Order Report, TC Order Report, and OMS Order Report in the sidebar under 'Order & OMS Validation' to get started.")
    else:
        # Create a button to run the validation
        if st.button("Run Order Validation & Analysis", type="primary", use_container_width=True):
            with st.spinner("Processing reports and running validations..."):
                try:
                    res = process_and_validate_orders(order_pending, order_tc, order_oms, seller_contacts)
                    st.session_state["order_enriched_df"] = res["enriched_pending_df"]
                    st.session_state["order_disc_df"] = res["discrepancies_df"]
                    st.session_state["order_summary"] = res["summary"]
                    st.session_state["order_groups"] = res["seller_groups"]
                    st.session_state["order_id_col"] = res["pending_order_id_col"]
                    st.success("Validation complete! See results below.")
                except Exception as e:
                    st.error(f"Error during order processing: {str(e)}")
                    st.exception(e)
                    
        # Check if we have results in session_state
        if "order_summary" in st.session_state:
            summary = st.session_state["order_summary"]
            enriched_df = st.session_state["order_enriched_df"]
            disc_df = st.session_state["order_disc_df"]
            seller_groups = st.session_state["order_groups"]
            
            # Display metrics
            st.markdown("#### Key Metrics")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Pending Orders", summary["total_pending_orders"])
            m2.metric("Enriched SLAs from TC", summary["enriched_sla_count"])
            m3.metric("Validation Mismatches", summary["total_discrepancies"], delta=-summary["total_discrepancies"] if summary["total_discrepancies"] > 0 else 0, delta_color="inverse")
            m4.metric("Total Sellers / Stores", summary["total_sellers"])
            
            # Display layout of tables
            st.markdown("#### Detailed Results")
            sub_tab1, sub_tab2, sub_tab3 = st.tabs([
                "Enriched Pending Orders", 
                "OMS vs TC Discrepancies", 
                "Seller Grouping & Email Center"
            ])
            
            with sub_tab1:
                st.markdown("##### Enriched Pending Orders Report")
                st.dataframe(enriched_df, use_container_width=True, hide_index=True)
                
            with sub_tab2:
                st.markdown("##### Validation Failures & Status Discrepancies")
                if disc_df.empty:
                    st.success("No status discrepancies or validation failures identified!")
                else:
                    st.warning(f"Found {len(disc_df)} discrepancies/warnings.")
                    st.dataframe(disc_df, use_container_width=True, hide_index=True)
                    
            with sub_tab3:
                st.markdown("##### SMTP Email Configuration")
                # Load credentials from secrets or default to empty
                secrets_smtp = st.secrets.get("smtp", {}) if st.secrets else {}
                
                # Show expander for email settings
                with st.expander("Configure SMTP Email Settings"):
                    c_host = st.text_input("SMTP Server Host", value=secrets_smtp.get("host", "smtp.office365.com"), key="smtp_host")
                    c_port = st.text_input("SMTP Port", value=str(secrets_smtp.get("port", 587)), key="smtp_port")
                    c_user = st.text_input("SMTP Username", value=secrets_smtp.get("user", ""), key="smtp_user")
                    c_pass = st.text_input("SMTP Password", type="password", value=secrets_smtp.get("password", ""), key="smtp_pass")
                    c_sender = st.text_input("Sender Email Address", value=secrets_smtp.get("sender_email", c_user), key="smtp_sender")
                    c_tls = st.checkbox("Use TLS", value=secrets_smtp.get("use_tls", True), key="smtp_tls")
                    
                    if st.button("Test Connection"):
                        is_ok, msg = test_smtp_connection(c_host, c_port, c_user, c_pass, c_tls)
                        if is_ok:
                            st.success(msg)
                        else:
                            st.error(msg)
                            
                # Display Sellers list
                st.markdown("##### Send Daily Report to Sellers")
                smtp_config = {
                    "host": c_host,
                    "port": c_port,
                    "user": c_user,
                    "password": c_pass,
                    "sender_email": c_sender,
                    "use_tls": c_tls
                }
                
                st.info("You can send the filtered Excel report directly to each seller. Make sure to specify their email address below.")
                
                send_all_btn = st.button("Send Reports to All Sellers", type="secondary", use_container_width=True)
                
                seller_email_inputs = {}
                for idx, (seller_name, info) in enumerate(seller_groups.items()):
                    with st.container():
                        col1, col2, col3, col4 = st.columns([3, 4, 2, 2])
                        col1.markdown(f"**{seller_name}**")
                        col2.write(f"Orders: {len(info['df'])}")
                        
                        default_email = info["email"]
                        recipient = col3.text_input("Email", value=default_email, key=f"email_{seller_name}_{idx}", label_visibility="collapsed")
                        seller_email_inputs[seller_name] = recipient
                        
                        send_single = col4.button("Send Email", key=f"send_{seller_name}_{idx}", use_container_width=True)
                        if send_single:
                            if not recipient:
                                st.error("Please enter a valid email address.")
                            else:
                                with st.spinner(f"Sending email to {seller_name}..."):
                                    ok, msg = send_seller_report_email(smtp_config, seller_name, recipient, info["df"], disc_df)
                                    if ok:
                                        st.success(f"Sent to {seller_name}!")
                                    else:
                                        st.error(msg)
                                        
                if send_all_btn:
                    success_count = 0
                    fail_count = 0
                    progress_bar = st.progress(0)
                    total_sellers = len(seller_groups)
                    
                    for i, (seller_name, info) in enumerate(seller_groups.items()):
                        recipient = seller_email_inputs[seller_name]
                        if not recipient:
                            st.warning(f"Skipping {seller_name} - No email specified.")
                            fail_count += 1
                        else:
                            ok, msg = send_seller_report_email(smtp_config, seller_name, recipient, info["df"], disc_df)
                            if ok:
                                success_count += 1
                            else:
                                st.error(f"Failed for {seller_name}: {msg}")
                                fail_count += 1
                        progress_bar.progress((i + 1) / total_sellers)
                    
                    st.success(f"Finished sending! Success: {success_count}, Failed/Skipped: {fail_count}")

