# -*- coding: utf-8 -*-
import io
import os
import gc
import tempfile
import json
import requests
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="Pending Order & OMS Validation",
    page_icon="email",
    layout="wide",
)

from styles import inject_css
from order_processor import process_and_validate_orders
from email_sender import test_smtp_connection, send_seller_report_email
import excel_formatter

inject_css()

# Load default SMTP config from config.json
smtp_defaults = {}
try:
    with open("config.json", "r") as config_file:
        cfg = json.load(config_file)
        smtp_defaults = cfg.get("smtp_config", {})
except Exception:
    pass

# Custom title and introduction
st.title("Pending Order SLA Enrichment & OMS Status Validation")
st.write("Upload the daily SLA Report (GSheet Link), Marketplace Order Reports (TC Reports), and OMS Report (Sales Order file) in the sidebar to run validations and email reports directly to the sellers.")

# == SMTP Email Configuration (always available, independent of validation) ===
st.markdown("### 📧 SMTP Email Configuration")
try:
    secrets_smtp = st.secrets.get("smtp", {}) if st.secrets else {}
except Exception:
    secrets_smtp = {}
with st.expander("Configure SMTP Email Settings", expanded=False):
    c_host = st.text_input("SMTP Server Host", value=smtp_defaults.get("host", "smtp.office365.com"), key="smtp_host")
    c_port = st.text_input("SMTP Port", value=str(smtp_defaults.get("port", 587)), key="smtp_port")
    c_user = st.text_input("SMTP Username", value=smtp_defaults.get("user", ""), key="smtp_user")
    c_pass = st.text_input("SMTP Password", type="password", value=smtp_defaults.get("password", ""), key="smtp_pass")
    c_sender = st.text_input("Sender Email Address", value=smtp_defaults.get("sender_email", smtp_defaults.get("user", "")), key="smtp_sender")
    c_tls = st.checkbox("Use TLS", value=smtp_defaults.get("use_tls", True), key="smtp_tls")
    
    if st.button("Test Connection"):
        is_ok, msg = test_smtp_connection(c_host, c_port, c_user, c_pass, c_tls)
        if is_ok:
            st.success(msg)
            try:
                with open("config.json", "r") as f:
                    cfg_data = json.load(f)
            except Exception:
                cfg_data = {}
            cfg_data["smtp_config"] = {
                "host": c_host,
                "port": int(c_port) if c_port.isdigit() else 587,
                "user": c_user,
                "password": c_pass,
                "sender_email": c_sender,
                "use_tls": c_tls
            }
            try:
                with open("config.json", "w") as f:
                    json.dump(cfg_data, f, indent=4)
                st.info("Saved SMTP configuration to config.json!")
            except Exception as save_err:
                st.warning(f"Could not save config.json: {save_err}")
        else:
            st.error(msg)

smtp_config = {
    "host": c_host,
    "port": c_port,
    "user": c_user,
    "password": c_pass,
    "sender_email": c_sender,
    "use_tls": c_tls
}

# == Sidebar ==================================================================-
with st.sidebar:
    st.markdown("## Configuration")
    st.markdown("Upload the daily reports below:")
    
    st.markdown("**1. Pending Order Report**")
    pending_file_upload = st.file_uploader("Upload Excel/CSV File", type=["xlsx","xls","csv"], key="pending_file_upload")
    gsheet_url = st.text_input("Or enter Google Sheet Link", placeholder="https://docs.google.com/spreadsheets/d/...")
    
    order_pending = pending_file_upload if pending_file_upload is not None else (gsheet_url.strip() if gsheet_url.strip() else None)
        
    order_tc = st.file_uploader("2. TC Order Report", type=["xlsx","xls","csv"], key="order_tc")
    marketplace_reports = st.file_uploader("3. Market Place Reports (Multiple Upload)", type=["xlsx","xls","csv"], accept_multiple_files=True, key="marketplace_reports")
    order_oms = st.file_uploader("4. OMS Report (Sales Order file)", type=["xlsx","xls","csv"], key="order_oms")

    st.markdown("---")
    run_btn = st.button("Run Order Validation", use_container_width=True, type="primary")

# == Main Screen ===============================================================
# Check if files are uploaded
has_pending = order_pending is not None
has_tc_uploaded = order_tc is not None
has_mp_uploaded = (marketplace_reports is not None and len(marketplace_reports) > 0)
has_oms = order_oms is not None

# Backward compatibility for other checks in app.py
has_tc = has_tc_uploaded or has_mp_uploaded

is_valid_combo = False
mode_desc = ""

# 1. Mode 1: 1 & 4 Uploading Option -> Pending order creation (Pending Report + OMS)
if has_pending and has_oms and not has_tc_uploaded and not has_mp_uploaded:
    is_valid_combo = True
    mode_desc = "1. Pending Order Creation (Pending Report + OMS)"
# 2. Mode 2: 2 & 3 Uploading Option -> Market Place Order Check (TC + Marketplace)
elif has_tc_uploaded and has_mp_uploaded and not has_pending and not has_oms:
    is_valid_combo = True
    mode_desc = "2. Market Place Order Check (TC + Marketplace)"
# 3. Mode 3: 2 & 4 Uploading Option -> Order Status Reconciliation (TC + OMS)
elif has_tc_uploaded and has_oms and not has_mp_uploaded and not has_pending:
    is_valid_combo = True
    mode_desc = "3. Order Status Reconciliation (TC + OMS)"
# 4. Mode 4: 2, 3 & 4 Uploading Option -> Order Flow Check (TC + Marketplace + OMS)
elif has_tc_uploaded and has_mp_uploaded and has_oms and not has_pending:
    is_valid_combo = True
    mode_desc = "4. Order Flow Check (TC + Marketplace + OMS)"
# 5. Mode 5: Full Validation Mode (Pending Report + TC + OMS + Marketplace)
elif has_pending and has_tc_uploaded and has_oms:
    is_valid_combo = True
    mode_desc = "Full Pending Order Validation (Pending Report + TC + OMS)"

if not is_valid_combo:
    st.info(
        "💡 **Please upload one of the following combinations to start validation:**\n\n"
        "1. **Option 1 & 4 (Pending Order Creation)**: Upload Pending Order Report (1) and OMS Report (4).\n"
        "2. **Option 2 & 3 (Market Place Order Check)**: Upload TC Order Report (2) and Market Place Reports (3).\n"
        "3. **Option 2 & 4 (Order Status Reconciliation)**: Upload TC Order Report (2) and OMS Report (4).\n"
        "4. **Option 2, 3 & 4 (Order Flow Check)**: Upload TC Order Report (2), Market Place Reports (3), and OMS Report (4)."
    )
else:
    # Display active mode banner
    st.success(f"🎯 **Active Mode: {mode_desc}**")

    # Trigger validation by clicking sidebar button
    if run_btn:
        with st.spinner("Processing reports and running validations..."):
            try:
                res = process_and_validate_orders(
                    pending_file=order_pending,
                    tc_file=order_tc,
                    marketplace_file=marketplace_reports,
                    oms_file=order_oms,
                    contacts_file=None
                )
                st.session_state["order_enriched_df"] = res["enriched_pending_df"]
                st.session_state["order_disc_df"] = res["discrepancies_df"]
                st.session_state["order_summary"] = res["summary"]
                st.session_state["order_groups"] = res["seller_groups"]
                st.session_state["order_id_col"] = res["pending_order_id_col"]
                st.session_state["order_country_reports"] = res["country_reports"]
                st.session_state["order_ref_date_dmy"] = res.get("ref_date_dmy", "")
                st.success("Validation complete! See results below.")
                
                # Automatically share discrepancies to Slack group
                disc_df_to_slack = res["discrepancies_df"]
                if not disc_df_to_slack.empty:
                    smtp_cfg = {}
                    if st.session_state.get("smtp_host"):
                        smtp_cfg = {
                            "host": st.session_state.get("smtp_host"),
                            "port": int(st.session_state.get("smtp_port")) if str(st.session_state.get("smtp_port")).isdigit() else 587,
                            "user": st.session_state.get("smtp_user"),
                            "password": st.session_state.get("smtp_pass"),
                            "sender_email": st.session_state.get("smtp_sender"),
                            "use_tls": st.session_state.get("smtp_tls", True)
                        }
                    
                    if not smtp_cfg or not smtp_cfg.get("host") or not smtp_cfg.get("user"):
                        try:
                            with open("config.json", "r") as config_file:
                                cfg = json.load(config_file)
                                smtp_cfg = cfg.get("smtp_config", {})
                        except Exception:
                            pass
                    
                    if smtp_cfg and smtp_cfg.get("host") and smtp_cfg.get("user"):
                        from email_sender import send_discrepancies_to_slack_email
                        success_slack, msg_slack = send_discrepancies_to_slack_email(
                            smtp_config=smtp_cfg,
                            discrepancies_df=disc_df_to_slack,
                            ref_date_str=res.get("ref_date_dmy", "")
                        )
                        if success_slack:
                            st.success("📢 Automatically shared status discrepancies report to Slack channel `#order-related-issues`!")
                        else:
                            st.warning(f"⚠️ Could not automatically share discrepancies to Slack: {msg_slack}")
            except Exception as e:
                st.error(f"Error during order processing: {str(e)}")
                st.exception(e)
                
    # Check if we have results in session_state
    if "order_summary" in st.session_state:
        summary = st.session_state["order_summary"]
        enriched_df = st.session_state["order_enriched_df"]
        disc_df = st.session_state["order_disc_df"]
        seller_groups = st.session_state["order_groups"]
        country_reports = st.session_state.get("order_country_reports", {})
        
        mode = summary.get("mode", "standard")
        has_pending = (mode != "tc_marketplace")

        # Display metrics
        st.markdown("### Key Metrics")
        if mode == "tc_oms":
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Number of Orders", summary["total_pending_orders"])
            m2.metric("Reflecting in OMS", summary["pushed_count"])
            m3.metric("Missing in OMS", summary["not_pushed_count"])
            mismatch_order_count = disc_df["Order ID"].nunique() if not disc_df.empty and "Order ID" in disc_df.columns else 0
            m4.metric("Mismatch Order Count", mismatch_order_count)
            
            if summary["all_imported_to_tc"]:
                st.success("🎉 **All active TC orders are successfully verified in OMS!**")
        elif mode in ("order_status_reconciliation", "tc_oms"):
            m1, m2, m3 = st.columns(3)
            m1.metric("Total TC Orders", summary["total_pending_orders"])
            m2.metric("Reflecting in OMS", summary["pushed_count"])
            m3.metric("Status Discrepancies in OMS", summary["not_pushed_count"])
            
            if summary["all_imported_to_tc"]:
                st.success("🎉 **All TC orders are successfully verified in OMS!**")
            else:
                st.error(f"⚠️ **Found {summary['not_pushed_count']} status discrepancies between TC and OMS!**")
        elif mode == "order_flow_check":
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Marketplace & TC Orders", summary["total_pending_orders"])
            m2.metric("Reflecting in OMS", summary["pushed_count"])
            m3.metric("Missing & Mismatches", summary["not_pushed_count"])
            
            if summary["all_imported_to_tc"]:
                st.success("🎉 **All Marketplace orders reflect in TC and match OMS!**")
            else:
                st.error(f"⚠️ **Found {summary['not_pushed_count']} missing or status mismatch orders!**")
        elif mode == "tc_marketplace":
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Marketplace Orders", summary["total_pending_orders"])
            m2.metric("Reflected in TC", summary["pushed_count"])
            m3.metric("Missing from TC", summary["not_pushed_count"])
            
            if summary["all_imported_to_tc"]:
                st.success("🎉 **All Marketplace orders are successfully imported to TC!**")
            else:
                st.error(f"⚠️ **Found {summary['not_pushed_count']} orders missing from TC!**")
        else:
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total Pending Orders", summary["total_pending_orders"])
            m2.metric("Successfully Pushed", summary["pushed_count"])
            m3.metric("Not Pushed to OMS", summary["not_pushed_count"])
            m4.metric("Unpaid Orders", summary["unpaid_count"])
            m5.metric("Status Discrepancies", summary["total_discrepancies"], 
                      delta=summary["total_discrepancies"] if summary["total_discrepancies"] > 0 else None, 
                      delta_color="inverse")
        
        # Render Downloads & Detailed Results based on mode
        if mode == "gsheet_oms":
            # Mode 1 ONLY: Pending Order Creation (File 1 + File 4).
            # This country-wise report / email center section is intentionally
            # NOT shown for any other mode (Market Place Order Check, Order
            # Status Reconciliation, Order Flow Check, or Full Validation).
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Country SLA & Pivot Reports (Styled)</h3>', unsafe_allow_html=True)
            
            c_cols = st.columns(3)
            ref_date_dmy = st.session_state.get("order_ref_date_dmy", "")
            for idx, country in enumerate(["SG", "MY", "PH"]):
                c_data = country_reports.get(country, {})
                raw_df = c_data.get("raw_df", pd.DataFrame())
                pivot_df = c_data.get("pivot_df", pd.DataFrame())
                summary_df = c_data.get("summary_df", pd.DataFrame())
                
                if not raw_df.empty:
                    wb = excel_formatter.generate_excel_workbook(country, raw_df, pivot_df, summary_df, ref_date_dmy)
                    c_buffer = io.BytesIO()
                    wb.save(c_buffer)
                    
                    with c_cols[idx]:
                        st.download_button(
                            label=f"📥 Download {country} Report",
                            data=c_buffer.getvalue(),
                            file_name=f"Pending order report - {country} {ref_date_dmy if ref_date_dmy else datetime.today().strftime('%d-%m-%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"dl_{country}"
                        )
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Consolidated download container
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Consolidated QC Report</h3>', unsafe_allow_html=True)
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                export_enriched_df = enriched_df.drop(columns=["Correct Order Number", "SLA Source"], errors="ignore")
                export_enriched_df.to_excel(writer, sheet_name="SLA Report", index=False)
                disc_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                
                excel_formatter.format_data_sheet(writer.sheets["SLA Report"], export_enriched_df)
                excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], disc_df)
                
                for country in ["SG", "MY", "PH"]:
                    c_data = country_reports.get(country, {})
                    pivot_df = c_data.get("pivot_df", pd.DataFrame())
                    raw_df = c_data.get("raw_df", pd.DataFrame())
                    summary_df = c_data.get("summary_df", pd.DataFrame())
                    
                    if not raw_df.empty:
                        excel_formatter.add_country_sheets_to_workbook(
                            writer.book, country, raw_df, pivot_df, summary_df, ref_date_dmy
                        )
                
                wb = writer.book
                sheet_order = []
                for c in ["SG", "MY", "PH"]:
                    if f"{c} Summary" in wb.sheetnames:
                        sheet_order.append(wb[f"{c} Summary"])
                if "SLA Report" in wb.sheetnames:
                    sheet_order.append(wb["SLA Report"])
                if "Status Discrepancies" in wb.sheetnames:
                    sheet_order.append(wb["Status Discrepancies"])
                for c in ["SG", "MY", "PH"]:
                    if f"{c} Data" in wb.sheetnames:
                        sheet_order.append(wb[f"{c} Data"])
                wb._sheets = sheet_order
            
            st.download_button(
                label="📥 Download Detailed Excel QC Report",
                data=excel_buffer.getvalue(),
                file_name=f"Pending order report - Consolidated {ref_date_dmy if ref_date_dmy else datetime.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_consolidated"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("### Detailed Results")
            sub_tab1, sub_tab2, sub_tab3 = st.tabs([
                "SLA Report", 
                "Status Discrepancies", 
                "Country Reports & Email Center"
            ])
            
            with sub_tab1:
                st.markdown("#### Enriched SLA Report (Main Sheet)")
                display_enriched = enriched_df.drop(columns=["Correct Order Number", "SLA Source"], errors="ignore")
                st.dataframe(display_enriched, use_container_width=True, hide_index=True)
                
            with sub_tab2:
                st.markdown("#### Validation Failures & Status Discrepancies (Separate Sheet)")
                if disc_df.empty:
                    st.success("No status discrepancies or validation failures identified!")
                else:
                    st.warning(f"Found {len(disc_df)} discrepancies/warnings.")
                    st.dataframe(disc_df, use_container_width=True, hide_index=True)
                    
            with sub_tab3:
                st.markdown("#### Country Pivot Summary & Email sharing")

                # Hardcoded default recipients per country. Currently the same
                # sample addresses for all three - update per country here as
                # real recipients are confirmed. The "From" address is NOT set
                # here; it's taken from the SMTP login/sender email configured
                # above.
                COUNTRY_EMAIL_DEFAULTS = {
                    "SG": {"to": "s.sudharsan@graas.ai", "cc": "yesuraja@graas.ai"},
                    "MY": {"to": "s.sudharsan@graas.ai", "cc": "yesuraja@graas.ai"},
                    "PH": {"to": "s.sudharsan@graas.ai", "cc": "yesuraja@graas.ai"},
                }

                country_sel = st.selectbox("Select Country", ["SG", "MY", "PH"])
                
                c_data = country_reports.get(country_sel, {})
                c_summary = c_data.get("summary_df", pd.DataFrame())
                c_pivot = c_data.get("pivot_df", pd.DataFrame())
                c_raw = c_data.get("raw_df", pd.DataFrame())
                
                if c_summary.empty and c_pivot.empty:
                    st.info(f"No order data found for country {country_sel}.")
                else:
                    st.markdown(f"##### Highlight Metrics for {country_sel}")
                    metrics_dict = c_summary.set_index("Metric")["Count"].to_dict() if not c_summary.empty else {}
                    
                    c1, c2, c3, c4, c5, c6 = st.columns(6)
                    c1.metric("Overdue (SLA breached)", metrics_dict.get("Overdue (SLA breached)", 0), delta="Breached" if metrics_dict.get("Overdue (SLA breached)", 0) > 0 else None, delta_color="inverse")
                    c2.metric("Handover Today (Today SLA)", metrics_dict.get("Handover today (Today SLA)", 0))
                    c3.metric("Order Status at New", metrics_dict.get("Order Status at New", 0))
                    c4.metric("Within SLA (Future)", metrics_dict.get("Within SLA (Future)", 0))
                    c5.metric("Not reflecting in OM", metrics_dict.get("Not reflecting in OM", 0))
                    c6.metric("Unpaid Orders", metrics_dict.get("Unpaid Orders", 0))
                    
                    st.markdown(f"##### Pivot Table: Channel & OMS Status vs Dates ({country_sel})")
                    st.dataframe(c_pivot, use_container_width=True, hide_index=True)
                    
                    with st.expander(f"View Raw Data ({country_sel})"):
                        st.dataframe(c_raw, use_container_width=True, hide_index=True)
                        
                    st.markdown("---")
                    st.markdown("##### 📧 Share Country Report with Seller Partner")
                    st.caption(f"From: uses the Sender/SMTP login email configured above. Subject: \"PUMA - {country_sel} Pending order Report on {ref_date_dmy}\".")
                    with st.expander(f"Send {country_sel} Report via Email", expanded=True):
                        defaults = COUNTRY_EMAIL_DEFAULTS.get(country_sel, {"to": "", "cc": ""})
                        to_val = st.text_input("To:", value=defaults["to"], key=f"to_{country_sel}")
                        cc_val = st.text_input("Cc:", value=defaults["cc"], key=f"cc_{country_sel}")
                        st.caption(f"Preview - this report will be sent **To: {to_val or '(none)'}  |  Cc: {cc_val or '(none)'}**")
                        
                        if st.button("Send Report", key=f"send_country_btn_{country_sel}", type="primary", use_container_width=True):
                            if not smtp_config.get("host"):
                                st.error("❌ SMTP config not found. Please configure the SMTP Email details in the setup section above.")
                            else:
                                with st.spinner(f"Generating and sending PUMA {country_sel} report..."):
                                    try:
                                        import io
                                        wb_to_send = excel_formatter.generate_excel_workbook(country_sel, c_raw, c_pivot, c_summary, ref_date_dmy)
                                        c_buf = io.BytesIO()
                                        wb_to_send.save(c_buf)
                                        excel_bytes = c_buf.getvalue()
                                        
                                        from email_sender import send_country_report_email
                                        success, msg = send_country_report_email(
                                            smtp_config=smtp_config,
                                            country=country_sel,
                                            to_email=to_val,
                                            cc_email=cc_val,
                                            excel_bytes=excel_bytes,
                                            ref_date_str=ref_date_dmy,
                                            pivot_df=c_pivot,
                                            summary_df=c_summary
                                        )
                                        if success:
                                            st.success(f"✅ {msg}")
                                        else:
                                            st.error(f"❌ {msg}")
                                    except Exception as e:
                                        st.error(f"❌ An error occurred: {str(e)}")

        elif mode == "tc_marketplace":
            # Mode 2: Market Place Order Check (File 2 + File 3)
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Missing Orders Report (Styled)</h3>', unsafe_allow_html=True)
            
            excel_bytes_data = excel_formatter.generate_fast_excel_bytes({"Missing Orders": enriched_df})
                
            st.download_button(
                label="📥 Download Missing Orders Report",
                data=excel_bytes_data,
                file_name=f"Missing Orders Report - {datetime.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_missing_orders"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("### Detailed Results")
            if enriched_df.empty:
                st.success("🎉 All Marketplace orders are reflected in the TC Order Report!")
            else:
                st.warning(f"⚠️ Found {len(enriched_df)} orders missing from TC Order Report.")
                st.dataframe(enriched_df, use_container_width=True, hide_index=True)

        elif mode == "order_flow_check":
            # Mode 4: Order Flow Check (File 2 + File 3 + File 4)
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Order Flow Check Report (Styled)</h3>', unsafe_allow_html=True)
            
            cols_to_drop_consolidated = ["Details", "Final Remarks", "_country", "Correct Order Number", "SLA Source", "SLA Date", "SLA", "sla_status"]
            cols_to_drop_mismatches = ["SLA Date", "SLA", "sla_status", "Details", "Final Remarks", "Correct Order Number", "SLA Source", "_country"]
            
            export_enriched_df = enriched_df.drop(columns=[c for c in cols_to_drop_consolidated if c in enriched_df.columns], errors="ignore")
            export_disc_df = disc_df.drop(columns=[c for c in cols_to_drop_mismatches if c in disc_df.columns], errors="ignore")

            excel_bytes_data = excel_formatter.generate_fast_excel_bytes({
                "Consolidated Report": export_enriched_df,
                "Missing & Mismatch Orders": export_disc_df
            })
                
            st.download_button(
                label="📥 Download Order Flow Check Report",
                data=excel_bytes_data,
                file_name=f"Order Flow Check Report - {datetime.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_order_flow_check"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("### Detailed Results")
            sub_tab1, sub_tab2 = st.tabs([
                "Consolidated Report", 
                "Missing & Mismatch Orders"
            ])
            
            with sub_tab1:
                st.markdown("#### Consolidated Report (Sheet 1)")
                st.dataframe(export_enriched_df, use_container_width=True, hide_index=True)
                
            with sub_tab2:
                st.markdown("#### Missing & Mismatch Orders (Sheet 2)")
                if export_disc_df.empty:
                    st.success("🎉 All Marketplace orders are reflected in TC and all TC orders match OMS!")
                else:
                    st.warning(f"⚠️ Found {len(export_disc_df)} missing or mismatch orders.")
                    st.dataframe(export_disc_df, use_container_width=True, hide_index=True)

        else:
            # Mode 3: Order Status Reconciliation (File 2 + File 4: TC + OMS)
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Status Reconciliation Report (Styled)</h3>', unsafe_allow_html=True)
            
            cols_to_drop_consolidated = ["Details", "Final Remarks", "_country", "Correct Order Number", "SLA Source", "SLA Date", "SLA", "sla_status"]
            cols_to_drop_mismatches = ["SLA Date", "SLA", "sla_status", "Details", "Final Remarks", "Correct Order Number", "SLA Source", "_country"]
            
            export_enriched_df = enriched_df.drop(columns=[c for c in cols_to_drop_consolidated if c in enriched_df.columns], errors="ignore")
            export_disc_df = disc_df.drop(columns=[c for c in cols_to_drop_mismatches if c in disc_df.columns], errors="ignore")

            excel_bytes_data = excel_formatter.generate_fast_excel_bytes({
                "Consolidated Report": export_enriched_df,
                "Status Mismatches": export_disc_df
            })
                
            st.download_button(
                label="📥 Download Status Reconciliation Report",
                data=excel_bytes_data,
                file_name=f"Status reconciliation Report - {datetime.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_status_reconciliation"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("### Detailed Results")
            sub_tab1, sub_tab2 = st.tabs([
                "Consolidated Report", 
                "Status Mismatches"
            ])
            
            with sub_tab1:
                st.markdown("#### Consolidated Report (Sheet 1)")
                st.dataframe(export_enriched_df, use_container_width=True, hide_index=True)
                
            with sub_tab2:
                st.markdown("#### Status Mismatches (Sheet 2)")
                if export_disc_df.empty:
                    st.success("🎉 No status mismatches found!")
                else:
                    st.warning(f"⚠️ Found {len(export_disc_df)} status mismatches.")
                    st.dataframe(export_disc_df, use_container_width=True, hide_index=True)
