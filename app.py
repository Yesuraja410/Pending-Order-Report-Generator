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

# Custom title and introduction
st.title("Pending Order SLA Enrichment & OMS Status Validation")
st.write("Upload the daily SLA Report (GSheet Link), Marketplace Order Reports (TC Reports), and OMS Report (Sales Order file) in the sidebar to run validations and email reports directly to the sellers.")

# == Sidebar ==================================================================-
with st.sidebar:
    st.markdown("## Configuration")
    st.markdown("Upload the daily reports below:")
    
    gsheet_url = st.text_input("1. Pending Order Report (Google Sheet Link)", placeholder="https://docs.google.com/spreadsheets/d/...")
    order_pending = gsheet_url if gsheet_url.strip() else None
        
    order_tc = st.file_uploader("2. TC Order Report", type=["xlsx","xls","csv"], key="order_tc")
    marketplace_reports = st.file_uploader("3. Market Place Reports (Multiple Upload)", type=["xlsx","xls","csv"], accept_multiple_files=True, key="marketplace_reports")
    order_oms = st.file_uploader("4. OMS Report (Sales Order file)", type=["xlsx","xls","csv"], key="order_oms")
    
    st.markdown("---")
    run_btn = st.button("Run Order Validation", use_container_width=True, type="primary")

# == Main Screen ===============================================================
# Check if files are uploaded
has_pending = order_pending is not None
has_tc = (order_tc is not None) or (marketplace_reports is not None and len(marketplace_reports) > 0)
has_oms = order_oms is not None

is_valid_combo = False
mode_desc = ""

# 1. Mode 1: GSheet + OMS Report Alone
if has_pending and has_oms and not has_tc:
    is_valid_combo = True
    mode_desc = "GSheet + OMS Validation Mode"
# 2. Mode 2: TC Order Report + Marketplace Reports alone
elif has_tc and not has_pending and not has_oms:
    is_valid_combo = True
    mode_desc = "TC + Marketplace Reconciliation Mode"
# 3. Mode 3/4: TC/Marketplace + OMS (with or without GSheet)
elif has_tc and has_oms:
    is_valid_combo = True
    if has_pending:
        mode_desc = "Full Validation Mode (GSheet + TC + OMS)"
    else:
        mode_desc = "TC + OMS Validation Mode"

if not is_valid_combo:
    st.info(
        "💡 **Please upload one of the following combinations to start validation:**\n\n"
        "1. **GSheet + OMS Mode**: Upload GSheet Link (1) and OMS Report (4).\n"
        "2. **TC + Marketplace Mode**: Upload TC Order Report (2) and Market Place Reports (3).\n"
        "3. **TC + OMS Mode**: Upload TC Order Report (2) or Market Place Reports (3), and OMS Report (4) (optionally add GSheet Link for Full SLA Enrichment)."
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
        
        has_pending = not enriched_df.empty
        is_reconciliation_mode = "all_imported_to_tc" in summary

        # Display metrics
        st.markdown("### Key Metrics")
        if is_reconciliation_mode:
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Marketplace Orders", summary["total_pending_orders"])
            m2.metric("Reflected in TC", summary["pushed_count"])
            m3.metric("Missing from TC", summary["not_pushed_count"])
            
            if summary["all_imported_to_tc"]:
                st.success("🎉 **All Marketplace orders are successfully imported to TC!**")
            else:
                st.error(f"⚠️ **Found {summary['not_pushed_count']} orders missing from TC!**")
        elif has_pending:
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total Pending Orders", summary["total_pending_orders"])
            m2.metric("Successfully Pushed", summary["pushed_count"])
            m3.metric("Not Pushed to OMS", summary["not_pushed_count"])
            m4.metric("Unpaid Orders", summary["unpaid_count"])
            m5.metric("Status Discrepancies", summary["total_discrepancies"], 
                      delta=summary["total_discrepancies"] if summary["total_discrepancies"] > 0 else None, 
                      delta_color="inverse")
        else:
            m1, = st.columns(1)
            m1.metric("Status Discrepancies Found", summary["total_discrepancies"], 
                      delta=summary["total_discrepancies"] if summary["total_discrepancies"] > 0 else None, 
                      delta_color="inverse")
        
        if has_pending:
            # Download Section matching screenshot
            # Country-specific reports download container
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
                # Sheet 1: SLA Report (dropping Correct Order Number and SLA Source as requested)
                export_enriched_df = enriched_df.drop(columns=["Correct Order Number", "SLA Source"], errors="ignore")
                export_enriched_df.to_excel(writer, sheet_name="SLA Report", index=False)
                # Sheet 2: Status Discrepancies
                disc_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                
                # Format using excel_formatter
                excel_formatter.format_data_sheet(writer.sheets["SLA Report"], export_enriched_df)
                excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], disc_df)
                
                # Country-specific styled Summary and Data sheets
                for country in ["SG", "MY", "PH"]:
                    c_data = country_reports.get(country, {})
                    pivot_df = c_data.get("pivot_df", pd.DataFrame())
                    raw_df = c_data.get("raw_df", pd.DataFrame())
                    summary_df = c_data.get("summary_df", pd.DataFrame())
                    
                    if not raw_df.empty:
                        excel_formatter.add_country_sheets_to_workbook(
                            writer.book, country, raw_df, pivot_df, summary_df, ref_date_dmy
                        )
                
                # Reorder sheets to show all summaries first, then main reports, then data sheets
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
        else:
            # Reconciliation Mode: Single styled Discrepancy Report download container
            st.markdown('<div class="download-container">', unsafe_allow_html=True)
            st.markdown('<h3 class="download-header">📥 Download Status Discrepancies Report (Styled)</h3>', unsafe_allow_html=True)
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                disc_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], disc_df)
                
            st.download_button(
                label="📥 Download Discrepancy Report",
                data=excel_buffer.getvalue(),
                file_name=f"Status Discrepancies Report - {datetime.today().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_discrepancies"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Display layout of tables
        st.markdown("### Detailed Results")
        if has_pending:
            sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs([
                "SLA Report", 
                "Status Discrepancies", 
                "Seller Grouping & Email Center",
                "Country Pivots & Highlights"
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
                    
            with sub_tab4:
                st.markdown("#### Country Pivots & Highlight Metrics")
                country_sel = st.selectbox("Select Country to View Summary & Pivot", ["SG", "MY", "PH"])
                
                c_data = country_reports.get(country_sel, {})
                c_summary = c_data.get("summary_df", pd.DataFrame())
                c_pivot = c_data.get("pivot_df", pd.DataFrame())
                c_raw = c_data.get("raw_df", pd.DataFrame())
                
                if c_summary.empty and c_pivot.empty:
                    st.info(f"No order data found for country {country_sel}.")
                else:
                    # 1. Highlight metrics using st.columns & metric cards
                    st.markdown(f"##### Highlight Metrics for {country_sel}")
                    
                    # Fetch count for each metric
                    metrics_dict = c_summary.set_index("Metric")["Count"].to_dict() if not c_summary.empty else {}
                    
                    c1, c2, c3, c4, c5, c6 = st.columns(6)
                    c1.metric("Overdue (SLA breached)", metrics_dict.get("Overdue (SLA breached)", 0), delta="Breached" if metrics_dict.get("Overdue (SLA breached)", 0) > 0 else None, delta_color="inverse")
                    c2.metric("Handover Today (Today SLA)", metrics_dict.get("Handover today (Today SLA)", 0))
                    c3.metric("Order Status at New", metrics_dict.get("Order Status at New", 0))
                    c4.metric("Within SLA (Future)", metrics_dict.get("Within SLA (Future)", 0))
                    c5.metric("Not reflecting in OM", metrics_dict.get("Not reflecting in OM", 0))
                    c6.metric("Unpaid Orders", metrics_dict.get("Unpaid Orders", 0))
                    
                    # 2. Display Pivot Table
                    st.markdown(f"##### Pivot Table: Channel & OMS Status vs Dates ({country_sel})")
                    st.dataframe(c_pivot, use_container_width=True, hide_index=True)
                    
                    # 3. Display Raw Data
                    with st.expander(f"View Raw Data ({country_sel})"):
                        st.dataframe(c_raw, use_container_width=True, hide_index=True)
                        
                    # 4. Email Sending Section
                    st.markdown("---")
                    st.markdown("##### 📧 Share Country Report with Seller Partner")
                    with st.expander(f"Send {country_sel} Report via Email", expanded=False):
                        to_val = st.text_input("To:", value="yesuraja@graas.ai", key=f"to_{country_sel}")
                        cc_val = st.text_input("Cc:", value="sudharsan.s@graas.ai", key=f"cc_{country_sel}")
                        
                        if st.button("Send Report", key=f"send_country_btn_{country_sel}", type="primary", use_container_width=True):
                            # Load SMTP config
                            smtp_config = st.session_state.get("smtp_config", {})
                            if not smtp_config or not smtp_config.get("host"):
                                import json
                                try:
                                    with open("config.json", "r") as config_file:
                                        cfg = json.load(config_file)
                                        smtp_config = cfg.get("smtp_config", {})
                                except Exception:
                                    pass
                            
                            if not smtp_config or not smtp_config.get("host"):
                                st.error("❌ SMTP config not found. Please configure the SMTP Email details in the 'Email Setup' tab.")
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
                                            ref_date_str=ref_date_dmy
                                        )
                                        if success:
                                            st.success(f"✅ {msg}")
                                        else:
                                            st.error(f"❌ {msg}")
                                    except Exception as e:
                                        st.error(f"❌ An error occurred: {str(e)}")
                        
            with sub_tab3:
                st.markdown("#### SMTP Email Configuration")
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
                st.markdown("#### Send Daily Report to Sellers")
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
        else:
            # Reconciliation Mode: Only display status discrepancies
            st.markdown("#### Validation Failures & Status Discrepancies")
            if disc_df.empty:
                st.success("No status discrepancies or validation failures identified!")
            else:
                st.warning(f"Found {len(disc_df)} discrepancies/warnings.")
                st.dataframe(disc_df, use_container_width=True, hide_index=True)

