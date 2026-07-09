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
st.write("Upload the daily SLA Report, TC Report (All file), and OMS Report (Sales Order file) in the sidebar to run validations and email reports directly to the sellers.")

# == Sidebar ==================================================================-
with st.sidebar:
    st.markdown("## Configuration")
    st.markdown("Upload the daily reports below:")
    
    pending_source = st.selectbox(
        "1. Pending Order Report Source",
        ["Google Sheet Link", "Upload File", "Fetch from DB (MCP Server)"]
    )
    
    order_pending = None
    if pending_source == "Google Sheet Link":
        gsheet_url = st.text_input("Pending Order Report (Google Sheet Link)", placeholder="https://docs.google.com/spreadsheets/d/...")
        order_pending = gsheet_url if gsheet_url.strip() else None
    elif pending_source == "Upload File":
        uploaded_pending = st.file_uploader("Upload Pending Order Report (SLA Report)", type=["xlsx","xls","csv"], key="order_pending_upload")
        order_pending = uploaded_pending
    else:
        # Fetch from DB (MCP Server)
        token_file = r"C:\Users\Yesuraja\.gemini\antigravity\brain\abf6c61b-3147-45f7-90e4-f03458ddd1ae\scratch\token_data.json"
        has_token = os.path.exists(token_file)
        
        if not has_token:
            st.warning("⚠️ MCP Authorization required.")
            client_id = "bVsPfvtJYPMP_MblEbImVw"
            redirect_uri = "https://oauth.pstmn.io/v1/callback"
            
            if "pkce_verifier" not in st.session_state:
                import secrets, hashlib, base64
                verifier = secrets.token_urlsafe(64)
                sha = hashlib.sha256(verifier.encode('utf-8')).digest()
                challenge = base64.urlsafe_b64encode(sha).decode('utf-8').replace('=', '')
                st.session_state["pkce_verifier"] = verifier
                st.session_state["pkce_challenge"] = challenge
                st.session_state["pkce_state"] = secrets.token_hex(16)
                
            auth_url = (
                "https://mcp.graas.ai/authorize?"
                f"response_type=code&"
                f"client_id={client_id}&"
                f"redirect_uri={redirect_uri}&"
                f"scope=read:data%20read:schema&"
                f"state={st.session_state['pkce_state']}&"
                f"code_challenge={st.session_state['pkce_challenge']}&"
                f"code_challenge_method=S256"
            )
            
            st.markdown(f"[🔗 Click to Authorize MCP]({auth_url})")
            redirect_url_input = st.text_input("Paste the redirected URL here:", key="mcp_redirect_url")
            if st.button("Complete Authorization"):
                if redirect_url_input.strip():
                    try:
                        from urllib.parse import urlparse, parse_qs
                        parsed = urlparse(redirect_url_input.strip())
                        params = parse_qs(parsed.query)
                        code = params.get("code", [None])[0]
                        if not code:
                            st.error("No authorization code found in the URL. Please verify.")
                        else:
                            token_url = "https://mcp.graas.ai/token"
                            payload = {
                                "grant_type": "authorization_code",
                                "client_id": client_id,
                                "code": code,
                                "redirect_uri": redirect_uri,
                                "code_verifier": st.session_state["pkce_verifier"]
                            }
                            r = requests.post(token_url, data=payload, headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=15)
                            if r.status_code == 200:
                                token_data = r.json()
                                os.makedirs(os.path.dirname(token_file), exist_ok=True)
                                with open(token_file, "w") as tf:
                                    json.dump(token_data, tf)
                                st.success("Authorization successful!")
                                st.rerun()
                            else:
                                st.error(f"Failed to fetch token: {r.text}")
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
        else:
            st.success("✅ MCP Authorization active.")
            if st.button("Disconnect / Re-authorize"):
                if os.path.exists(token_file):
                    os.remove(token_file)
                st.session_state.pop("pkce_verifier", None)
                st.rerun()
            order_pending = "mcp"
        
    order_tc = st.file_uploader("2. TC Report (All file)", type=["xlsx","xls","csv"], key="order_tc")
    order_oms = st.file_uploader("3. OMS Report (Sales Order file)", type=["xlsx","xls","csv"], key="order_oms")
    seller_contacts = st.file_uploader("4. Seller Contact List (Optional)", type=["xlsx","xls","csv"], key="seller_contacts")
    
    st.markdown("---")
    run_btn = st.button("Run Order Validation", use_container_width=True, type="primary")

# == Main Screen ===============================================================
# Check if files are uploaded
if not (order_tc and order_oms):
    st.info("Please upload at least the TC Report (All file) and OMS Report (Sales Order file) in the sidebar to get started.")
else:
    # Display active mode banner
    if order_pending == "mcp":
        st.success("🎯 **MCP Database Automation Mode**: Pending Order Report will be queried directly from PUMA Database, enriched with TC Report, and reconciled against OMS Report.")
    elif order_pending:
        st.success("🎯 **Full Validation Mode (GSheet SLA)**: Pending Order Report, TC Report, and OMS Report are ready for processing.")
    else:
        st.success("🎯 **Validation Mode (TC Pending)**: Pending orders will be extracted from TC Report (status NEW, READY TO SHIP, ACCEPTED/PICKED) and reconciled against OMS Report.")

    # Trigger validation either by clicking sidebar button or main screen button
    if run_btn or st.button("Run Validation & Analysis", type="primary", use_container_width=True):
        with st.spinner("Processing reports and running validations..."):
            try:
                res = process_and_validate_orders(order_pending, order_tc, order_oms, seller_contacts)
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

        # Display metrics
        st.markdown("### Key Metrics")
        if has_pending:
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
                # Sheet 1: SLA Report
                enriched_df.to_excel(writer, sheet_name="SLA Report", index=False)
                # Sheet 2: Status Discrepancies
                disc_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                
                # Format using excel_formatter
                excel_formatter.format_data_sheet(writer.sheets["SLA Report"], enriched_df)
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
                st.dataframe(enriched_df, use_container_width=True, hide_index=True)
                
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
                    
                    c1, c2, c3, c4, c5 = st.columns(5)
                    c1.metric("Overdue (SLA breached)", metrics_dict.get("Overdue (SLA breached)", 0), delta="Breached" if metrics_dict.get("Overdue (SLA breached)", 0) > 0 else None, delta_color="inverse")
                    c2.metric("Handover Today (Today SLA)", metrics_dict.get("Handover today (Today SLA)", 0))
                    c3.metric("Order Status at New", metrics_dict.get("Order Status at New", 0))
                    c4.metric("Within SLA (Future)", metrics_dict.get("Within SLA (Future)", 0))
                    c5.metric("Not reflecting in OM", metrics_dict.get("Not reflecting in OM", 0))
                    
                    # 2. Display Pivot Table
                    st.markdown(f"##### Pivot Table: Channel & OMS Status vs Dates ({country_sel})")
                    st.dataframe(c_pivot, use_container_width=True, hide_index=True)
                    
                    # 3. Display Raw Data
                    with st.expander(f"View Raw Data ({country_sel})"):
                        st.dataframe(c_raw, use_container_width=True, hide_index=True)
                        
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

