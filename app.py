# -*- coding: utf-8 -*-
import io
import os
import gc
import tempfile
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

inject_css()

# Custom title and introduction
st.title("Pending Order SLA Enrichment & OMS Status Validation")
st.write("Upload the daily SLA Report, TC Report (All file), and OMS Report (Sales Order file) in the sidebar to run validations and email reports directly to the sellers.")

# == Sidebar ==================================================================-
with st.sidebar:
    st.markdown("## Configuration")
    st.markdown("Upload the daily reports below:")
    
    order_pending = st.file_uploader("1. Pending Order Report (SLA Report)", type=["xlsx","xls","csv"], key="order_pending")
    order_tc = st.file_uploader("2. TC Report (All file)", type=["xlsx","xls","csv"], key="order_tc")
    order_oms = st.file_uploader("3. OMS Report (Sales Order file)", type=["xlsx","xls","csv"], key="order_oms")
    seller_contacts = st.file_uploader("4. Seller Contact List (Optional)", type=["xlsx","xls","csv"], key="seller_contacts")
    
    st.markdown("---")
    run_btn = st.button("Run Order Validation", use_container_width=True, type="primary")

# == Main Screen ===============================================================
# Check if files are uploaded
if not (order_pending and order_tc and order_oms):
    st.info("Please upload the Pending Order Report, TC Report, and OMS Report in the sidebar to get started.")
else:
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
        st.markdown("### Key Metrics")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Pending Orders", summary["total_pending_orders"])
        m2.metric("Successfully Pushed", summary["pushed_count"])
        m3.metric("Not Pushed to OMS", summary["not_pushed_count"])
        m4.metric("Unpaid Orders", summary["unpaid_count"])
        m5.metric("Status Discrepancies", summary["total_discrepancies"], 
                  delta=summary["total_discrepancies"] if summary["total_discrepancies"] > 0 else None, 
                  delta_color="inverse")
        
        # Download Section matching screenshot
        st.markdown('<div class="download-container">', unsafe_allow_html=True)
        st.markdown('<h3 class="download-header">📥 Download Validation Reports</h3>', unsafe_allow_html=True)
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            # Sheet 1: SLA Report
            enriched_df.to_excel(writer, sheet_name="SLA Report", index=False)
            # Sheet 2: Status Discrepancies
            disc_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
        
        st.download_button(
            label="📥 Download Detailed Excel QC Report",
            data=excel_buffer.getvalue(),
            file_name=f"SLA_Validation_Report_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_consolidated"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Display layout of tables
        st.markdown("### Detailed Results")
        sub_tab1, sub_tab2, sub_tab3 = st.tabs([
            "SLA Report", 
            "Status Discrepancies", 
            "Seller Grouping & Email Center"
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
