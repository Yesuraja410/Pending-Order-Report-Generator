# -*- coding: utf-8 -*-
# VERSION: v1 - Email Sender via SMTP
import smtplib
import os
import io
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import excel_formatter

def test_smtp_connection(host, port, user, password, use_tls=True):
    """Test connection to the SMTP server."""
    try:
        if not host or not user or not password:
            return False, "SMTP configuration details are incomplete."
            
        port = int(port)
        if use_tls:
            server = smtplib.SMTP(host, port, timeout=10)
            server.ehlo()
            server.starttls()
            server.ehlo()
        else:
            server = smtplib.SMTP_SSL(host, port, timeout=10)
            
        server.login(user, password)
        server.quit()
        return True, "Successfully connected to SMTP server!"
    except Exception as e:
        return False, f"SMTP Connection failed: {str(e)}"

def send_seller_report_email(smtp_config, seller_name, recipient_email, seller_df, discrepancies_df=None):
    """
    Generate an Excel sheet for the seller, build a nice HTML summary, and send it via SMTP.
    """
    if seller_df.empty:
        return False, "No data available for this seller."

    if not recipient_email or "@" not in recipient_email:
        return False, f"Invalid or missing recipient email address: '{recipient_email}'"

    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("password")
    use_tls = smtp_config.get("use_tls", True)
    sender_email = smtp_config.get("sender_email", user)

    # == 1. Create the Excel Attachment in Memory ==============================
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        seller_df.to_excel(writer, sheet_name="Pending Orders", index=False)
        excel_formatter.format_data_sheet(writer.sheets["Pending Orders"], seller_df)
        
        if discrepancies_df is not None and not discrepancies_df.empty:
            # Filter discrepancies for this seller's orders
            order_ids = set(seller_df.iloc[:, 0].dropna().apply(lambda x: str(x).strip()).tolist())
            
            # Find the Order ID column in discrepancies
            disc_id_col = next((c for c in discrepancies_df.columns if "order" in c.lower()), "")
            if disc_id_col:
                seller_disc = discrepancies_df[discrepancies_df[disc_id_col].astype(str).str.strip().isin(order_ids)]
                if not seller_disc.empty:
                    seller_disc.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                    excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], seller_disc)
                    
    excel_data = excel_buffer.getvalue()

    # == 2. Build HTML Body ==================================================-
    total_orders = len(seller_df)
    
    # Check if there are urgent SLAs (e.g. within today/tomorrow)
    # Since we don't have standard date parsing here, we just display the first 5 pending orders as a preview
    preview_df = seller_df.head(10)
    
    # Convert preview to HTML table
    table_html = preview_df.to_html(classes="table", index=False, border=0)
    
    # Custom CSS style for email
    email_html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: #333333;
                background-color: #f9f9f9;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 700px;
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 30px;
                margin: 0 auto;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            }}
            .header {{
                border-bottom: 2px solid #0066cc;
                padding-bottom: 15px;
                margin-bottom: 20px;
            }}
            .header h2 {{
                color: #0066cc;
                margin: 0;
                font-size: 24px;
            }}
            .summary-box {{
                background-color: #f0f7ff;
                border-left: 4px solid #0066cc;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 0 4px 4px 0;
            }}
            .summary-title {{
                font-weight: bold;
                font-size: 16px;
                color: #004499;
                margin-bottom: 5px;
            }}
            .table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 15px;
                margin-bottom: 20px;
                font-size: 13px;
            }}
            .table th {{
                background-color: #f2f2f2;
                color: #444444;
                font-weight: bold;
                text-align: left;
                padding: 10px;
                border-bottom: 1px solid #dddddd;
            }}
            .table td {{
                padding: 10px;
                border-bottom: 1px solid #eeeeee;
            }}
            .footer {{
                font-size: 12px;
                color: #888888;
                border-top: 1px solid #e0e0e0;
                padding-top: 15px;
                margin-top: 30px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Daily Pending Order & SLA Report</h2>
                <p style="margin: 5px 0 0 0; color: #666666;">Store: <strong>{seller_name}</strong></p>
            </div>
            
            <p>Dear Seller partner,</p>
            <p>Please find attached the daily Pending Order and SLA Status validation report for your store. Kindly review the details to ensure prompt fulfillment and address any status mismatches highlighted.</p>
            
            <div class="summary-box">
                <div class="summary-title">Report Summary</div>
                Total Pending Orders: <strong>{total_orders}</strong><br>
                Please check the attached Excel sheet for the full list and any discrepancies identified.
            </div>

            <h3>Pending Orders Preview (Top 10 Rows)</h3>
            {table_html}

            <p style="font-size: 14px;"><strong>Note:</strong> A complete report with all order statuses and validation checks has been attached to this email as an Excel spreadsheet.</p>

            <p>Best regards,<br>
            <strong>Operations & Analytics Team</strong></p>
            
            <div class="footer">
                This is an automated report generated by the Status Validation Analyzer. Please do not reply directly to this email.
            </div>
        </div>
    </body>
    </html>
    """

    # == 3. Build Multipart Message ==========================================-
    msg = MIMEMultipart()
    msg["From"] = f"Operations Team <{sender_email}>"
    msg["To"] = recipient_email
    msg["Subject"] = f"Pending Order & SLA Report - {seller_name}"

    msg.attach(MIMEText(email_html, "html"))

    # Attach Excel file
    attachment = MIMEApplication(excel_data, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=f"Pending_Orders_Report_{seller_name.replace(' ', '_')}.xlsx")
    msg.attach(attachment)

    # == 4. Send Email via SMTP ==============================================-
    try:
        if use_tls:
            server = smtplib.SMTP(host, port, timeout=15)
            server.ehlo()
            server.starttls()
            server.ehlo()
        else:
            server = smtplib.SMTP_SSL(host, port, timeout=15)
            
        server.login(user, password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"
