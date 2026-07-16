# -*- coding: utf-8 -*-
import os
import io
import json
import requests
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Import formatting and processing modules
import excel_formatter
from order_processor import parse_country_and_channel, compute_sla_status

def load_config():
    config_path = "config.json"
    if not os.path.exists(config_path):
        raise FileNotFoundError("Configuration file 'config.json' not found. Please verify it exists.")
    with open(config_path, "r") as f:
        return json.load(f)

def get_mcp_token():
    token_path = r"C:\Users\Yesuraja\.gemini\antigravity\brain\abf6c61b-3147-45f7-90e4-f03458ddd1ae\scratch\token_data.json"
    if not os.path.exists(token_path):
        raise FileNotFoundError("MCP Access Token not found. Please log in first via the Streamlit dashboard sidebar.")
    with open(token_path, "r") as f:
        token_data = json.load(f)
    return token_data.get("access_token")

def fetch_pending_orders_db(access_token):
    mcp_url = "https://mcp.graas.ai/mcp/GED"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    init_payload = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {
                "name": "Antigravity-Automation",
                "version": "1.0.0"
            }
        }
    }
    
    try:
        r_init = requests.post(mcp_url, json=init_payload, headers=headers, timeout=15)
        if r_init.status_code == 401:
            raise PermissionError("Access token expired or unauthorized. Please re-authenticate via the Streamlit sidebar.")
    except Exception as e:
        raise ConnectionError(f"Failed to connect to MCP server: {e}")

    query = """
    SELECT 
        oi.ORDER_ID AS "orderID",
        oi.SELLER_ID AS "merchantID",
        oi.ORDER_CREATED_REPORT_TS AS "timeOrderCreated",
        oi.ORDER_ID AS "orderNumber",
        oi.ORDER_STATUS AS "orderStatus",
        oi.ORDER_ITEM_STATUS AS "orderItems.orderStatus",
        oi.PAYMENT_STATUS AS "paymentStatus",
        o.PAYMENT_METHOD AS "paymentMethods",
        oi.SELLER_SKU AS "orderItems.customSKU",
        NULL AS "shippingDeadLine",
        o.SHIPPING_METHOD AS "courierName",
        NULL AS "airwaybill",
        oi.SHIPPING_STATUS AS "omsStatus",
        oi.CHANNEL_NAME AS "storeName"
    FROM ORDER_ITEMS_METRICS oi
    LEFT JOIN ORDER_METRICS o ON oi.ORDER_ID = o.ORDER_ID
    WHERE oi.ORDER_STATUS IN ('UNPAID', 'INITIATED', 'PROCESSING', 'ACCEPTED', 'AWAITING_COLLECTION', 'AWAITING_SHIPMENT')
    """
    
    call_payload = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "query_data",
            "arguments": {
                "sql_query": query,
                "question": "Fetch pending orders list for automated reporting"
            }
        }
    }
    
    print("Executing database query against Snowflake views...")
    r = requests.post(mcp_url, json=call_payload, headers=headers, timeout=40)
    if r.status_code == 401:
        raise PermissionError("Access token expired or unauthorized. Please re-authenticate via the Streamlit sidebar.")
        
    if r.status_code != 200:
        raise ValueError(f"MCP Server returned status code {r.status_code}: {r.text}")
        
    res = r.json()
    if "error" in res:
        raise ValueError(f"MCP Server Error: {res['error'].get('message', 'Unknown error')}")
        
    result_data = res.get("result", {})
    structured = result_data.get("structuredContent", {})
    
    columns = []
    rows = []
    
    if structured:
        columns = structured.get("columns", [])
        rows = structured.get("rows", [])
    else:
        content_list = result_data.get("content", [])
        if content_list:
            try:
                parsed = json.loads(content_list[0].get("text", "{}"))
                if isinstance(parsed, dict) and "rows" in parsed:
                    columns = parsed.get("columns", [])
                    rows = parsed.get("rows", [])
            except Exception:
                pass
                
    if not rows:
        return pd.DataFrame()
        
    return pd.DataFrame(rows, columns=columns)

def calculate_sla_date(created_date_str):
    try:
        dt = pd.to_datetime(created_date_str)
        sla_dt = dt + pd.Timedelta(days=2)
        return sla_dt.strftime('%d-%m-%Y')
    except Exception:
        return ""

def send_report_email(smtp_config, store_name, recipient, excel_bytes, order_count):
    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("password")
    use_tls = smtp_config.get("use_tls", True)
    sender_email = smtp_config.get("sender_email", user)

    email_html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #333333; background-color: #f9f9f9; padding: 20px; }}
            .container {{ max-width: 650px; background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 30px; margin: 0 auto; }}
            .header {{ border-bottom: 2px solid #C00000; padding-bottom: 10px; margin-bottom: 20px; }}
            .header h2 {{ color: #C00000; margin: 0; }}
            .summary-box {{ background-color: #fff8f8; border-left: 4px solid #C00000; padding: 15px; margin-bottom: 20px; }}
            .footer {{ font-size: 11px; color: #888888; border-top: 1px solid #e0e0e0; padding-top: 10px; margin-top: 30px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Daily Pending Orders & SLA Report (Automated)</h2>
                <p>Store Code: <strong>{store_name}</strong></p>
            </div>
            <p>Dear Seller Partner,</p>
            <p>Please find attached the daily Pending Order and SLA Status validation report for your store.</p>
            
            <div class="summary-box">
                <strong>Report Summary:</strong><br>
                Total Pending Orders: <strong>{order_count}</strong><br>
                Please review the attached Excel spreadsheet which contains the summarized Pivot Tables, SLA indicators, and detailed order lists.
            </div>
            
            <p>Best regards,<br>
            <strong>Operations & Fulfillment Team</strong></p>
            
            <div class="footer">
                This is an automated report generated directly from our database. Please contact support if you have any questions.
            </div>
        </div>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = f"Operations Team <{sender_email}>"
    msg["To"] = recipient
    msg["Subject"] = f"Pending Order & SLA Report - {store_name} ({datetime.today().strftime('%d-%m-%Y')})"
    msg.attach(MIMEText(email_html, "html"))

    attachment = MIMEApplication(excel_bytes, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=f"Pending_Orders_Report_{store_name.replace(' ', '_')}.xlsx")
    msg.attach(attachment)

    if use_tls:
        server = smtplib.SMTP(host, port, timeout=20)
        server.ehlo()
        server.starttls()
        server.ehlo()
    else:
        server = smtplib.SMTP_SSL(host, port, timeout=20)
        
    server.login(user, password)
    server.sendmail(sender_email, recipient, msg.as_string())
    server.quit()

def run_automation():
    print("="*60)
    print("--- PUMA Seller Pending Orders Email Automation (MCP DB) ---")
    print("="*60)
    
    try:
        config = load_config()
        smtp_config = config["smtp"]
        seller_emails = config["seller_emails"]
        print("[OK] Configuration 'config.json' loaded successfully.")
    except Exception as e:
        print(f"[ERROR] Failed to load configuration: {e}")
        return

    try:
        access_token = get_mcp_token()
        print("[OK] MCP Access Token loaded successfully.")
    except Exception as e:
        print(f"[ERROR] {e}")
        return

    try:
        df_raw = fetch_pending_orders_db(access_token)
        print(f"[OK] Fetched {len(df_raw)} raw pending orders from database.")
    except Exception as e:
        print(f"[ERROR] Failed to retrieve orders from database: {e}")
        return

    if df_raw.empty:
        print("No pending orders retrieved from database. Processing completed.")
        return

    ref_date_str = datetime.today().strftime('%d-%m-%Y')
    ref_date_iso = datetime.today().strftime('%Y-%m-%d')
    print(f"Reference Date: {ref_date_str}")

    # Map database fields to application structure
    df_raw["Order ID"] = df_raw["orderID"].astype(str)
    df_raw["Store Name"] = df_raw["storeName"].fillna(df_raw["merchantID"]).astype(str)
    df_raw["SLA"] = df_raw["timeOrderCreated"].apply(calculate_sla_date)
    df_raw["Order Date"] = df_raw["timeOrderCreated"].astype(str).apply(lambda x: x.split("T")[0] if "T" in x else x)
    df_raw["OMS Order Status"] = df_raw["omsStatus"].fillna("UNKNOWN").astype(str)
    df_raw["Correct Order Number"] = df_raw["orderNumber"].astype(str)
    
    # Calculate SLA status
    df_raw["sla_status"] = df_raw["SLA"].apply(lambda x: compute_sla_status(x, ref_date_iso))

    # Group by Store and process reports
    unique_stores = df_raw["Store Name"].unique()
    success_count = 0
    fail_count = 0

    for store in unique_stores:
        print(f"\nProcessing report for Store: {store}...")
        store_df = df_raw[df_raw["Store Name"] == store].copy()
        
        country, channel = parse_country_and_channel(store)
        
        # Apply the required filter for the pivot and summary sheets:
        # Ignore only OMS shipped orders (keep unpaid orders)
        filtered_store_df = store_df[
            (store_df["OMS Order Status"].astype(str).str.lower() != "shipped")
        ].copy()
        
        # Generate Pivot Table
        if not filtered_store_df.empty:
            pivot_df = filtered_store_df.pivot_table(
                index=["omsStatus", "omsStatus"],  # Dummy level to match schema
                columns="Order Date",
                values="Correct Order Number",
                aggfunc="count",
                fill_value=0
            )
            pivot_df.index.names = ["Channel", "OMS Order Status"]
            pivot_df = pivot_df.reset_index()
            pivot_df["Channel"] = f"{channel} {country}"
            
            pivot_df["Grand Total"] = pivot_df.iloc[:, 2:].sum(axis=1)
            pivot_df.loc[len(pivot_df)] = ["Grand Total", ""] + pivot_df.iloc[:, 2:].sum(axis=0).tolist()

            new_cols = []
            for col in pivot_df.columns:
                try:
                    dt = pd.to_datetime(col)
                    new_cols.append(dt.strftime('%d-%m-%Y'))
                except Exception:
                    new_cols.append(col)
            pivot_df.columns = new_cols
        else:
            # Empty pivot structure matching schema
            pivot_df = pd.DataFrame(columns=["Channel", "OMS Order Status", "Grand Total"])
            pivot_df.loc[0] = ["Grand Total", "", 0]

        # Generate summary metrics
        summary_metrics = [
            {"Metric": "Overdue (SLA breached)", "Count": int((filtered_store_df["sla_status"] == "Breached").sum()) if not filtered_store_df.empty else 0},
            {"Metric": "Handover today (Today SLA)", "Count": int((filtered_store_df["sla_status"] == "Today").sum()) if not filtered_store_df.empty else 0},
            {"Metric": "Order Status at New", "Count": int((filtered_store_df["OMS Order Status"].astype(str).str.lower() == "new").sum()) if not filtered_store_df.empty else 0},
            {"Metric": "Within SLA (Future)", "Count": int((filtered_store_df["sla_status"] == "Future").sum()) if not filtered_store_df.empty else 0},
            {"Metric": "Not reflecting in OM", "Count": int((filtered_store_df["OMS Order Status"] == "Not in OMS").sum()) if not filtered_store_df.empty else 0},
            {"Metric": "Unpaid Orders", "Count": int(filtered_store_df["Final Remarks"].astype(str).str.contains("Unpaid", case=False).sum()) if not filtered_store_df.empty else 0}
        ]
        summary_df = pd.DataFrame(summary_metrics)

        # Drop helper columns from raw sheet data
        cols_to_drop = ["Correct Order Number", "Order Date", "orderID", "merchantID", "timeOrderCreated", "orderNumber", "orderStatus", "orderItems.orderStatus", "paymentStatus", "paymentMethods", "orderItems.customSKU", "shippingDeadLine", "courierName", "airwaybill", "omsStatus", "storeName"]
        raw_export_df = store_df.drop(columns=[c for c in cols_to_drop if c in store_df.columns])

        # Generate styled Excel sheets in memory
        try:
            wb = excel_formatter.generate_excel_workbook(country, raw_export_df, pivot_df, summary_df, ref_date_str)
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_bytes = excel_buffer.getvalue()
        except Exception as e:
            print(f"   [FAILED] Excel report formatting error for {store}: {e}")
            fail_count += 1
            continue

        # Resolve recipient email address
        recipient = seller_emails.get(store, seller_emails.get("default"))
        if not recipient or "@" not in recipient:
            print(f"   [FAILED] No valid email address configured for store {store}.")
            fail_count += 1
            continue

        # Send email
        try:
            print(f"   Sending report to {recipient} via email...")
            send_report_email(smtp_config, store, recipient, excel_bytes, len(store_df))
            print(f"   [SUCCESS] Daily report sent to {store} partner ({recipient}) successfully.")
            success_count += 1
        except Exception as e:
            print(f"   [FAILED] SMTP email delivery error for {store}: {e}")
            fail_count += 1

    print("\n" + "="*60)
    print("--- Automation Execution Summary ---")
    print(f"Total Stores Processed: {len(unique_stores)}")
    print(f"Emails Sent Successfully: {success_count}")
    print(f"Errors/Failures: {fail_count}")
    print("="*60 + "\n")

if __name__ == "__main__":
    run_automation()
