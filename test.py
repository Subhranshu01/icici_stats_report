import os
import requests
import pandas as pd
import configparser
from datetime import datetime, timedelta
from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage


FILE_NAME = "dynatrace_metrics.xlsx"

# Fetch metrics and update Excel
def fetch_and_store_metrics(controller_name, url, api_token):
    yesterday_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    sheet_name = controller_name.lower()

    headers = {
        "Authorization": f"Api-Token {api_token}",
        "accept": "application/json"
    }
    response = requests.get(url, headers=headers)
    print(f"📡 {controller_name}: Status Code {response.status_code}")
    if response.status_code != 200:
        print("❌ Error:", response.text)
        return

    json_data = response.json()
    result = json_data.get("result", [])
    data_dict = {}

    for metric in result:
        metric_id = metric["metricId"]
        for entry in metric["data"]:
            method_name = entry["dimensionMap"]["dt.entity.service_method.name"]
            value = entry["values"][0]
            if method_name not in data_dict:
                data_dict[method_name] = {}
            if "count.total" in metric_id:
                data_dict[method_name]["total_hits"] = int(value)
            elif "errors.server.count" in metric_id:
                data_dict[method_name]["failure_count"] = int(value)
            elif ":avg" in metric_id:
                data_dict[method_name]["avg"] = round(value / 1_000_000, 2)
            elif "percentile(90.0)" in metric_id:
                data_dict[method_name]["p90"] = round(value / 1_000_000, 2)
            elif "percentile(95.0)" in metric_id:
                data_dict[method_name]["p95"] = round(value / 1_000_000, 2)
            elif "percentile(99.0)" in metric_id:
                data_dict[method_name]["p99"] = round(value / 1_000_000, 2)
            elif "errors.server.rate" in metric_id:
                data_dict[method_name]["failure_rate"] = round(value, 2)

    records = []
    for method, values in data_dict.items():
        records.append({
            "Date": yesterday_str,
            "request": method,
            "total hits": values.get("total_hits", 0),
            "failure count": values.get("failure_count", 0),
            "average": values.get("avg", 0),
            "p90": values.get("p90", 0),
            "p95": values.get("p95", 0),
            "p99": values.get("p99", 0),
            "failure rate": values.get("failure_rate", 0)
        })

    df = pd.DataFrame(records)

    try:
        if os.path.exists(FILE_NAME):
            with pd.ExcelFile(FILE_NAME, engine="openpyxl") as xls:
                existing_sheets = xls.sheet_names
                existing_df = pd.read_excel(xls, sheet_name=sheet_name) if sheet_name in existing_sheets else pd.DataFrame()
            combined_df = pd.concat([df, existing_df], ignore_index=True)
        else:
            combined_df = df
    except Exception:
        print("⚠️ Corrupted workbook. Recreating it.")
        os.remove(FILE_NAME)
        combined_df = df

    try:
        writer = pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="replace")
    except FileNotFoundError:
        writer = pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w")

    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()
    print(f"✅ Sheet '{sheet_name}' updated.")

# Send email with Excel file
def send_email_report():
    EMAIL_USER = os.environ["EMAIL_USER"]
    EMAIL_PASS = os.environ["EMAIL_PASS"]
    SMTP_SERVER = os.environ["EMAIL_HOST"]
    SMTP_PORT = int(os.environ["EMAIL_PORT"])
    TO_EMAIL = os.environ["EMAIL_TO"]

    msg = EmailMessage()
    msg["Subject"] = "📊 Dynatrace Daily Metrics Report"
    msg["From"] = EMAIL_USER
    msg["To"] = TO_EMAIL
    msg.set_content("Hi,\n\nPlease find attached the latest Dynatrace metrics workbook.\n\nRegards,\nAutomated Bot")

    with open(FILE_NAME, "rb") as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=FILE_NAME)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)
        print("📧 Email sent successfully.")
    except Exception as e:
        print("❌ Email failed:", e)

# 🔄 Run job
if __name__ == "__main__":
    API_TOKEN = os.environ["API_TOKEN"]
    LoginController_url = os.environ["LOGINCONTROLLER_URL"]
    MotorInsurance_url = os.environ["MOTORINSURANCE_URL"]

    fetch_and_store_metrics("LoginController", LoginController_url, API_TOKEN)
    fetch_and_store_metrics("LoanController", MotorInsurance_url, API_TOKEN)
    send_email_report()
