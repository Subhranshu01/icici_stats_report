import os
import requests
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
import smtplib
import pytz
from email.message import EmailMessage
from openpyxl.styles import Alignment, PatternFill

india_tz = pytz.timezone("Asia/Kolkata")
now_ist = datetime.now(india_tz)
FILE_NAME = "Product Category wise Internal APIs Performance Report.xlsx"

DEBUG = os.environ.get("DEBUG", "0") == "1"

def is_new_api_response(json_data):
    return any(
        "Dimension" in entry.get("dimensionMap", {}) or "dimension" in entry.get("dimensionMap", {})
        for metric in json_data.get("result", [])
        for entry in metric.get("data", [])
    )

def _assign_field(dct, method, key, new_value):
    """Assign new_value to dct[method][key] without overwriting a non-zero existing value with zero.
    Always overwrite None or missing. Print debug if enabled.
    """
    existing = dct[method].get(key)
    # if existing is not None and existing != 0 and new_value == 0 --> skip overwrite
    if existing is not None and existing != 0 and (new_value == 0 or new_value == 0.0):
        if DEBUG:
            print(f"⚠️ skip overwrite {method}.{key}: existing={existing} new={new_value}")
        return
    dct[method][key] = new_value
    if DEBUG:
        print(f"✅ set {method}.{key} = {new_value}")

def apply_formatting(file_path):
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    wb = load_workbook(file_path)

    for sheet in wb.worksheets:
        header_row = next(sheet.iter_rows(min_row=1, max_row=1))
        col_map = {cell.value: cell.column for cell in header_row}

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if cell.value is not None:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

            # p90
            if col_map.get("p90"):
                cell_p90 = row[col_map.get("p90") - 1]
                if isinstance(cell_p90.value, (int, float)) and cell_p90.value > 3:
                    cell_p90.fill = red_fill

            # p95
            if col_map.get("p95"):
                cell_p95 = row[col_map.get("p95") - 1]
                if isinstance(cell_p95.value, (int, float)) and cell_p95.value > 3:
                    cell_p95.fill = red_fill

            # failure rate
            if col_map.get("failure rate"):
                cell_fr = row[col_map.get("failure rate") - 1]
                if isinstance(cell_fr.value, (int, float)) and cell_fr.value > 10:
                    cell_fr.fill = red_fill

        for col in sheet.columns:
            max_length = max((len(str(cell.value)) if cell.value else 0) for cell in col)
            adjusted_width = max_length + 2
            col_letter = col[0].column_letter
            sheet.column_dimensions[col_letter].width = adjusted_width

    wb.save(file_path)
    print("🎨 Formatting + 🔴 highlights applied to all sheets.")

def fetch_and_store_metrics(controller_name, url, api_token):
    yesterday_str = (now_ist - timedelta(days=1)).strftime("%Y-%m-%d")
    sheet_name = controller_name

    headers = {
        "Authorization": f"Api-Token {api_token}",
        "accept": "application/json"
    }

    try:
        response = requests.get(url, headers=headers, timeout=30)
    except Exception as e:
        print(f"📡 {controller_name}: Request failed:", e)
        return

    print(f"📡 {controller_name}: Status Code {response.status_code}")
    if response.status_code != 200:
        print("❌ Error:", response.text)
        return

    json_data = response.json()
    is_new = is_new_api_response(json_data)
    if DEBUG:
        print(f"🧪 is_new={is_new} for {controller_name}")

    result = json_data.get("result", [])
    data_dict = {}

    for metric in result:
        metric_id = metric.get("metricId", "")
        for entry in metric.get("data", []):
            dim_map = entry.get("dimensionMap", {}) or {}
            # choose method name based on schema
            if is_new:
                method_name = dim_map.get("Dimension") or dim_map.get("dimension")
            else:
                method_name = dim_map.get("dt.entity.service_method.name") or dim_map.get("dt.entity.service_method") or None

            if method_name:
                method_name = str(method_name).strip()
            else:
                # fallback: pick first value from dimensionMap if available
                if dim_map:
                    first_val = next(iter(dim_map.values()))
                    method_name = str(first_val).strip()
                else:
                    method_name = "unknown"

            values_list = entry.get("values", []) or []
            value = values_list[0] if values_list else 0

            if DEBUG:
                print(f"📛 Metric ID: {metric_id} → Method: {method_name} → Value: {value}")

            if method_name not in data_dict:
                data_dict[method_name] = {}

            # New API detection branch: use substrings that appear in new metricId values
            if is_new:
                # total hits
                if "portfoliotrackcontroller_total_count" in metric_id or "total_count" in metric_id:
                    _assign_field(data_dict, method_name, "total_hits", int(value))
                    continue
                # failure count
                if "failed_req" in metric_id or "failed" in metric_id or "errors.server.count" in metric_id:
                    _assign_field(data_dict, method_name, "failure_count", int(value))
                    continue
                # avg response time (Dynatrace sometimes reports in microseconds)
                if "calc:service.portfoliotrackcontroller_avg_responsetime:splitBy(Dimension):sort(value(auto,descending))):names" in metric_id :
                    _assign_field(data_dict, method_name, "avg", round(value / 1_000_000, 2))
                    continue
                # percentiles
                if "(calc:service.portfoliotrackcontroller_avg_responsetime:splitBy(Dimension):percentile(90.0):sort(value(percentile(90.0)" in metric_id :
                    _assign_field(data_dict, method_name, "p90", round(value / 1_000_000, 2))
                    continue
                if "(calc:service.portfoliotrackcontroller_avg_responsetime:splitBy(Dimension):percentile(95.0):sort(value(percentile(95.0)" in metric_id :
                    _assign_field(data_dict, method_name, "p95", round(value / 1_000_000, 2))
                    continue
                if "percentile(99.0)" in metric_id or "p99" in metric_id:
                    _assign_field(data_dict, method_name, "p99", round(value / 1_000_000, 2))
                    continue
                # failure rate
                if "errors.server.rate" in metric_id or "failure_rate" in metric_id:
                    _assign_field(data_dict, method_name, "failure_rate", round(value, 2))
                    continue

                # fallback: try generic patterns also used by old API
                if "total_count" in metric_id or "count.total" in metric_id:
                    _assign_field(data_dict, method_name, "total_hits", int(value))
                elif "errors.server.count" in metric_id:
                    _assign_field(data_dict, method_name, "failure_count", int(value))
                elif ":avg" in metric_id:
                    _assign_field(data_dict, method_name, "avg", round(value / 1_000_000, 2))
            else:
                # Old API branch
                if "count.total" in metric_id or "total_count" in metric_id:
                    _assign_field(data_dict, method_name, "total_hits", int(value))
                elif "errors.server.count" in metric_id:
                    _assign_field(data_dict, method_name, "failure_count", int(value))
                elif ":avg" in metric_id or "avg_responsetime" in metric_id:
                    _assign_field(data_dict, method_name, "avg", round(value / 1_000_000, 2))
                elif "percentile(90.0)" in metric_id:
                    _assign_field(data_dict, method_name, "p90", round(value / 1_000_000, 2))
                elif "percentile(95.0)" in metric_id:
                    _assign_field(data_dict, method_name, "p95", round(value / 1_000_000, 2))
                elif "percentile(99.0)" in metric_id:
                    _assign_field(data_dict, method_name, "p99", round(value / 1_000_000, 2))
                elif "errors.server.rate" in metric_id:
                    _assign_field(data_dict, method_name, "failure_rate", round(value, 2))

    if DEBUG:
        print(f"\n📦 Final data_dict for {controller_name}:")
        for method, vals in data_dict.items():
            print(f"🔍 {method} → {vals}")

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
    update_workbook(sheet_name, df)

def update_workbook(sheet_name, df_new):
    all_sheets = {}

    if os.path.exists(FILE_NAME):
        with pd.ExcelFile(FILE_NAME, engine="openpyxl") as xls:
            for name in xls.sheet_names:
                try:
                    all_sheets[name] = pd.read_excel(xls, sheet_name=name)
                except Exception:
                    all_sheets[name] = pd.DataFrame()
        print(f"📄 Existing sheets loaded: {list(all_sheets.keys())}")
    else:
        print("🆕 No existing workbook found. Creating new one.")

    existing_df = all_sheets.get(sheet_name, pd.DataFrame())
    combined_df = pd.concat([df_new, existing_df], ignore_index=True)
    all_sheets[sheet_name] = combined_df

    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    print(f"✅ Sheet '{sheet_name}' updated with new data.")

def send_email_report():
    EMAIL_USER = os.environ["EMAIL_USER"]
    EMAIL_PASS = os.environ["EMAIL_PASS"]
    SMTP_SERVER = os.environ["EMAIL_HOST"]
    SMTP_PORT = int(os.environ["EMAIL_PORT"])
    TO_EMAIL = os.environ["EMAIL_TO"]

    msg = EmailMessage()
    msg["Subject"] = "📊 Dynatrace Metrics Report"
    msg["From"] = EMAIL_USER
    msg["To"] = TO_EMAIL
    msg.set_content("Hi,\n\nAttached is the updated Product Category wise Internal APIs Performance Report.\n\nRegards,\nSubhranshu")

    with open(FILE_NAME, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=FILE_NAME
        )

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)
        print("📤 Email sent successfully.")
    except Exception as e:
        print("❌ Email sending failed:", e)

if __name__ == "__main__":
    API_TOKEN = os.environ["API_TOKEN"]
    controller_urls = {
        "LoginController": os.environ.get("LOGINCONTROLLER_URL"),
        "MotorInsurance": os.environ.get("MOTORINSURANCE_URL"),
        "CreditTrack": os.environ.get("CREDITTRACK_URL"),
        "DigiGold & Silver": os.environ.get("DIGIGOLD_URL"),
        "HealthInsurance": os.environ.get("HEALTHINSURANCE_URL"),
        "LR Rewards": os.environ.get("LRREWARDS_URL"),
        "Personal Loan": os.environ.get("PERSONALLOAN_URL"),
        "MutualFund": os.environ.get("MUTUALFUND_URL"),
        "Fixed Deposit": os.environ.get("FIXEDDEPOSIT_URL"),
        "BusinessLoanController": os.environ.get("BLOAN_URL"),
        "TermInsuranceBuyController(LI)": os.environ.get("TERM_URL"),
        "Stocks": os.environ.get("STOCKS_URL"),
        "GoldLoanController": os.environ.get("GOLD_URL"),
        "Home Loan": os.environ.get("HOMELOAN_URL"),
        "PortfolioTrack": os.environ.get("PORTFOLIO_URL"),
        "SpendTrack": os.environ.get("SPENDTRACK_URL"),
    }

    for name, url in controller_urls.items():
        if not url:
            print(f"⚠️ Skipping {name}: URL not set in environment")
            continue
        fetch_and_store_metrics(name, url, API_TOKEN)

    apply_formatting(FILE_NAME)
    send_email_report()
