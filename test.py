import os
import requests
import pandas as pd
import smtplib
from email.message import EmailMessage

# Read credentials from environment variables
API_TOKEN = os.environ["API_TOKEN"]
EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASS = os.environ["EMAIL_PASS"]
SMTP_SERVER = os.environ["EMAIL_HOST"]
SMTP_PORT = int(os.environ["EMAIL_PORT"])
TO_EMAIL = os.environ["EMAIL_TO"]  # ✅ fixed

# URL is hardcoded
url = 'https://abcdapm.adityabirlacapital.com/e/aae1eca1-1d7c-4a0b-a10c-257b60cfe008/api/v2/metrics/query?metricSelector=(builtin:service.keyRequest.count.total:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names,(builtin:service.keyRequest.errors.server.count:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):avg:sort(value(avg,descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(90.0):sort(value(percentile(90.0),descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(95.0):sort(value(percentile(95.0),descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(99.0):sort(value(percentile(99.0),descending))):names,(builtin:service.keyRequest.errors.server.rate:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names&from=-1d/d&to=now/d&resolution=Inf&mzSelector=mzId(-2729326265202688410)'

headers = {
    "Authorization": f"Api-Token {API_TOKEN}",
    "accept": "application/json"
}
response = requests.get(url, headers=headers)
print("Status Code:", response.status_code)
if response.status_code != 200:
    print("Error:", response.text)
    exit()

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
            data_dict[method_name]["avg"] = round(value / 1000000, 2)
        elif "percentile(90.0)" in metric_id:
            data_dict[method_name]["p90"] = round(value / 1000000, 2)
        elif "percentile(95.0)" in metric_id:
            data_dict[method_name]["p95"] = round(value / 1000000, 2)
        elif "percentile(99.0)" in metric_id:
            data_dict[method_name]["p99"] = round(value / 1000000, 2)
        elif "errors.server.rate" in metric_id:
            data_dict[method_name]["failure_rate"] = round(value, 2)

records = []
for method, values in data_dict.items():
    records.append({
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
excel_file = "dynatrace_metrics.xlsx"
df.to_excel(excel_file, index=False)

# Send Email with attachment
msg = EmailMessage()
msg["Subject"] = "Dynatrace Daily Metrics Report"
msg["From"] = EMAIL_USER
msg["To"] = TO_EMAIL
msg.set_content("Hi,\n\nPlease find attached the latest Dynatrace metrics report.\n\nRegards,\nAutomated Bot")

# Attach the Excel file
with open("dynatrace_metrics.xlsx", "rb") as f:
    file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="dynatrace_metrics.xlsx")

# Send the email
try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)
    print("✅ Email sent successfully.")
except Exception as e:
    print("❌ Failed to send email.")
    print("Error:", e)