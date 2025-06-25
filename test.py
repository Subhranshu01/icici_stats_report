import requests
import pandas as pd
import configparser

# Load token from config.ini
config = configparser.ConfigParser()
config.read("config.ini")
API_TOKEN = config["auth"]["api_token"]

# Paste your full Dynatrace metrics URL here
url = 'https://abcdapm.adityabirlacapital.com/e/aae1eca1-1d7c-4a0b-a10c-257b60cfe008/api/v2/metrics/query?metricSelector=(builtin:service.keyRequest.count.total:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names,(builtin:service.keyRequest.errors.server.count:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):avg:sort(value(avg,descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(90.0):sort(value(percentile(90.0),descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(95.0):sort(value(percentile(95.0),descending))):names,(builtin:service.keyRequest.response.server:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):percentile(99.0):sort(value(percentile(99.0),descending))):names,(builtin:service.keyRequest.errors.server.rate:filter(or(in("dt.entity.service_method",entitySelector("type(service_method),fromRelationship.isServiceMethodOfService(type(SERVICE),entityName.equals(~"LoginController~"))")))):splitBy("dt.entity.service_method"):sort(value(auto,descending))):names&from=-1d/d&to=now/d&resolution=Inf&mzSelector=mzId(-2729326265202688410)'

# Send request
headers = {
    "Authorization": f"Api-Token {API_TOKEN}",
    "accept": "application/json"
}
response = requests.get(url, headers=headers)

# Handle response
print("Status Code:", response.status_code)
if response.status_code != 200:
    print("Error:", response.text)
    exit()

json_data = response.json()
result = json_data.get("result", [])

# Parse and aggregate data

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
            data_dict[method_name]["avg"] = round(value /1000000, 2)
        elif "percentile(90.0)" in metric_id:
            data_dict[method_name]["p90"] = round(value / 1000000, 2)
        elif "percentile(95.0)" in metric_id:
            data_dict[method_name]["p95"] = round(value / 1000000, 2)
        elif "percentile(99.0)" in metric_id:
            data_dict[method_name]["p99"] = round(value / 1000000, 2)
        elif "errors.server.rate" in metric_id:
            data_dict[method_name]["failure_rate"] = round(value, 2)

# Convert to DataFrame
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
df.to_excel("dynatrace_metrics.xlsx", index=False)
print("✅ Data written to dynatrace_metrics.xlsx")
