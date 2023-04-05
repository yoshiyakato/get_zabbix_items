import json
import requests
import datetime
from openpyxl import Workbook

# Zabbix APIのURL、ユーザー名、パスワードを設定
ZABBIX_URL = "http://09092310973:317567@localhost/zabbix/api_jsonrpc.php"
ZABBIX_USER = "Admin"
ZABBIX_PASSWORD = "zabbix"

# セッションを作成
session = requests.Session()
session.headers.update({"Content-Type": "application/json-rpc"})

def call_api(method, params, auth=None):
    payload = {
        "jsonrpc": "2.0",
        "method": method,
        "params": params,
        "id": 1,
        "auth": auth
    }
    response = session.post(ZABBIX_URL, data=json.dumps(payload))

    if response.status_code != 200:
        print(f"Error: Received status code {response.status_code} from Zabbix API")
        print(response.text)
        return None

    json_data = response.json()
    if "error" in json_data:
        print("Error in Zabbix API response:")
        print(json_data["error"])
        return None

    return json_data["result"]

def login(username, password):
    result = call_api("user.login", {"user": username, "password": password}, auth=None)
    session.headers.update({"auth": result})
    return result

def get_all_hosts(auth):
    return call_api("host.get", {"output": ["hostid", "name", "host"]}, auth)

def get_items_by_host(host_id, auth):
    return call_api("item.get", {"output": ["itemid", "name", "key_", "value_type"], "hostids": host_id}, auth)

# Zabbixにログイン
auth = login(ZABBIX_USER, ZABBIX_PASSWORD)
if not auth:
    print("ログインに失敗しました。ユーザー名とパスワードを確認してください。")
    exit(1)

# すべてのホストを取得
hosts = get_all_hosts(auth)
if hosts is None:
    print("ホストの取得に失敗しました。")
    exit(1)

# Excelワークブックとワークシートを作成
workbook = Workbook()
sheet = workbook.active
sheet.title = "Zabbix_Items"

# ヘッダー行を作成
header = ["Host ID", "Host Name", "Host", "Item ID", "Item Name", "Item Key", "Value Type"]
sheet.append(header)

# すべてのホストのアイテムを取得してExcelに追加
for host in hosts:
    host_id = host["hostid"]
    host_name = host["name"]
    host_host = host["host"]

    items = get_items_by_host(host_id,auth)
    for item in items:
        item_id = item["itemid"]
        item_name = item["name"]
        item_key = item["key_"]
        value_type = item["value_type"]

        row = [host_id, host_name, host_host, item_id, item_name, item_key, value_type]
        sheet.append(row)

# 現在の日時を取得
now = datetime.datetime.now()
timestamp = now.strftime("%Y%m%d_%H%M%S")
# 現在の日時を取得
now = datetime.datetime.now()
timestamp = now.strftime("%Y%m%d_%H%M%S")

# Excelファイル名にタイムスタンプを追加
output_file = f"zabbix_items_{timestamp}.xlsx"

# Excelファイルを保存
workbook.save(output_file)
print(f"Excelファイルにアイテムが出力されました: {output_file}")
