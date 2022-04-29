#coding:utf-8
import requests
import json
import csv
import time

def get_token():
    data = {
        "jsonrpc": "2.0",
        "method": "user.login",
        "params": {
            "user": username,
            "password": password
        },
        "id": 0
    }
    r = requests.get(zaurl, headers=header, data=json.dumps(data))
    auth = json.loads(r.text)
    return auth["result"]

if __name__ == "__main__":
    zaurl = "http://192.168.1.25/zabbix/api_jsonrpc.php"  # 访问zabbix页面的url
    header = {"Content-Type": "application/json"}
    username = "Admin"  # zabbix的账户与密码
    password = "zabbix"
    hostIp = "192.168.1.42"
    token = get_token()
    print(token)
