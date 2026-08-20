#!/usr/bin/env python3
import json
import urllib.request
import urllib.error

INDEXNOW_KEY = "7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f"
HOST = "excelarsiv.com"
KEY_LOCATION = f"https://{HOST}/7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f.txt"
URL_LIST = ["https://excelarsiv.com/"]
ENDPOINTS = ["https://api.indexnow.org/indexnow", "https://www.bing.com/indexnow"]

def ping():
    payload = {"host": HOST, "key": INDEXNOW_KEY, "keyLocation": KEY_LOCATION, "urlList": URL_LIST}
    data = json.dumps(payload).encode('utf-8')
    headers = {'Content-Type': 'application/json; charset=utf-8', 'User-Agent': 'IndexNow-Notifier/2.0'}
    for ep in ENDPOINTS:
        try:
            req = urllib.request.Request(ep, data=data, headers=headers, method='POST')
            with urllib.request.urlopen(req, timeout=10) as resp:
                print(f"✅ {ep} -> HTTP {resp.getcode()}")
        except Exception as e:
            print(f"⚠️ {ep} -> {e}")

if __name__ == "__main__":
    ping()
