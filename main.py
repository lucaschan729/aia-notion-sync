import os
import sys
import requests

TOKEN = os.environ["NOTION_TOKEN"].strip()
TARGET = os.environ["NOTION_DATABASE_ID"].strip()
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2022-06-28",
}


def get(kind):
    url = f"https://api.notion.com/v1/{kind}/{TARGET}"
    return requests.get(url, headers=HEADERS, timeout=30)


db = get("databases")
if db.status_code == 200:
    data = db.json()
    title = "".join(t["plain_text"] for t in data.get("title", []))
    print("結果：這是資料庫")
    print("名稱：", title)
    print("欄位：")
    for name, prop in data["properties"].items():
        print(" -", name, "/", prop["type"])
    sys.exit(0)

page = get("pages")
if page.status_code == 200:
    print("結果：這是普通頁面，不是資料庫")
    sys.exit(0)

print("連線失敗")
print("資料庫查詢：", db.status_code, db.json().get("message"))
print("頁面查詢：", page.status_code, page.json().get("message"))
sys.exit(1)
