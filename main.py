import os
import requests

TOKEN = os.environ["NOTION_TOKEN"].strip()
TARGET = os.environ.get("NOTION_DATABASE_ID", "").strip()
API = "https://api.notion.com/v1"


def headers(version):
    return {
        "Authorization": f"Bearer {TOKEN}",
        "Notion-Version": version,
        "Content-Type": "application/json",
    }


def title_of(obj):
    if isinstance(obj.get("title"), list):
        return "".join(t.get("plain_text", "") for t in obj["title"]) or "(無標題)"
    for prop in obj.get("properties", {}).values():
        if isinstance(prop, dict) and prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", [])) or "(無標題)"
    return "(無標題)"


print("=== 1. integration 能看到的內容 ===")
items, cursor = [], None
while True:
    body = {"page_size": 100}
    if cursor:
        body["start_cursor"] = cursor
    r = requests.post(f"{API}/search", headers=headers("2022-06-28"), json=body, timeout=30)
    if r.status_code != 200:
        print("搜尋失敗：", r.status_code, r.text[:300])
        break
    data = r.json()
    items += data.get("results", [])
    if not data.get("has_more"):
        break
    cursor = data.get("next_cursor")

if not items:
    print("甚麼都看不到 → 頁面尚未分享給 integration")
for it in items:
    kind = "資料庫" if it["object"] == "database" else "頁面"
    same = "  ← 與目前設定的 ID 相同" if it["id"] == TARGET else ""
    print(f"[{kind}] {title_of(it)} | {it['id'].replace('-', ' ')}{same}")

print()
print("=== 2. 用目前設定的 ID 逐種類型測試 ===")
tests = [
    ("資料庫", "databases", "2022-06-28"),
    ("頁面", "pages", "2022-06-28"),
    ("區塊", "blocks", "2022-06-28"),
    ("資料來源", "data_sources", "2025-09-03"),
]
for label, path, version in tests:
    r = requests.get(f"{API}/{path}/{TARGET}", headers=headers(version), timeout=30)
    print(f"{label}：{r.status_code}", "成功" if r.status_code == 200 else "找不到")
