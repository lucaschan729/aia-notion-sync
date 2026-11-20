import os
import requests

TOKEN = os.environ["NOTION_TOKEN"].strip()
TARGET = os.environ["NOTION_DATABASE_ID"].strip()
API = "https://api.notion.com/v1"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2025-09-03",
    "Content-Type": "application/json",
}


def text_of(rich):
    return "".join(t.get("plain_text", "") for t in rich or [])


def value_of(prop):
    kind = prop.get("type")
    v = prop.get(kind)
    if v is None:
        return ""
    if kind in ("title", "rich_text"):
        return text_of(v)
    if kind in ("select", "status"):
        return v.get("name", "")
    if kind == "multi_select":
        return ", ".join(x.get("name", "") for x in v)
    if kind == "date":
        return v.get("start", "")
    if kind == "files":
        return ", ".join(x.get("name", "") for x in v)
    if kind in ("number", "checkbox", "url"):
        return str(v)
    return f"({kind})"


def show_blocks(block_id, depth=0, counter=None):
    counter = counter if counter is not None else [0]
    r = requests.get(f"{API}/blocks/{block_id}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code != 200:
        print("  " * depth + f"讀取失敗：{r.status_code} {r.text[:200]}")
        return
    table_rows = 0
    for b in r.json().get("results", []):
        if counter[0] >= 80:
            print("  " * depth + "（內容太多，只顯示前 80 項）")
            return
        kind = b["type"]
        body = b.get(kind, {}) or {}
        if kind == "table_row":
            table_rows += 1
            if table_rows > 3:
                continue
            info = " | ".join(text_of(c) for c in body.get("cells", []))
        elif kind in ("child_page", "child_database"):
            info = body.get("title", "")
        elif kind == "table":
            info = f"{body.get('table_width')} 欄"
        elif kind in ("file", "pdf", "image"):
            info = body.get("name", "") or text_of(body.get("caption"))
        else:
            info = text_of(body.get("rich_text"))
        counter[0] += 1
        print("  " * depth + f"- [{kind}] {info[:80]}")
        if b.get("has_children") and depth < 3 and kind != "child_page":
            show_blocks(b["id"], depth + 1, counter)
    if table_rows > 3:
        print("  " * depth + f"（表格共 {table_rows} 行，只顯示前 3 行）")


print("=== 1. 資料來源 ===")
r = requests.get(f"{API}/data_sources/{TARGET}", headers=HEADERS, timeout=30)
r.raise_for_status()
ds = r.json()
print("名稱：", text_of(ds.get("title")) or "(無標題)")
print("所屬：", ds.get("parent"))
print("欄位：")
for name, prop in ds.get("properties", {}).items():
    print(f"  - {name}（{prop.get('type')}）")

print()
print("=== 2. 現有記錄 ===")
r = requests.post(f"{API}/data_sources/{TARGET}/query", headers=HEADERS, json={"page_size": 20}, timeout=30)
r.raise_for_status()
rows = r.json().get("results", [])
print("記錄數：", len(rows))
for row in rows:
    print("記錄：", row["id"])
    for name, prop in row.get("properties", {}).items():
        print(f"  {name} = {value_of(prop)}")

print()
print("=== 3. 每筆記錄頁面內的內容 ===")
for row in rows[:3]:
    title = ""
    for prop in row.get("properties", {}).values():
        if prop.get("type") == "title":
            title = text_of(prop.get("title"))
    print(f"--- {title} ---")
    show_blocks(row["id"])
