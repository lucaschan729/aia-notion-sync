import io
import os
import subprocess
import sys

import requests

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])
from openpyxl import load_workbook

TOKEN = os.environ["NOTION_TOKEN"].strip()
TARGET = os.environ["NOTION_DATABASE_ID"].strip()
API = "https://api.notion.com/v1"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2025-09-03",
    "Content-Type": "application/json",
}


def cell_text(value):
    if value is None:
        return ""
    return str(value).replace("\n", " ")[:25]


def show_row(ws, n):
    row = next(ws.iter_rows(min_row=n, max_row=n))
    cells = [cell_text(c.value) for c in row]
    while cells and not cells[-1]:
        cells.pop()
    print(f"  {n}: " + " | ".join(cells))


def show_workbook(content):
    wb = load_workbook(io.BytesIO(content))
    print("工作表：", wb.sheetnames)
    for ws in wb.worksheets:
        print()
        print(f"### {ws.title}（{ws.max_row} 行 × {ws.max_column} 欄）")
        merged = [str(m) for m in ws.merged_cells.ranges]
        if merged:
            print("合併儲存格：", ", ".join(merged[:8]), "…" if len(merged) > 8 else "")
        head = min(8, ws.max_row)
        print("開頭：")
        for n in range(1, head + 1):
            show_row(ws, n)
        tail_start = max(head + 1, ws.max_row - 3)
        if tail_start <= ws.max_row:
            print("結尾：")
            for n in range(tail_start, ws.max_row + 1):
                show_row(ws, n)


r = requests.post(f"{API}/data_sources/{TARGET}/query", headers=HEADERS, json={"page_size": 20}, timeout=30)
r.raise_for_status()
found = False
for page in r.json().get("results", []):
    r2 = requests.get(f"{API}/blocks/{page['id']}/children?page_size=100", headers=HEADERS, timeout=30)
    r2.raise_for_status()
    for block in r2.json().get("results", []):
        if block["type"] != "file":
            continue
        info = block["file"]
        name = info.get("name", "")
        if not name.lower().endswith(".xlsx"):
            continue
        url = (info.get("file") or info.get("external") or {}).get("url")
        if not url:
            print("找不到下載連結：", name)
            continue
        found = True
        print("=== 檔案：", name, "===")
        data = requests.get(url, timeout=60)
        data.raise_for_status()
        show_workbook(data.content)

if not found:
    print("沒有找到 .xlsx 附件")
