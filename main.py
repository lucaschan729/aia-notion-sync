import datetime
import io
import json
import os
import re
import statistics
import subprocess
import sys
import zipfile
from collections import Counter

import requests

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

TOKEN = os.environ["NOTION_TOKEN"].strip()
TARGET = os.environ["NOTION_DATABASE_ID"].strip()
API = "https://api.notion.com/v1"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2025-09-03",
    "Content-Type": "application/json",
}
ROOT = "https://www.aia.com.hk/content/dam/hk-wise/json/further-product-information/"
PAGE = "https://www.aia.com.hk/zh-hk/products/further-product-information/participating-products/fulfillment-ratio"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
DIV_NAMES = {"AD": "周年红利", "TD": "终期红利", "RB": "复归红利", "TB": "终期分红"}
PERCENT = False


# ---------- 下載友邦資料 ----------
def fetch_json(name):
    this_year = datetime.date.today().year
    for y in (this_year + 1, this_year, this_year - 1, this_year - 2):
        try:
            r = requests.get(f"{ROOT}{y}/{name}", headers={"User-Agent": UA, "Referer": PAGE}, timeout=60)
            if r.status_code == 200:
                obj = json.loads(r.content.decode("utf-8-sig"))
                print(f"已下載 {name}（資料夾 {y}，申報年度 {obj.get('report_year')}）")
                return obj
        except Exception as e:
            print(f"{y}/{name} 失敗：{str(e)[:100]}")
    raise SystemExit(f"找不到 {name}")


def zh(v):
    if isinstance(v, dict):
        v = v.get("zh-cn") or v.get("zh-hk") or v.get("en") or ""
    if isinstance(v, list):
        v = "".join(str(x) for x in v)
    return str(v).strip()


def to_number(text):
    try:
        return float(text.replace(",", "").replace("%", "").strip())
    except ValueError:
        return None


def convert(raw):
    n = to_number(raw)
    if n is None:
        return "/"
    return round(n / 100 if PERCENT else n, 2)


def all_rows(obj, keys):
    for p in obj.get("pData", []):
        for k in keys:
            for row in p.get(k) or []:
                yield row


def year_columns(obj, keys):
    cols = []
    for row in all_rows(obj, keys):
        labels = [zh(d.get("year")) for d in row.get("data", [])]
        if len(labels) > len(cols):
            cols = labels
    return cols


def values_for(row, cols):
    data = row.get("data", [])
    if len(data) == len(cols):
        return [convert(zh(d.get("ratio"))) for d in data]
    m = {zh(d.get("year")): zh(d.get("ratio")) for d in data}
    return [convert(m.get(c, "")) for c in cols]


def build_fulfillment(obj):
    cols = year_columns(obj, DIV_NAMES)
    out = []
    for p in obj.get("pData", []):
        name, kind = zh(p.get("productNm")), zh(p.get("type"))
        keys = [k for k in DIV_NAMES if isinstance(p.get(k), list)]
        n = max((len(p[k]) for k in keys), default=0)
        for i in range(n):
            for k in keys:
                if i < len(p[k]):
                    row = p[k][i]
                    out.append([name, kind, zh(row.get("currency")), DIV_NAMES[k]] + values_for(row, cols))
    return cols, out


def build_cash(obj):
    cols = year_columns(obj, ["row"])
    out = []
    for p in obj.get("pData", []):
        name, kind = zh(p.get("productNm")), zh(p.get("type"))
        for row in p.get("row") or []:
            out.append([name, kind, zh(row.get("currency"))] + values_for(row, cols))
    return cols, out


def make_headers(cols, fixed):
    n = len(cols)
    top = fixed + [f"保单年度{i + 1}" for i in range(n)]
    second = [""] * len(fixed) + list(cols)
    if n >= 2 and not re.fullmatch(r"\d{4}", cols[-1]):
        top[-1] = f"保单年度{n - 1}+"
        second[-1] = f"（{cols[-2]}之前）"
    return top, second


# ---------- 生成 Excel ----------
def merge_runs(ws, rows, col, key):
    start = 0
    for i in range(1, len(rows) + 1):
        if i == len(rows) or key(rows[i]) != key(rows[start]):
            if i - start > 1:
                ws.merge_cells(start_row=3 + start, start_column=col, end_row=2 + i, end_column=col)
            start = i


def add_sheet(wb, title, top, second, rows, n_fixed, merges):
    ws = wb.create_sheet(title)
    ws.append(top)
    ws.append(second)
    for row in rows:
        ws.append(row)
    for c in range(1, n_fixed + 1):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
    for col, key in merges:
        merge_runs(ws, rows, col, key)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
    for cell in ws[1] + ws[2]:
        cell.font = Font(bold=True)
    widths = [42, 14, 10, 12][:n_fixed]
    for i in range(1, len(top) + 1):
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1] if i <= n_fixed else 12
    ws.freeze_panes = ws.cell(row=3, column=n_fixed + 1)


# ---------- 讀取 Notion 現有檔案 ----------
def fix_xlsx(content):
    src = zipfile.ZipFile(io.BytesIO(content))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = re.sub(r"<(\w+:)?family\b[^>]*/>", "", data.decode("utf-8")).encode("utf-8")
            dst.writestr(item.filename, data)
    return out.getvalue()


def find_old_file(year):
    r = requests.post(f"{API}/data_sources/{TARGET}/query", headers=HEADERS, json={"page_size": 50}, timeout=30)
    r.raise_for_status()
    for page in r.json().get("results", []):
        r2 = requests.get(f"{API}/blocks/{page['id']}/children?page_size=100", headers=HEADERS, timeout=30)
        r2.raise_for_status()
        for block in r2.json().get("results", []):
            if block["type"] != "file":
                continue
            info = block["file"]
            name = info.get("name", "")
            url = (info.get("file") or info.get("external") or {}).get("url")
            if name.lower().endswith(".xlsx") and str(year) in name and url:
                data = requests.get(url, timeout=60)
                data.raise_for_status()
                print("核對對象：", name)
                try:
                    return load_workbook(io.BytesIO(data.content), data_only=True)
                except ValueError:
                    return load_workbook(io.BytesIO(fix_xlsx(data.content)), data_only=True)
    return None


def norm_text(v):
    return "" if v is None else re.sub(r"\s+", "", str(v))


def norm_val(v):
    if v is None or str(v).strip() == "":
        return ""
    s = str(v).strip().replace("／", "/")
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return s


def read_old(ws, n_fixed, key_idx, ncols):
    out = {}
    last = ["", "", ""]
    for row in ws.iter_rows(min_row=3, values_only=True):
        cells = list(row) + [None] * (n_fixed + ncols)
        fixed = [norm_text(c) for c in cells[:n_fixed]]
        if not any(fixed):
            continue
        for i in range(3):
            if fixed[i]:
                last[i] = fixed[i]
            else:
                fixed[i] = last[i]
        key = tuple(fixed[i] for i in key_idx)
        out.setdefault(key, [norm_val(c) for c in cells[n_fixed:n_fixed + ncols]])
    return out


def index_new(rows, n_fixed, key_idx):
    out = {}
    for row in rows:
        key = tuple(norm_text(row[i]) for i in key_idx)
        out.setdefault(key, [norm_val(c) for c in row[n_fixed:]])
    return out


def compare(label, old, new):
    same = [k for k in new if k in old and old[k] == new[k]]
    diff = [k for k in new if k in old and old[k] != new[k]]
    only_new = [k for k in new if k not in old]
    only_old = [k for k in old if k not in new]
    print()
    print(f"【{label}】新檔 {len(new)} 行｜舊檔 {len(old)} 行")
    print(f"  完全相同 {len(same)}｜數值不同 {len(diff)}｜只在新檔 {len(only_new)}｜只在舊檔 {len(only_old)}")
    for k in diff[:6]:
        print("  不同：", "/".join(k)[:45])
        print("    舊：", " ".join(old[k]))
        print("    新：", " ".join(new[k]))
    for k in only_new[:6]:
        print("  只在新檔：", "/".join(k)[:60])
    for k in only_old[:6]:
        print("  只在舊檔：", "/".join(k)[:60])


# ---------- 主流程 ----------
fr = fetch_json("fulfillment-ratio.json")
cv = fetch_json("total-cash-value-ratio.json")
year = fr.get("report_year")

raws = [zh(d.get("ratio")) for obj, keys in ((fr, DIV_NAMES), (cv, ["row"]))
        for row in all_rows(obj, keys) for d in row.get("data", [])]
nums = [n for n in (to_number(x) for x in raws) if n is not None]
PERCENT = any("%" in x for x in raws) or (bool(nums) and statistics.median(nums) > 10)
texts = Counter(x for x in raws if to_number(x) is None)

print()
print("=== 原始資料 ===")
print("數值寫法：", "百分比（會除以 100）" if PERCENT else "倍數（直接使用）")
print("數字例子：", [x for x in raws if to_number(x) is not None][:8])
print("非數字內容：", dict(texts.most_common(8)))
print("產品內的鍵：", dict(Counter(k for p in fr.get("pData", []) for k in p)))

f_cols, f_rows = build_fulfillment(fr)
c_cols, c_rows = build_cash(cv)
print("年份欄（分紅）：", f_cols)
print("年份欄（總現金）：", c_cols)

f_top, f_second = make_headers(f_cols, ["产品名称", "产品种类", "货币", "红利类型"])
c_top, c_second = make_headers(c_cols, ["产品名称", "产品种类", "货币"])
wb = Workbook()
wb.remove(wb.active)
add_sheet(wb, f"{year}分紅實現率", f_top, f_second, f_rows, 4,
          [(1, lambda r: r[0]), (2, lambda r: (r[0], r[1])), (3, lambda r: (r[0], r[2]))])
add_sheet(wb, f"{year}總現金價值比率", c_top, c_second, c_rows, 3,
          [(1, lambda r: r[0]), (2, lambda r: (r[0], r[1]))])
buf = io.BytesIO()
wb.save(buf)

print()
print("=== 生成結果 ===")
print(f"分紅實現率 {len(f_rows)} 行｜總現金價值比率 {len(c_rows)} 行｜檔案 {len(buf.getvalue()) // 1024} KB")
print("標題：", " | ".join(f_top[-3:]), "→", " | ".join(str(x) for x in f_second[-3:]))
for row in f_rows[:2] + c_rows[:1]:
    print("  ", " | ".join(str(x) for x in row)[:150])

print()
print("=== 與 Notion 現有檔案核對 ===")
old_wb = find_old_file(year)
if old_wb is None:
    print(f"Notion 沒有 {year} 年的檔案，略過核對")
else:
    sheets = old_wb.worksheets
    compare("分紅實現率", read_old(sheets[0], 4, (0, 2, 3), len(f_cols)), index_new(f_rows, 4, (0, 2, 3)))
    if len(sheets) > 1:
        compare("總現金價值比率", read_old(sheets[1], 3, (0, 2), len(c_cols)), index_new(c_rows, 3, (0, 2)))
