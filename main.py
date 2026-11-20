import datetime
import difflib
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
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
ROOT = "https://www.aia.com.hk/content/dam/hk-wise/json/further-product-information/"
PAGE = "https://www.aia.com.hk/zh-hk/products/further-product-information/participating-products/fulfillment-ratio"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
DIV_NAMES = {"AD": "周年红利", "TD": "终期红利", "RB": "复归红利", "TB": "终期分红"}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
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


def clean(text):
    text = re.sub(r"<sup>.*?</sup>", "", text, flags=re.S | re.I)
    return re.sub(r"<[^>]+>", "", text).strip()


def to_number(text):
    try:
        return float(clean(text).replace(",", "").replace("%", "").strip())
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
        name, kind = clean(zh(p.get("productNm"))), clean(zh(p.get("type")))
        keys = [k for k in DIV_NAMES if isinstance(p.get(k), list)]
        n = max((len(p[k]) for k in keys), default=0)
        for i in range(n):
            for k in keys:
                if i < len(p[k]):
                    row = p[k][i]
                    out.append([name, kind, clean(zh(row.get("currency"))), DIV_NAMES[k]] + values_for(row, cols))
    return cols, out


def build_cash(obj):
    cols = year_columns(obj, ["row"])
    out = []
    for p in obj.get("pData", []):
        name, kind = clean(zh(p.get("productNm"))), clean(zh(p.get("type")))
        for row in p.get("row") or []:
            out.append([name, kind, clean(zh(row.get("currency")))] + values_for(row, cols))
    return cols, out


def make_headers(cols, fixed):
    n = len(cols)
    top = fixed + [f"保单年度{i + 1}" for i in range(n)]
    second = [""] * len(fixed) + list(cols)
    if n >= 2 and not re.fullmatch(r"\d{4}", cols[-1]):
        top[-1] = f"保单年度{n - 1}+"
        second[-1] = f"（{cols[-2]}之前）"
    return top, second


def average_row(rows, n_fixed, label_cells):
    out = list(label_cells)
    width = max(len(r) for r in rows) if rows else n_fixed
    for c in range(n_fixed, width):
        nums = [r[c] for r in rows if isinstance(r[c], (int, float))]
        out.append(round(sum(nums) / len(nums), 2) if nums else "/")
    return out


# ---------- 生成 Excel ----------
def merge_runs(ws, rows, col, key):
    start = 0
    for i in range(1, len(rows) + 1):
        if i == len(rows) or key(rows[i]) != key(rows[start]):
            if i - start > 1:
                ws.merge_cells(start_row=3 + start, start_column=col, end_row=2 + i, end_column=col)
            start = i


def add_sheet(wb, title, top, second, rows, footer, n_fixed, merges):
    ws = wb.create_sheet(title)
    ws.append(top)
    ws.append(second)
    for row in rows:
        ws.append(row)
    for row in footer:
        ws.append(row)
    for c in range(1, n_fixed + 1):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
    for col, key in merges:
        merge_runs(ws, rows, col, key)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
    bold = Font(bold=True)
    for cell in ws[1] + ws[2]:
        cell.font = bold
    for i in range(len(footer)):
        for cell in ws[3 + len(rows) + i]:
            cell.font = bold
    widths = [42, 14, 10, 12][:n_fixed]
    for i in range(1, len(top) + 1):
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1] if i <= n_fixed else 12
    ws.freeze_panes = ws.cell(row=3, column=n_fixed + 1)


# ---------- Notion ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def query_pages():
    r = requests.post(f"{API}/data_sources/{TARGET}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
    return check(r, "讀取年份資料庫").get("results", [])


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=HEADERS, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            url = (info.get("file") or info.get("external") or {}).get("url")
            out.append((block["id"], info.get("name", ""), url))
    return out


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


def load_manual_file(page_id):
    for _, name, url in file_blocks(page_id):
        if name.lower().endswith(".xlsx") and AUTO not in name and url:
            data = requests.get(url, timeout=60)
            data.raise_for_status()
            print("核對對象：", name)
            try:
                return load_workbook(io.BytesIO(data.content), data_only=True)
            except ValueError:
                return load_workbook(io.BytesIO(fix_xlsx(data.content)), data_only=True)
    return None


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name, _ in file_blocks(page_id) if AUTO in name]
    r = requests.post(f"{API}/file_uploads", headers=HEADERS,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                      headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": VERSION},
                      files={"file": (filename, content, XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id}}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=HEADERS, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=HEADERS, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 核對 ----------
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
    diff = [(k, k) for k in new if k in old and old[k] != new[k]]
    only_new = [k for k in new if k not in old]
    only_old = [k for k in old if k not in new]
    used, renamed_same, left_new = set(), [], []
    for k in only_new:
        best, score = None, 0.0
        for o in only_old:
            if o in used or o[1:] != k[1:]:
                continue
            s = difflib.SequenceMatcher(None, k[0], o[0]).ratio()
            if s > score:
                best, score = o, s
        if best is not None and score >= 0.6:
            used.add(best)
            if new[k] == old[best]:
                renamed_same.append((k, best))
            else:
                diff.append((k, best))
        else:
            left_new.append(k)
    left_old = [o for o in only_old if o not in used]
    print()
    print(f"【{label}】新檔 {len(new)} 行｜舊檔 {len(old)} 行")
    print(f"  完全相同 {len(same)}｜名稱寫法不同但數值相同 {len(renamed_same)}｜數值不同 {len(diff)}")
    print(f"  只在新檔 {len(left_new)}｜只在舊檔 {len(left_old)}")
    for k, o in renamed_same[:3]:
        print("  名稱不同：", k[0][:28], "←", o[0][:28])
    for k, o in diff[:8]:
        print("  不同：", "/".join(k)[:50], "" if k == o else f"（舊名 {o[0][:20]}）")
        print("    舊：", " ".join(old[o]))
        print("    新：", " ".join(new[k]))
    for k in left_new[:10]:
        print("  只在新檔：", "/".join(k)[:60])
    for k in left_old[:10]:
        print("  只在舊檔：", "/".join(k)[:60])


# ---------- 主流程 ----------
fr = fetch_json("fulfillment-ratio.json")
cv = fetch_json("total-cash-value-ratio.json")
year = fr.get("report_year")
if cv.get("report_year") != year:
    raise SystemExit(f"兩個資料檔的申報年度不同：{year} / {cv.get('report_year')}，暫不更新")

raws = [zh(d.get("ratio")) for obj, keys in ((fr, DIV_NAMES), (cv, ["row"]))
        for row in all_rows(obj, keys) for d in row.get("data", [])]
nums = [n for n in (to_number(x) for x in raws) if n is not None]
PERCENT = any("%" in x for x in raws) or (bool(nums) and statistics.median(nums) > 10)
texts = Counter(clean(x) for x in raws if to_number(x) is None)
print()
print("=== 原始資料 ===")
print("數值寫法：", "百分比（會除以 100）" if PERCENT else "倍數（直接使用）")
print("非數字內容（一律填 /）：", dict(texts.most_common(8)))

f_cols, f_rows = build_fulfillment(fr)
c_cols, c_rows = build_cash(cv)
f_footer = []
for label in DIV_NAMES.values():
    part = [r for r in f_rows if r[3] == label]
    if part:
        f_footer.append(average_row(part, 4, ["平均值", "", "", label]))
c_footer = [average_row(c_rows, 3, ["平均值", "", ""])]

f_top, f_second = make_headers(f_cols, ["产品名称", "产品种类", "货币", "红利类型"])
c_top, c_second = make_headers(c_cols, ["产品名称", "产品种类", "货币"])
wb = Workbook()
wb.remove(wb.active)
add_sheet(wb, f"{year}分紅實現率", f_top, f_second, f_rows, f_footer, 4,
          [(1, lambda r: r[0]), (2, lambda r: (r[0], r[1])), (3, lambda r: (r[0], r[2]))])
add_sheet(wb, f"{year}總現金價值比率", c_top, c_second, c_rows, c_footer, 3,
          [(1, lambda r: r[0]), (2, lambda r: (r[0], r[1]))])
buf = io.BytesIO()
wb.save(buf)
content = buf.getvalue()

print()
print("=== 生成結果 ===")
print(f"分紅實現率 {len(f_rows)} 行｜總現金價值比率 {len(c_rows)} 行｜檔案 {len(content) // 1024} KB")
print("紅利類型（新檔）：", dict(Counter(r[3] for r in f_rows)))
for row in f_footer + c_footer:
    print("  ", " | ".join(str(x) for x in row))

pages = query_pages()
page = next((p for p in pages if str(year) in page_title(p)), None)

print()
print("=== 與手動整理的檔案核對 ===")
old_wb = load_manual_file(page["id"]) if page else None
if old_wb is None:
    print(f"沒有 {year} 年的手動檔案，略過核對")
else:
    sheets = old_wb.worksheets
    old_f = read_old(sheets[0], 4, (0, 2, 3), len(f_cols))
    print("紅利類型（舊檔）：", dict(Counter(k[2] for k in old_f)))
    compare("分紅實現率", old_f, index_new(f_rows, 4, (0, 2, 3)))
    if len(sheets) > 1:
        compare("總現金價值比率", read_old(sheets[1], 3, (0, 2), len(c_cols)), index_new(c_rows, 3, (0, 2)))

print()
print("=== 上傳到 Notion ===")
if page is None:
    body = {"parent": {"type": "data_source_id", "data_source_id": TARGET},
            "properties": {"title": {"title": [{"text": {"content": f"{year}年"}}]}}}
    r = requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30)
    page = check(r, "建立年度頁面")
    print(f"已建立頁面：{year}年")
else:
    print("使用現有頁面：", page_title(page))
upload_file(page["id"], f"友邦總現金價值&分紅實現率（{year}年）{AUTO}.xlsx", content)
print("完成")
