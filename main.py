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

import requests

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

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
HEAD_ROW = 3
FIRST = 4
HIGH = 1.0
LOW_LIMIT = 0.8


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF1F4E79")
LIGHT = solid("FFDDEBF7")
BAND = solid("FFF3F7FC")
WHITE = solid("FFFFFFFF")
GOOD = solid("FFE2F0D9")
LOW = solid("FFFCE4D6")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
GREY_FONT = Font(color="FFAAAAAA")
TITLE_FONT = Font(bold=True, size=14, color="FF1F4E79")
NOTE_FONT = Font(size=10, color="FF555555")
SECTION_FONT = Font(bold=True, size=12, color="FF1F4E79")
THIN = Side(style="thin", color="FFB0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


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


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


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


def last_is_before(cols):
    return len(cols) >= 2 and not re.fullmatch(r"\d{4}", cols[-1])


def make_header(cols, fixed):
    n = len(cols)
    out = list(fixed)
    for i, c in enumerate(cols):
        if i == n - 1 and last_is_before(cols):
            out.append(f"保单年度{n - 1}+\n{cols[-2]}之前")
        else:
            out.append(f"保单年度{i + 1}\n{c}")
    return out


# ---------- 平均值 ----------
def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def build_summary(n_fixed, cols, groups):
    n = len(cols)
    width = n_fixed + n
    head = []
    for i in range(n):
        if i == 0:
            head.append("所有年份平均值")
        elif i == n - 1:
            head.append(f"持滿{n - 1}年+平均值")
        elif i == n - 2:
            head.append(f"第{n - 1}年平均值")
        else:
            head.append(f"持滿{i + 1}年平均值")
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    year_head = list(cols)
    if last_is_before(cols):
        year_head[-1] = f"{cols[-2]}之前"
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": year_head, "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:.2f}" if is_num(v) else "/"


# ---------- 生成 Excel ----------
def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed):
    ws = wb.create_sheet(title)
    last_col = len(header)
    data_end = HEAD_ROW + len(rows)

    # 標題和說明
    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 60

    # 表頭
    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 40

    # 產品數據
    band = False
    for i, row in enumerate(rows):
        rr = FIRST + i
        if i > 0 and row[0] != rows[i - 1][0]:
            band = not band
        base = BAND if band else WHITE
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT if c == 1 else CENTER
            cell.fill = base
            if c == 1:
                cell.font = BLACK_BOLD
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = "0%"
                    if v >= HIGH:
                        cell.fill = GOOD
                    elif v < LOW_LIMIT:
                        cell.fill = LOW
                else:
                    cell.font = GREY_FONT
    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    # 平均值表
    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = 24
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append(r)
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append(r)
        r += 1
    table_end = r - 1
    for rr in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=rr, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for rr in (table_start, year_head_row):
        ws.row_dimensions[rr].height = 30
    for vr in value_rows:
        light(ws.cell(row=vr, column=lc))
        for c in range(lc + 1, last_col + 1):
            ws.cell(row=vr, column=c).number_format = "0.00%"
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 30

    # 欄寬、凍結、列印
    widths = [40, 12, 14, 14]
    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else 13
        if i == lc:
            w = 16
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


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
        if any("平均" in str(c) for c in row if c is not None):
            break
        cells = list(row) + [None] * (n_fixed + ncols)
        fixed = [norm_text(c) for c in cells[:n_fixed]]
        if not any(fixed):
            continue
        if n_fixed == 4:
            fixed[3] = fixed[3].replace("/红利", "")
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
    print(f"【{label}】新檔 {len(new)} 行｜舊檔 {len(old)} 行")
    print(f"  完全相同 {len(same)}｜名稱不同但數值相同 {len(renamed_same)}｜數值不同 {len(diff)}"
          f"｜只在新檔 {len(left_new)}｜只在舊檔 {len(left_old)}")
    for k, o in diff[:6]:
        print("  不同：", "/".join(k)[:50])
        print("    舊：", " ".join(old[o]))
        print("    新：", " ".join(new[k]))


def to_frac(v):
    if is_num(v):
        return float(v)
    if isinstance(v, str) and v.strip().endswith("%"):
        try:
            return float(v.strip()[:-1]) / 100
        except ValueError:
            return None
    return None


def old_summary(ws, col, ncols):
    section, out = None, {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        cells = list(row) + [None] * (col + ncols + 2)
        label = norm_text(cells[col]).replace("紅", "红").replace("/红利", "")
        if label == "分红類型":
            section = "cum"
        elif label == "當年平均值":
            section = "year"
        elif section and label:
            out[(section, label)] = [to_frac(c) for c in cells[col + 1:col + 1 + ncols]]
    return out


def check_summary(old, summary):
    mine = [("cum", line) for line in summary["cum"]] + [("year", line) for line in summary["year"]]
    for section, line in mine:
        name = ("累計 " if section == "cum" else "當年 ") + line[0]
        vals = line[1:]
        o = old.get((section, line[0]))
        if o is None:
            print(f"  {name}：舊檔沒有這一行")
            continue
        ok = 0
        for a, b in zip(o, vals):
            if is_num(a) and is_num(b):
                ok += abs(a - b) < 0.00006
            elif not is_num(a) and not is_num(b):
                ok += 1
        print(f"  {name}：{ok}/{len(vals)} 格相同")
        if ok < len(vals):
            print("    舊：", " ".join(fmt_pct(x) for x in o))
            print("    新：", " ".join(fmt_pct(x) for x in vals))


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

f_cols, f_rows = build_fulfillment(fr)
c_cols, c_rows = build_cash(cv)
f_groups = [("所有分红", f_rows)] + [(label, [r for r in f_rows if r[3] == label]) for label in DIV_NAMES.values()]
f_summary = build_summary(4, f_cols, f_groups)
c_summary = build_summary(3, c_cols, [("所有产品", c_rows)])

today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
legend = ("「/」＝沒有數字（產品當時尚未推出、已停售，或未有數字可顯示）。"
          "綠色＝100% 或以上；橙色＝低於 80%。\n"
          f"資料來源：友邦香港網站（申報年度 {year}）｜更新日期：{today}")
f_notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時預期的金額。"
           "100% 即與預期一致，高於 100% 即比預期多。\n" + legend)
c_notes = ("總現金價值比率＝實際總現金價值 ÷ 銷售時預期的總現金價值。"
           "100% 即與預期一致，高於 100% 即比預期多。\n" + legend)

f_header = make_header(f_cols, ["产品名称", "产品种类", "货币", "红利类型"])
c_header = make_header(c_cols, ["产品名称", "产品种类", "货币"])
wb = Workbook()
wb.remove(wb.active)
add_sheet(wb, f"{year}分紅實現率", f"友邦分紅實現率（申報年度 {year}）", f_notes,
          f_header, f_rows, f_summary, "分紅類型", 4)
add_sheet(wb, f"{year}總現金價值比率", f"友邦總現金價值比率（申報年度 {year}）", c_notes,
          c_header, c_rows, c_summary, "總現金價值比率", 3)
buf = io.BytesIO()
wb.save(buf)
content = buf.getvalue()

print()
print("=== 生成結果 ===")
print(f"分紅實現率 {len(f_rows)} 行｜總現金價值比率 {len(c_rows)} 行｜檔案 {len(content) // 1024} KB")

pages = query_pages()
page = next((p for p in pages if str(year) in page_title(p)), None)

print()
print("=== 與手動整理的檔案核對 ===")
old_wb = load_manual_file(page["id"]) if page else None
if old_wb is None:
    print(f"沒有 {year} 年的手動檔案，略過核對")
else:
    sheets = old_wb.worksheets
    compare("分紅實現率", read_old(sheets[0], 4, (0, 2, 3), len(f_cols)), index_new(f_rows, 4, (0, 2, 3)))
    if len(sheets) > 1:
        compare("總現金價值比率", read_old(sheets[1], 3, (0, 2), len(c_cols)), index_new(c_rows, 3, (0, 2)))
    print("【分紅實現率平均值】")
    check_summary(old_summary(sheets[0], 3, len(f_cols)), f_summary)

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
