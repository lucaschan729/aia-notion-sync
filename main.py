import datetime
import io
import json
import math
import os
import re
import statistics
import subprocess
import sys

import requests

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
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
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"   # 產品名稱顏色（深灰黑）
CUR_COLOR = "FF7030A0"    # 貨幣顏色（紫）
SEP_COLOR = "FFE9ECEF"    # 產品之間分隔行的顏色（淺灰）
CAT_COLORS = ["FF2F5597", "FF2E7D32", "FFC55A11", "FF00838F", "FF7F6000"]
CAT_SEEN = {}


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


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
    return "/" if n is None else n


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


def to_fractions(rows, n_fixed):
    """網站數字如果是 95 這種寫法，換成 0.95，方便用百分比格式顯示。"""
    nums = [v for r in rows for v in r[n_fixed:] if is_num(v)]
    scale = 100 if nums and statistics.median(nums) > 3 else 1
    for r in rows:
        for c in range(n_fixed, len(r)):
            if is_num(r[c]):
                r[c] = round(r[c] / scale, 6)


def report_year(obj):
    m = re.search(r"\d{4}", zh(obj.get("report_year")))
    return m.group(0) if m else ""


def last_is_before(cols):
    return len(cols) >= 2 and not re.fullmatch(r"\d{4}", clean(cols[-1]))


def make_header(cols, fixed):
    n = len(cols)
    out = list(fixed)
    for i, c in enumerate(cols):
        if i == n - 1 and last_is_before(cols):
            out.append(f"保单年度{n - 1}+\n{clean(c)}")
        else:
            out.append(f"保单年度{i + 1}\n{clean(c)}")
    return out


# ---------- 平均值 ----------
def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


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
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": [clean(c) for c in cols], "year": [year_line]}


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


def heat(cell, v, base):
    """100% 或以上綠底；低於 80% 紅底；沒有數字用淺灰字。"""
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"


def cat_color(cat):
    if cat not in CAT_SEEN:
        CAT_SEEN[cat] = CAT_COLORS[len(CAT_SEEN) % len(CAT_COLORS)]
    return CAT_SEEN[cat]


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    """按最長的文字決定欄闊；超過上限的文字會自動縮小字體。"""
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 14]
    caps = [52, 24, 12, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(n_fixed)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            # 產品之間的分隔行：整行合併
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = Font(color=cat_color(v))
            elif c == 3:
                cell.font = CUR_FONT
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, today, items):
    """items：[(簡稱, 工作表名稱, 資料行, 固定欄數)]"""
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    banner(1, f"友邦分紅表現概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：友邦香港網站｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for _, title, rows, n_fixed in items:
        nums = numbers(rows, n_fixed)
        low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
        vals = [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low]
        r += 1
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
        link = ws.cell(row=r, column=1)
        link.font = LINK_FONT
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{title}'!A1")
        ws.cell(row=r, column=5).number_format = "0.00%"
        heat(ws.cell(row=r, column=5), vals[4], WHITE)
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "按產品種類的整體平均值")
    r += 1
    table_head(r, ["產品種類"] + [short for short, _, _, _ in items])
    cats = []
    for _, _, rows, _ in items:
        for row in rows:
            if row[1] not in cats:
                cats.append(row[1])
    for cat in cats:
        r += 1
        cell = ws.cell(row=r, column=1, value=cat or "（未分類）")
        cell.border = BORDER
        cell.alignment = LEFT_FIT
        cell.font = Font(bold=True, color=cat_color(cat))
        for k, (_, _, rows, n_fixed) in enumerate(items):
            v = mean_or_slash(numbers([row for row in rows if row[1] == cat], n_fixed))
            c2 = ws.cell(row=r, column=k + 2, value=v)
            c2.border = BORDER
            c2.alignment = CENTER_FIT
            c2.number_format = "0.00%"
            heat(c2, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：網站沒有提供該年度的數字")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "使用提示")
    tips = ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
            "每個工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
            "向右或向下捲動時，標題列和產品名稱會固定不動。",
            "平均值摘要在每個工作表最底部。",
            "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"]
    for t in tips:
        r += 1
        cell = ws.cell(row=r, column=1, value="・" + t)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = NOTE_FONT
        cell.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H


# ---------- Notion ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    r = requests.get(f"{API}/data_sources/{ident}", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    print(f"找不到友邦的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
          "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")
    raise SystemExit(1)


def query_pages(source):
    r = requests.post(f"{API}/data_sources/{source}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
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
            out.append((block["id"], block["file"].get("name", "")))
    return out


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name in file_blocks(page_id) if AUTO in name]
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


# ---------- 主流程 ----------
f_obj = fetch_json("fulfillment-ratio.json")
c_obj = fetch_json("total-cash-value-ratio.json")
f_year, c_year = report_year(f_obj), report_year(c_obj)
if f_year and c_year and f_year != c_year:
    raise SystemExit(f"兩份資料的申報年度不同：{f_year} / {c_year}，暫不更新")
year = f_year or c_year
if not year:
    raise SystemExit("讀不到申報年度，暫不更新")

f_cols, f_rows = build_fulfillment(f_obj)
c_cols, c_rows = build_cash(c_obj)
to_fractions(f_rows, 4)
to_fractions(c_rows, 3)

print()
print("=== 讀取結果 ===")
print(f"報告年度 {year}｜總現金價值比率 {len(c_rows)} 行｜分紅實現率 {len(f_rows)} 行")
print("分紅實現率年度欄：", " ".join(clean(c) for c in f_cols))
print("總現金價值年度欄：", " ".join(clean(c) for c in c_cols))
print("產品種類：", "｜".join(sorted({r[1] or "（空白）" for r in f_rows + c_rows})))
print("紅利類型：", "｜".join(sorted({r[3] for r in f_rows})))
for row in (f_rows[:2] + c_rows[:2]):
    print("  例：", " ".join(str(v) for v in row))
if len(c_cols) < 5 or len(f_cols) < 5 or len(c_rows) < 30 or len(f_rows) < 60:
    raise SystemExit("讀到的資料太少，友邦網站可能改了格式，暫不更新")

f_groups = [("所有分红", f_rows)]
for t in DIV_NAMES.values():
    sub = [r for r in f_rows if r[3] == t]
    if numbers(sub, 4):
        f_groups.append((t, sub))
f_summary = build_summary(4, f_cols, f_groups)
c_summary = build_summary(3, c_cols, [("所有產品", c_rows)])

today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
legend = ("綠色＝100% 或以上；紅色＝低於 80%；「/」＝網站沒有提供該年度的數字。\n"
          f"資料來源：友邦香港網站（報告年度 {year}）｜更新日期：{today}")
c_notes = ("總現金價值比率＝實際總現金價值 ÷ 銷售時說明的總現金價值。"
           "100% 即與銷售時一致，高於 100% 即比銷售時多。定義以友邦網站的說明為準。\n" + legend)
f_notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額。"
           "100% 即與預期一致，高於 100% 即比預期多。定義以友邦網站的說明為準。\n" + legend)

c_title, f_title = f"{year}總現金價值比率", f"{year}分紅實現率"
wb = Workbook()
wb.remove(wb.active)
add_sheet(wb, c_title, f"友邦總現金價值比率（報告年度 {year}）", c_notes,
          make_header(c_cols, ["产品名称", "产品种类", "货币"]), c_rows, c_summary,
          "總現金價值比率", 3, NAVY_HEX)
add_sheet(wb, f_title, f"友邦分紅實現率（報告年度 {year}）", f_notes,
          make_header(f_cols, ["产品名称", "产品种类", "货币", "红利类型"]), f_rows, f_summary,
          "分紅類型", 4, "2E75B6")
add_overview(wb, year, today,
             [("總現金價值比率", c_title, c_rows, 3), ("分紅實現率", f_title, f_rows, 4)])
buf = io.BytesIO()
wb.save(buf)
content = buf.getvalue()
print(f"Excel 檔案 {len(content) // 1024} KB｜工作表：", "、".join(wb.sheetnames))

print()
print("=== 上傳到 Notion ===")
source = resolve_source(TARGET)
pages = query_pages(source)
page = next((p for p in pages if str(year) in page_title(p)), None)
if page is None:
    body = {"parent": {"type": "data_source_id", "data_source_id": source},
            "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
    r = requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30)
    page = check(r, "建立年度頁面")
    print(f"已建立頁面：{year}")
else:
    print("使用現有頁面：", page_title(page))
upload_file(page["id"], f"友邦總現金價值&分紅實現率（{year}年）{AUTO}.xlsx", content)
print("完成")
