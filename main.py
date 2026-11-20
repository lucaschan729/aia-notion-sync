import json
import subprocess
import sys

import requests

BASE = "https://www.aia.com.hk/content/dam/hk-wise/json/further-product-information/2026/"
PAGE = "https://www.aia.com.hk/zh-hk/products/further-product-information/participating-products/fulfillment-ratio"
FILES = [
    ("分紅實現率", "fulfillment-ratio.json"),
    ("總現金價值比率", "total-cash-value-ratio.json"),
]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
MAX_LINES = 80


def short(value, n=40):
    return str(value).replace("\n", " ")[:n]


def kind(value):
    if isinstance(value, dict):
        return f"字典 {len(value)} 鍵"
    if isinstance(value, list):
        return f"列表 {len(value)} 項"
    return type(value).__name__


def describe(obj, indent, out):
    if len(out) >= MAX_LINES:
        return
    pad = "  " * indent
    if isinstance(obj, dict):
        if "zh-cn" in obj:
            out.append(f"{pad}（多語言，只顯示 zh-cn）")
            describe(obj["zh-cn"], indent, out)
            return
        keys = list(obj.keys())
        show = keys if len(keys) <= 14 else keys[:3]
        for k in show:
            v = obj[k]
            if isinstance(v, (dict, list)):
                out.append(f"{pad}{short(k, 30)}: {kind(v)}")
                describe(v, indent + 1, out)
            else:
                out.append(f"{pad}{short(k, 30)} = {short(v)}")
        if len(keys) > len(show):
            out.append(f"{pad}…共 {len(keys)} 個鍵")
    elif isinstance(obj, list):
        if not obj:
            return
        if all(not isinstance(x, (dict, list)) for x in obj):
            out.append(f"{pad}{short(obj[:14], 160)}")
        else:
            out.append(f"{pad}[第 1 項] {kind(obj[0])}")
            describe(obj[0], indent + 1, out)
    else:
        out.append(f"{pad}{short(obj)}")


def parse(raw):
    return json.loads(raw.decode("utf-8-sig"))


def fetch_direct(name):
    try:
        r = requests.get(BASE + name, headers={"User-Agent": UA, "Referer": PAGE}, timeout=60)
        if r.status_code == 200:
            return parse(r.content)
        print(f"直接下載 {name} 失敗：{r.status_code}")
    except Exception as e:
        print(f"直接下載 {name} 失敗：{short(e, 150)}")
    return None


def fetch_with_browser(names):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "playwright"])
    subprocess.check_call(
        [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    from playwright.sync_api import sync_playwright

    result = {}
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(user_agent=UA, locale="zh-HK")
        page = ctx.new_page()
        page.goto(PAGE, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)
        for name in names:
            r = ctx.request.get(BASE + name)
            if r.status == 200:
                result[name] = parse(r.body())
            else:
                print(f"瀏覽器下載 {name} 失敗：{r.status}")
        browser.close()
    return result


data = {}
for _, name in FILES:
    got = fetch_direct(name)
    if got is not None:
        data[name] = got
missing = [name for _, name in FILES if name not in data]
if missing:
    print("改用瀏覽器下載：", missing)
    data.update(fetch_with_browser(missing))
else:
    print("下載方式：直接下載（不需要瀏覽器）")

for label, name in FILES:
    print()
    print(f"=== {label}：{name} ===")
    obj = data.get(name)
    if obj is None:
        print("沒有取得資料")
        continue
    print("report_year：", obj.get("report_year"))
    print("頂層：", ", ".join(f"{k}（{kind(v)}）" for k, v in obj.items()))
    lines = []
    for k, v in obj.items():
        if isinstance(v, (dict, list)):
            lines.append(f"{k}: {kind(v)}")
            describe(v, 1, lines)
    for line in lines[:MAX_LINES]:
        print(line)
    if len(lines) >= MAX_LINES:
        print("（內容太多，只顯示前 80 行）")
