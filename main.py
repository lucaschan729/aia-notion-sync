import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "playwright"])
subprocess.check_call(
    [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)
from playwright.sync_api import sync_playwright

BASE = "https://www.aia.com.hk/{lang}/products/further-product-information/participating-products/{slug}"
PAGES = [
    ("分紅實現率（繁體）", BASE.format(lang="zh-hk", slug="fulfillment-ratio")),
    ("總現金價值比率（繁體）", BASE.format(lang="zh-hk", slug="total-cash-value-ratio")),
    ("分紅實現率（簡體）", BASE.format(lang="zh-cn", slug="fulfillment-ratio")),
]
KEYWORDS = ["活出精彩", "才雋", "才隽", "倍之寶", "倍之宝"]
SKIP = ("google", "analytics", "demdex", "facebook", "doubleclick", "omtrdc",
        "tiktok", "linkedin", "hotjar", "cookielaw", "onetrust", "clarity")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")


def short(text, n):
    return re.sub(r"\s+", " ", text or "")[:n]


def probe(browser, label, url, detail):
    print()
    print(f"=== {label} ===")
    ctx = browser.new_context(user_agent=UA, locale="zh-HK", viewport={"width": 1366, "height": 900})
    page = ctx.new_page()
    responses = []
    page.on("response", lambda r: responses.append(r))
    try:
        resp = page.goto(url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(10000)
    except Exception as e:
        print("打開失敗：", short(str(e), 200))
        ctx.close()
        return

    print("狀態：", resp.status if resp else "?", "| 標題：", short(page.title(), 60))
    print("最終網址：", page.url[:120])
    html = page.content()
    body = page.inner_text("body")
    print("頁面文字長度：", len(body))
    print("HTML 內含產品名稱：", [k for k in KEYWORDS if k in html] or "沒有")
    m = re.search(r"申[報报]年度.{0,30}", body)
    print("申報年度：", short(m.group(0), 50) if m else "找不到")
    print("表格數：", page.locator("table").count(),
          "| iframe：", [f.url[:80] for f in page.frames[1:]][:5])

    selects = page.locator("select")
    print("下拉選單數：", selects.count())
    for i in range(min(selects.count(), 4)):
        opts = selects.nth(i).locator("option").all_inner_texts()
        print(f"  選單{i + 1}：{len(opts)} 項，例如 {[short(o, 20) for o in opts[:4]]}")

    btn = page.get_by_text(re.compile("展示[詳详]情"))
    print("「展示詳情」按鈕數：", btn.count())
    if detail and btn.count():
        snippet = btn.first.evaluate(
            "el => { let n = el; for (let i = 0; i < 4 && n.parentElement; i++) n = n.parentElement; return n.outerHTML; }"
        )
        print("按鈕附近 HTML：", short(snippet, 900))

    print("資料請求：")
    shown = 0
    for r in responses:
        if r.request.resource_type not in ("xhr", "fetch"):
            continue
        if any(s in r.url for s in SKIP):
            continue
        try:
            text = r.text()
        except Exception:
            text = ""
        flag = "  ← 含產品名稱" if any(k in text for k in KEYWORDS) else ""
        print(f"  [{r.status}] {r.url[:130]} | {len(text)} 字{flag}")
        if flag:
            print("     內容開頭：", short(text, 250))
        shown += 1
        if shown >= 20:
            break
    if not shown:
        print("  （沒有）")
    ctx.close()


with sync_playwright() as p:
    browser = p.chromium.launch()
    for i, (label, url) in enumerate(PAGES):
        probe(browser, label, url, i == 0)
    browser.close()
