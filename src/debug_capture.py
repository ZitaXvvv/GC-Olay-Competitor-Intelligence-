"""
诊断脚本：打开浏览器，等你手动搜索一个品牌后按 Enter，
然后自动保存截图和页面 HTML，方便修复选择器。

运行方式:
    python src/debug_capture.py
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import BEBD_URL, COOKIES_FILE

from playwright.sync_api import sync_playwright

OUT_DIR = Path(__file__).parent.parent / "log" / "debug"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def load_cookies(context):
    if COOKIES_FILE.exists():
        with open(COOKIES_FILE, encoding="utf-8") as f:
            context.add_cookies(json.load(f))
        print("已加载 Cookie")
    else:
        print("无 Cookie 文件，请在浏览器中手动登录")


def capture(page, label: str):
    """保存截图 + 完整 HTML"""
    safe = label.replace(" ", "_")
    png = OUT_DIR / f"{safe}.png"
    html = OUT_DIR / f"{safe}.html"
    page.screenshot(path=str(png), full_page=True)
    html.write_text(page.content(), encoding="utf-8")
    print(f"  截图 → {png}")
    print(f"  HTML → {html}")


with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False)
    context = browser.new_context(
        locale="zh-CN",
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    load_cookies(context)
    page = context.new_page()
    page.goto(BEBD_URL)
    page.wait_for_timeout(2000)

    print("\n" + "=" * 60)
    print("【第一步】主页已打开，先不要做任何操作，按 Enter 抓主页搜索框")
    print("=" * 60)
    input("按 Enter ▶ ")

    print("\n=== 主页 URL ===")
    print(page.url)

    print("\n=== 主页所有 input 元素 ===")
    import re as _re
    hp_html = page.content()
    hp_inputs = _re.findall(r'<input[^>]*>', hp_html)
    for inp in hp_inputs[:10]:
        print(inp[:200])

    capture(page, "00_homepage")

    print("\n" + "=" * 60)
    print("【第二步】现在在浏览器里搜索「珀莱雅」，等结果出来后按 Enter")
    print("=" * 60)
    input("搜索结果页加载好后按 Enter ▶ ")

    print("\n正在保存当前页面状态...")
    capture(page, "01_search_result")

    # ── 排序 Tab 诊断 ──
    print("\n=== 排序 Tab 诊断（查找"备案时间"等元素）===")
    sort_info = page.evaluate("""
        () => {
            const all = [...document.querySelectorAll('span, li, div, a, label, button')];
            const results = [];
            for (const el of all) {
                if (el.children.length > 2) continue;
                const txt = el.textContent.trim();
                if (['综合排序','安全','用户评分','备案时间','美修指数'].some(k => txt === k || txt.startsWith(k))) {
                    const r = el.getBoundingClientRect();
                    results.push({
                        tag: el.tagName,
                        cls: el.className.slice(0, 80),
                        txt: txt.slice(0, 30),
                        x: Math.round(r.x),
                        y: Math.round(r.y),
                        w: Math.round(r.width),
                        h: Math.round(r.height),
                        inTable: !!el.closest('table, .ant-table, th'),
                        parent: el.parentElement ? (el.parentElement.tagName + '.' + (el.parentElement.className||'').slice(0,50)) : '',
                    });
                }
            }
            return results;
        }
    """)
    for item in sort_info:
        print(f"  [{item['txt']:10s}] <{item['tag']} class=\"{item['cls']}\">  pos=({item['x']},{item['y']}) size={item['w']}x{item['h']}  inTable={item['inTable']}  parent={item['parent']}")
    if not sort_info:
        print("  ⚠️  找不到任何排序 Tab 元素（可能还未加载，或选择器需要调整）")

    # 打印页面里所有 table、tr、td 数量，以及常见选择器命中情况
    print("\n=== 页面元素诊断 ===")
    checks = {
        "table":                 "table",
        "tbody tr":              "tbody tr",
        ".ant-table-row":        ".ant-table-row",
        ".el-table__row":        ".el-table__row",
        "tr[class*='row']":      "tr[class*='row']",
        "div[class*='row']":     "div[class*='row']",
        "td":                    "td",
    }
    for name, sel in checks.items():
        count = page.locator(sel).count()
        print(f"  {name:30s}: {count} 个")

    # 如果有 table，打印第一行的 td 文本
    if page.locator("table tbody tr").count() > 0:
        print("\n=== 第一个 <table> 第一行各列内容 ===")
        first_row = page.locator("table tbody tr").first
        cells = first_row.locator("td").all()
        for i, cell in enumerate(cells):
            print(f"  cells[{i}] = {cell.inner_text().strip()[:60]!r}")
    elif page.locator(".ant-table-row").count() > 0:
        print("\n=== 第一个 .ant-table-row 各列内容 ===")
        first_row = page.locator(".ant-table-row").first
        cells = first_row.locator("td").all()
        for i, cell in enumerate(cells):
            print(f"  cells[{i}] = {cell.inner_text().strip()[:60]!r}")

    print("\n按 Enter 关闭浏览器...")
    input()
    browser.close()

print(f"\n诊断文件已保存到: {OUT_DIR}")
print("请把上面的输出文字复制给我，我来修复选择器。")
