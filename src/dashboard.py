"""
CI New SKU Dashboard — Streamlit 前端
月历视图：每格显示产品名（中文 + 英文）、功效、Artwork PDF、mini-POC 链接

运行：
    streamlit run src/dashboard.py
"""

import json
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from config import (
    BRANDS,
    COL_DATE,
    COL_EFFECT,
    COL_INGREDIENTS,
    COL_LABEL_URL,
    COL_NAME,
    COL_PDF_URL,
    COL_POC_URL,
    COL_REG_NUM,
    COL_UPLOAD_DATE,
    EXCEL_PATH,
)

# ─────────────────────────────────────────────
# 页面配置
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CI New SKU · Competitor Intelligence",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TRANSLATION_CACHE = Path(__file__).parent / "translations_cache.json"

# ─────────────────────────────────────────────
# 本地翻译：美妆专用词典（零网络依赖）
# ─────────────────────────────────────────────

# 品牌中英文映射
_BRAND_MAP = {
    "珀莱雅": "PROYA", "谷雨": "Guyu", "欧诗漫": "OSM",
    "科颜氏": "Kiehl's", "欧莱雅": "L'Oréal", "雅诗兰黛": "Estée Lauder",
    "修丽可": "SkinCeuticals", "兰蔻": "Lancôme", "百雀羚": "Pechoin",
    "自然堂": "Chando", "薇诺娜": "Winona", "韩束": "Kans",
    "娇韵诗": "Clarins", "契尔氏": "Kiehl's",
}

# 产品剂型
_FORM_MAP = {
    "精华液": "Serum", "精华": "Serum", "面霜": "Face Cream", "眼霜": "Eye Cream",
    "乳液": "Lotion", "水乳": "Toner & Lotion", "乳": "Milk/Lotion",
    "面膜": "Sheet Mask", "睡眠面膜": "Sleeping Mask",
    "慕斯": "Mousse", "泡沫": "Foam",
    "洁面": "Cleanser", "卸妆": "Makeup Remover",
    "防晒乳": "Sunscreen Lotion", "防晒霜": "Sunscreen Cream",
    "防晒": "Sunscreen", "隔离": "Primer",
    "喷雾": "Mist", "喷": "Mist", "露": "Dew",
    "油": "Oil", "霜": "Cream", "啫喱": "Gel",
    "次抛": "Ampoule", "安瓶": "Ampoule",
    "精粹": "Essence", "素": "Essence",
    "唇膏": "Lip Balm", "唇霜": "Lip Cream",
    "爽肤水": "Toner", "化妆水": "Toner",
    "冻膜": "Jelly Mask", "凝露": "Gel Serum",
    "凝霜": "Gel Cream", "凝冻": "Gel",
}

# 功效关键词
_BENEFIT_MAP = {
    "水润": "Hydrating", "保湿": "Moisturizing", "水光": "Luminous",
    "滋润": "Nourishing", "锁水": "Moisture-Lock",
    "美白": "Brightening", "焕白": "Radiance", "亮白": "Whitening",
    "抗皱": "Anti-Wrinkle", "淡纹": "Wrinkle Reduction",
    "紧致": "Firming", "塑颜": "Sculpting",
    "修护": "Repairing", "修复": "Recovery",
    "舒缓": "Soothing", "镇静": "Calming",
    "控油": "Oil Control", "祛痘": "Acne-Clearing",
    "去角质": "Exfoliating", "焕肤": "Renewal",
    "提亮": "Illuminating", "焕亮": "Brightening",
    "双抗": "Dual Antioxidant", "抗氧": "Antioxidant",
    "源力": "Energy-Boost", "赋能": "Energizing",
    "胶原": "Collagen", "玻尿酸": "Hyaluronic Acid",
    "烟酰胺": "Niacinamide", "视黄醇": "Retinol",
    "特护": "Intensive Care", "舒敏": "Sensitive-Soothing",
    "净澈": "Purifying", "净透": "Clarifying",
    "凝时": "Time Freeze", "冻龄": "Age-Freeze",
    "极润": "Ultra-Moist", "嘭盈": "Plumping",
    "缎光": "Satin Glow", "盈润": "Dewy",
}


def translate_local(zh_name: str) -> str:
    """本地美妆词典翻译，不需要网络。"""
    if not zh_name:
        return ""

    result = zh_name

    # 1. 替换品牌名
    for zh, en in _BRAND_MAP.items():
        result = result.replace(zh, en)

    # 2. 替换剂型（从长到短，避免截断）
    for zh, en in sorted(_FORM_MAP.items(), key=lambda x: -len(x[0])):
        result = result.replace(zh, f" {en} ")

    # 3. 替换功效关键词
    for zh, en in sorted(_BENEFIT_MAP.items(), key=lambda x: -len(x[0])):
        result = result.replace(zh, f" {en} ")

    # 4. 清理多余空格，保留首尾
    parts = [p.strip() for p in result.split() if p.strip()]
    translated = " ".join(parts)

    # 5. 若仍有汉字（未覆盖词），保留原文追加
    import unicodedata
    has_cjk = any(unicodedata.category(c) in ("Lo",) for c in translated)
    return translated if not has_cjk or translated != zh_name else zh_name


@st.cache_resource
def get_translation_cache() -> dict:
    if TRANSLATION_CACHE.exists():
        with open(TRANSLATION_CACHE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def translate_batch(names: list[str], cache: dict, timeout_sec: int = 8) -> dict:
    """批量翻译：优先用本地词典，无需网络。"""
    for name in names:
        if name and name not in cache:
            cache[name] = translate_local(name)
    # 持久化（本地翻译很快，直接保存）
    try:
        with open(TRANSLATION_CACHE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return cache


# ─────────────────────────────────────────────
# 数据加载
# ─────────────────────────────────────────────
@st.cache_data(ttl=1800, show_spinner="⏳ 读取数据…")
def load_data() -> list[dict]:
    if not Path(EXCEL_PATH).exists():
        return []
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    records = []
    for brand_en, brand_cn in BRANDS.items():
        if brand_en not in wb.sheetnames:
            continue
        ws = wb[brand_en]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 5:
                continue
            def g(col):
                return row[col - 1] if len(row) >= col else None
            name_raw  = g(COL_NAME)
            effect    = g(COL_EFFECT)
            notif_raw = g(COL_DATE)
            upload_raw= g(COL_UPLOAD_DATE)
            ingr      = g(COL_INGREDIENTS)
            pdf       = g(COL_PDF_URL)
            label     = g(COL_LABEL_URL)
            poc       = g(COL_POC_URL)

            if not name_raw:
                continue

            notif_date = None
            for val in [notif_raw, upload_raw]:
                if val is None:
                    continue
                if hasattr(val, "date"):
                    notif_date = val.date()
                    break
                for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
                    try:
                        notif_date = datetime.strptime(str(val).strip(), fmt).date()
                        break
                    except ValueError:
                        pass
                if notif_date:
                    break
            if not notif_date:
                continue

            # Classify reg type from any cell
            import re
            reg_type = ""
            for cell in row:
                s = str(cell or "")
                if re.search(r"国妆特字|国妆特进字|卫妆特字", s):
                    reg_type = "特殊注册"
                    break
                if re.search(r"妆网备字|国妆备字|国妆备进字", s):
                    reg_type = "普通备案"
                    break

            def clean(v):
                s = str(v or "").strip()
                return s if s not in ("None", "NA", "") else ""

            records.append({
                "brand_en":   brand_en,
                "brand_cn":   brand_cn,
                "name":       str(name_raw).strip(),
                "effect":     clean(effect),
                "ingredients":clean(ingr),
                "notif_date": notif_date,
                "year":       notif_date.year,
                "month":      notif_date.month,
                "year_month": notif_date.strftime("%Y-%m"),
                "reg_type":   reg_type,
                "pdf_url":    clean(pdf),
                "label_url":  clean(label),
                "poc_url":    clean(poc),
            })
    wb.close()
    return records


# ─────────────────────────────────────────────
# CSS（现代设计风格）
# ─────────────────────────────────────────────
GLOBAL_CSS = """
<style>
  /* ── 全局字体 ── */
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  /* ── 隐藏 Streamlit 默认元素 ── */
  #MainMenu, footer, header { visibility: hidden; }
  .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }

  /* ── 顶部 Hero ── */
  .hero {
    background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 24px;
    color: white;
  }
  .hero h1 { font-size: 28px; font-weight: 700; margin: 0 0 4px 0; letter-spacing: -0.5px; }
  .hero p  { font-size: 14px; color: rgba(255,255,255,0.65); margin: 0; }

  /* ── 统计卡片 ── */
  .stat-row { display: flex; gap: 16px; margin-bottom: 24px; }
  .stat-card {
    flex: 1;
    background: white;
    border: 1px solid #e8ecf0;
    border-radius: 12px;
    padding: 18px 22px;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
  }
  .stat-card .label { font-size: 11px; font-weight: 600; color: #8898aa;
                      text-transform: uppercase; letter-spacing: .6px; }
  .stat-card .value { font-size: 32px; font-weight: 700; color: #1a2b4a;
                      line-height: 1.1; margin-top: 4px; }
  .stat-card.blue  .value { color: #1565c0; }
  .stat-card.green .value { color: #2e7d32; }
  .stat-card.red   .value { color: #c62828; }

  /* ── 日历表格 ── */
  .cal-wrap { overflow-x: auto; border-radius: 12px;
              box-shadow: 0 2px 12px rgba(0,0,0,.08); }
  .cal-table { border-collapse: collapse; width: 100%; min-width: 1100px;
               font-size: 12px; background: white; }

  /* 表头 */
  .cal-th {
    background: #1a2b4a;
    color: white;
    text-align: center;
    padding: 10px 6px;
    font-weight: 600;
    font-size: 11px;
    letter-spacing: .4px;
    white-space: nowrap;
  }
  .cal-th-brand {
    background: #1a2b4a;
    color: white;
    padding: 10px 14px;
    font-weight: 600;
    font-size: 11px;
    text-align: left;
    min-width: 100px;
    position: sticky;
    left: 0;
    z-index: 2;
  }

  /* 品牌列 */
  .cal-brand {
    background: #f8f9fb;
    border-right: 2px solid #e2e8f0;
    padding: 12px 14px;
    font-weight: 700;
    font-size: 12px;
    color: #1a2b4a;
    vertical-align: top;
    min-width: 100px;
    position: sticky;
    left: 0;
    z-index: 1;
  }
  .cal-brand small { display: block; color: #8898aa;
                     font-weight: 400; font-size: 10px; margin-top: 2px; }

  /* 数据格 */
  .cal-cell {
    border: 1px solid #f0f2f5;
    padding: 6px;
    vertical-align: top;
    min-width: 130px;
    max-width: 170px;
    background: white;
  }
  .cal-cell:hover { background: #fafbff; }

  /* 产品卡片 */
  .prod-card {
    background: white;
    border: 1px solid #e8ecf1;
    border-left: 3px solid #1565c0;
    border-radius: 8px;
    padding: 8px 10px;
    margin-bottom: 6px;
    box-shadow: 0 1px 3px rgba(0,0,0,.05);
    transition: box-shadow .15s;
  }
  .prod-card:hover { box-shadow: 0 3px 10px rgba(0,0,0,.1); }
  .prod-card.special { border-left-color: #c62828; }

  .prod-name { font-weight: 600; color: #1a2b4a; font-size: 12px;
               line-height: 1.35; margin-bottom: 2px; }
  .prod-en   { color: #5f7089; font-size: 10px; margin-bottom: 4px;
               font-style: italic; }
  .prod-eff  { color: #637382; font-size: 10px; margin-bottom: 5px;
               display: -webkit-box; -webkit-line-clamp: 2;
               -webkit-box-orient: vertical; overflow: hidden; }
  .prod-meta { display: flex; align-items: center; gap: 5px;
               margin-bottom: 5px; flex-wrap: wrap; }

  /* 徽章 */
  .badge {
    font-size: 9px; font-weight: 700; padding: 2px 7px;
    border-radius: 20px; white-space: nowrap;
  }
  .badge-n  { background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9; }
  .badge-s  { background: #ffebee; color: #c62828; border: 1px solid #ffcdd2; }

  /* 按钮行 */
  .btn-row  { display: flex; gap: 5px; flex-wrap: wrap; margin-top: 4px; }
  .btn {
    font-size: 10px; font-weight: 600;
    padding: 3px 9px; border-radius: 5px;
    text-decoration: none; display: inline-flex;
    align-items: center; gap: 3px; border: none;
    cursor: pointer; white-space: nowrap;
  }
  .btn-art { background: #1565c0; color: white; }
  .btn-art:hover { background: #0d47a1; color: white; }
  .btn-lbl { background: #6a1b9a; color: white; }
  .btn-poc { background: #2e7d32; color: white; }

  /* 月份徽章 */
  .month-pill { background: #3d7bd9; color: white; border-radius: 12px;
                padding: 1px 7px; font-size: 10px; margin-left: 4px;
                font-weight: 600; }

  /* 筛选栏 */
  .filter-bar {
    background: white;
    border: 1px solid #e8ecf0;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,.05);
  }
</style>
"""


def product_card_html(prod: dict, en_name: str) -> str:
    is_special = prod["reg_type"] == "特殊注册"
    badge = (f'<span class="badge badge-s">特殊注册</span>'
             if is_special
             else f'<span class="badge badge-n">普通备案</span>')
    eff = (prod["effect"][:55] + "…") if len(prod["effect"]) > 55 else prod["effect"]

    btns = ""
    if prod["pdf_url"]:
        btns += f'<a href="{prod["pdf_url"]}" target="_blank" class="btn btn-art">🖼 Artwork</a>'
    if prod["label_url"]:
        btns += f'<a href="{prod["label_url"]}" target="_blank" class="btn btn-lbl">🏷 Label</a>'
    if prod["poc_url"]:
        btns += f'<a href="{prod["poc_url"]}" target="_blank" class="btn btn-poc">🧪 POC</a>'

    en_block = f'<div class="prod-en">{en_name[:55]}</div>' if en_name else ""
    eff_block = f'<div class="prod-eff">{eff}</div>' if eff else ""
    btn_block = f'<div class="btn-row">{btns}</div>' if btns else ""

    card_cls = "prod-card special" if is_special else "prod-card"
    return f"""<div class="{card_cls}">
  <div class="prod-name">{prod["name"][:38]}</div>
  {en_block}
  {eff_block}
  <div class="prod-meta">{badge}</div>
  {btn_block}
</div>"""


def build_calendar_html(records, selected_brands, months, trans_cache) -> str:
    grouped: dict[str, dict[str, list]] = {}
    for r in records:
        if r["brand_en"] not in selected_brands:
            continue
        grouped.setdefault(r["brand_en"], {}).setdefault(r["year_month"], []).append(r)

    if not grouped:
        return "<p style='color:#888;padding:20px'>当前筛选条件下无数据</p>"

    html = '<div class="cal-wrap"><table class="cal-table"><thead><tr>'
    html += '<th class="cal-th-brand">品牌</th>'
    for ym, label in months:
        total = sum(len(grouped.get(b, {}).get(ym, [])) for b in selected_brands)
        pill = f'<span class="month-pill">{total}</span>' if total else ""
        html += f'<th class="cal-th">{label}{pill}</th>'
    html += "</tr></thead><tbody>"

    for brand_en in selected_brands:
        brand_data = grouped.get(brand_en)
        if not brand_data:
            continue
        brand_cn = BRANDS[brand_en]
        html += f'<tr><td class="cal-brand">{brand_en}<small>{brand_cn}</small></td>'
        for ym, _ in months:
            prods = brand_data.get(ym, [])
            html += '<td class="cal-cell"><div class="cell-grid">'
            for p in prods:
                en = trans_cache.get(p["name"], "")
                html += product_card_html(p, en)
            html += "</div></td>"
        html += "</tr>"

    html += "</tbody></table></div>"
    return html


# ─────────────────────────────────────────────
# 主界面
# ─────────────────────────────────────────────
def main():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

    # Hero 标题
    st.markdown("""
    <div class="hero">
      <h1>🔬 Competitor Intelligence · New SKU Tracker</h1>
      <p>竞品注册 & 备案月历 · 成分 · 功效宣称 · Artwork PDF · 功效证明</p>
    </div>
    """, unsafe_allow_html=True)

    records = load_data()
    if not records:
        st.error("⚠️ 未读到数据，请先运行：`python src/main.py --days 730`")
        st.stop()

    # 批量翻译（8秒超时自动跳过，不阻塞页面）
    trans_cache = get_translation_cache()
    all_names = list({r["name"] for r in records})
    trans_cache = translate_batch(all_names, trans_cache)

    # ── 筛选栏 ──
    all_brand_keys = list(BRANDS.keys())
    all_years = sorted({r["year"] for r in records}, reverse=True)

    with st.container():
        fc1, fc2, fc3 = st.columns([3, 1, 1.5])
        with fc1:
            selected_brands = st.multiselect(
                "📦 品牌筛选",
                all_brand_keys,
                default=all_brand_keys,
                label_visibility="collapsed",
                placeholder="选择品牌…",
            )
        with fc2:
            selected_year = st.selectbox("📅 年份", all_years, label_visibility="collapsed")
        with fc3:
            reg_type_filter = st.selectbox(
                "🏷 类型",
                ["全部", "普通备案（含'备'）", "特殊注册（含'特'）"],
                label_visibility="collapsed",
            )

    # 过滤
    filtered = [
        r for r in records
        if r["brand_en"] in selected_brands
        and r["year"] == selected_year
        and (
            reg_type_filter == "全部"
            or (reg_type_filter.startswith("普通") and r["reg_type"] == "普通备案")
            or (reg_type_filter.startswith("特殊") and r["reg_type"] == "特殊注册")
        )
    ]

    # ── 统计摘要 ──
    total   = len(filtered)
    normal  = sum(1 for r in filtered if r["reg_type"] == "普通备案")
    special = sum(1 for r in filtered if r["reg_type"] == "特殊注册")
    brands_active = len({r["brand_en"] for r in filtered})

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-card blue">
        <div class="label">Total Products</div>
        <div class="value">{total}</div>
      </div>
      <div class="stat-card green">
        <div class="label">🟢 普通备案</div>
        <div class="value">{normal}</div>
      </div>
      <div class="stat-card red">
        <div class="label">🔴 特殊注册</div>
        <div class="value">{special}</div>
      </div>
      <div class="stat-card">
        <div class="label">Active Brands</div>
        <div class="value">{brands_active}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 月历 ──
    month_labels = ["Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"]
    months = [(f"{selected_year}-{m:02d}", month_labels[m-1]) for m in range(1, 13)]

    calendar_html = build_calendar_html(filtered, selected_brands, months, trans_cache)
    st.markdown(calendar_html, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 原始数据下载 ──
    with st.expander("📊 原始数据 / 下载 CSV"):
        import pandas as pd
        df = pd.DataFrame(filtered).drop(columns=["year","month"], errors="ignore")
        # Add English names
        df["name_en"] = df["name"].map(lambda x: trans_cache.get(x, ""))
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("⬇️ 下载 CSV",
                           csv,
                           file_name=f"ci_newsku_{selected_year}.csv",
                           mime="text/csv")


if __name__ == "__main__":
    main()


sys.path.insert(0, str(Path(__file__).parent))
from config import (
    BRANDS,
    COL_DATE,
    COL_EFFECT,
    COL_INGREDIENTS,
    COL_LABEL_URL,
    COL_NAME,
    COL_PDF_URL,
    COL_POC_URL,
    COL_REG_NUM,
    COL_UPLOAD_DATE,
    EXCEL_PATH,
)

# ─────────────────────────────────────────────
# 页面配置
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CI New SKU Dashboard",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

TRANSLATION_CACHE = Path(__file__).parent / "translations_cache.json"


# ─────────────────────────────────────────────
# 翻译缓存
# ─────────────────────────────────────────────
@st.cache_resource
def load_translation_cache() -> dict:
    if TRANSLATION_CACHE.exists():
        with open(TRANSLATION_CACHE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_translation_cache(cache: dict):
    with open(TRANSLATION_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def translate_name(zh_name: str, cache: dict) -> str:
    if not zh_name or not zh_name.strip():
        return ""
    key = zh_name.strip()
    if key in cache:
        return cache[key]
    try:
        from deep_translator import GoogleTranslator
        en = GoogleTranslator(source="zh-CN", target="en").translate(key)
        cache[key] = en or ""
        save_translation_cache(cache)
        return en or ""
    except Exception:
        return ""


# ─────────────────────────────────────────────
# 数据加载
# ─────────────────────────────────────────────
@st.cache_data(ttl=1800, show_spinner="读取 Excel 数据…")
def load_data() -> list[dict]:
    if not Path(EXCEL_PATH).exists():
        return []

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    records = []

    for brand_en, brand_cn in BRANDS.items():
        if brand_en not in wb.sheetnames:
            continue
        ws = wb[brand_en]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 5:
                continue
            name_raw   = row[COL_NAME       - 1] if len(row) >= COL_NAME       else None
            effect_raw = row[COL_EFFECT     - 1] if len(row) >= COL_EFFECT     else None
            notif_raw  = row[COL_DATE       - 1] if len(row) >= COL_DATE       else None
            upload_raw = row[COL_UPLOAD_DATE- 1] if len(row) >= COL_UPLOAD_DATE else None
            reg_raw    = row[COL_REG_NUM    - 1] if len(row) >= COL_REG_NUM    else None
            ingr_raw   = row[COL_INGREDIENTS- 1] if len(row) >= COL_INGREDIENTS else None
            pdf_raw    = row[COL_PDF_URL    - 1] if len(row) >= COL_PDF_URL    else None
            label_raw  = row[COL_LABEL_URL  - 1] if len(row) >= COL_LABEL_URL  else None
            poc_raw    = row[COL_POC_URL    - 1] if len(row) >= COL_POC_URL    else None

            if not name_raw:
                continue

            # 日期：优先用备案时间(D)，再用上传时间(A)
            notif_date = None
            for val in [notif_raw, upload_raw]:
                if val is None:
                    continue
                if hasattr(val, "date"):
                    notif_date = val.date()
                    break
                for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
                    try:
                        notif_date = datetime.strptime(str(val).strip(), fmt).date()
                        break
                    except ValueError:
                        pass
                if notif_date:
                    break

            if not notif_date:
                continue

            reg_str  = str(reg_raw or "")
            reg_type = "特殊注册" if "特" in reg_str else ("普通备案" if "备" in reg_str else "")

            def clean(v):
                s = str(v or "").strip()
                return s if s not in ("None", "NA", "") else ""

            records.append({
                "brand_en":   brand_en,
                "brand_cn":   brand_cn,
                "name":       str(name_raw).strip(),
                "effect":     clean(effect_raw),
                "ingredients":clean(ingr_raw),
                "notif_date": notif_date,
                "year":       notif_date.year,
                "month":      notif_date.month,
                "year_month": notif_date.strftime("%Y-%m"),
                "reg_num":    reg_str,
                "reg_type":   reg_type,
                "pdf_url":    clean(pdf_raw),
                "label_url":  clean(label_raw),
                "poc_url":    clean(poc_raw),
            })

    wb.close()
    return records


# ─────────────────────────────────────────────
# HTML 产品卡片
# ─────────────────────────────────────────────
CSS = """
<style>
.cal-wrap   { overflow-x: auto; }
.cal-table  { border-collapse: collapse; width: 100%; font-size: 12px; }
.cal-th     { background:#1F5C99; color:white; text-align:center;
              padding:6px 4px; white-space:nowrap; min-width:100px; }
.cal-brand  { background:#f5f5f5; padding:8px 6px; font-weight:bold;
              white-space:nowrap; min-width:80px; vertical-align:top;
              border:1px solid #ddd; }
.cal-cell   { border:1px solid #e8e8e8; padding:4px; vertical-align:top;
              min-width:120px; max-width:200px; }
.prod-card  { background:#fff; border:1px solid #e0e0e0; border-radius:6px;
              padding:7px; margin-bottom:5px; }
.prod-name  { font-weight:600; color:#1a1a1a; font-size:12px;
              line-height:1.3; margin-bottom:2px; }
.prod-en    { color:#777; font-style:italic; font-size:10px; margin-bottom:3px; }
.prod-eff   { color:#444; font-size:10px; margin-bottom:3px;
              display:-webkit-box; -webkit-line-clamp:2;
              -webkit-box-orient:vertical; overflow:hidden; }
.badge-n    { background:#d4edda; color:#155724; font-size:9px;
              padding:1px 5px; border-radius:10px; }
.badge-s    { background:#f8d7da; color:#721c24; font-size:9px;
              padding:1px 5px; border-radius:10px; }
.reg-num    { color:#888; font-size:9px; margin-left:4px; }
.btn-row    { margin-top:4px; display:flex; gap:4px; flex-wrap:wrap; }
.btn-art    { background:#0d6efd; color:white; border:none; border-radius:4px;
              padding:3px 8px; font-size:10px; cursor:pointer;
              text-decoration:none; display:inline-block; }
.btn-lbl    { background:#6f42c1; color:white; border:none; border-radius:4px;
              padding:3px 8px; font-size:10px; cursor:pointer;
              text-decoration:none; display:inline-block; }
.btn-poc    { background:#198754; color:white; border:none; border-radius:4px;
              padding:3px 8px; font-size:10px; cursor:pointer;
              text-decoration:none; display:inline-block; }
.month-badge{ background:#1F5C99; color:white; border-radius:10px;
              padding:1px 6px; font-size:10px; margin-left:4px; }
</style>
"""


def product_card_html(prod: dict, en_name: str) -> str:
    badge = (f'<span class="badge-s">特殊注册</span>'
             if prod["reg_type"] == "特殊注册"
             else f'<span class="badge-n">普通备案</span>')

    reg_short = prod["reg_num"][:22] if prod["reg_num"] else ""
    eff_short = (prod["effect"][:60] + "…") if len(prod["effect"]) > 60 else prod["effect"]

    btns = ""
    if prod["pdf_url"]:
        btns += f'<a href="{prod["pdf_url"]}" target="_blank" class="btn-art">🖼 Artwork</a>'
    if prod["label_url"]:
        btns += f'<a href="{prod["label_url"]}" target="_blank" class="btn-lbl">🏷 Label</a>'
    if prod["poc_url"]:
        btns += f'<a href="{prod["poc_url"]}" target="_blank" class="btn-poc">🧪 POC</a>'

    en_block = f'<div class="prod-en">{en_name[:50]}</div>' if en_name else ""
    eff_block = f'<div class="prod-eff">{eff_short}</div>' if eff_short else ""

    return f"""
<div class="prod-card">
  <div class="prod-name">{prod["name"][:35]}</div>
  {en_block}
  {eff_block}
  <div>{badge}<span class="reg-num">{reg_short}</span></div>
  <div class="btn-row">{btns}</div>
</div>"""


def build_calendar_html(
    records: list[dict],
    selected_brands: list[str],
    months: list[tuple],          # [(year_month, label), ...]
    translation_cache: dict,
) -> str:
    # group: brand → year_month → list[prod]
    grouped: dict[str, dict[str, list]] = {}
    for r in records:
        if r["brand_en"] not in selected_brands:
            continue
        grouped.setdefault(r["brand_en"], {}).setdefault(r["year_month"], []).append(r)

    if not grouped:
        return "<p style='color:#888'>当前筛选条件下无数据</p>"

    html = CSS + '<div class="cal-wrap"><table class="cal-table"><tr>'
    html += '<th class="cal-th" style="min-width:80px">品牌</th>'

    for ym, label in months:
        total = sum(len(grouped.get(b, {}).get(ym, [])) for b in selected_brands)
        badge = f'<span class="month-badge">{total}</span>' if total else ""
        html += f'<th class="cal-th">{label}{badge}</th>'
    html += "</tr>"

    for brand_en in selected_brands:
        brand_data = grouped.get(brand_en)
        if not brand_data:
            continue
        brand_cn = BRANDS[brand_en]
        html += f'<tr><td class="cal-brand">{brand_en}<br><small style="color:#666">{brand_cn}</small></td>'

        for ym, _ in months:
            prods = brand_data.get(ym, [])
            html += '<td class="cal-cell">'
            for p in prods:
                en = translate_name(p["name"], translation_cache)
                html += product_card_html(p, en)
            html += "</td>"

        html += "</tr>"

    html += "</table></div>"
    return html


# ─────────────────────────────────────────────
# 成分 / 功效 详情弹窗
# ─────────────────────────────────────────────
def show_product_detail(records: list[dict], translation_cache: dict):
    """侧边栏：搜索产品查看完整成分和功效"""
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 产品详情查询")

    all_names = sorted({r["name"] for r in records if r["name"]})
    chosen = st.sidebar.selectbox("选择产品", ["— 请选择 —"] + all_names)

    if chosen == "— 请选择 —":
        return

    matches = [r for r in records if r["name"] == chosen]
    if not matches:
        return

    prod = matches[0]
    en_name = translate_name(prod["name"], translation_cache)

    st.sidebar.markdown(f"### {prod['name']}")
    if en_name:
        st.sidebar.markdown(f"*{en_name}*")
    st.sidebar.markdown(f"**品牌**: {prod['brand_en']} / {prod['brand_cn']}")
    st.sidebar.markdown(f"**备案号**: `{prod['reg_num']}`")
    st.sidebar.markdown(f"**类型**: {prod['reg_type']}")
    st.sidebar.markdown(f"**备案日期**: {prod['notif_date']}")

    if prod["effect"]:
        st.sidebar.markdown("**功效宣称**:")
        st.sidebar.markdown(f"> {prod['effect']}")

    if prod["ingredients"]:
        with st.sidebar.expander("📋 完整成分列表"):
            st.write(prod["ingredients"])

    if prod["pdf_url"]:
        st.sidebar.link_button("🖼 打开 Artwork PDF", prod["pdf_url"], use_container_width=True)
    if prod["label_url"]:
        st.sidebar.link_button("🏷 产品标签链接", prod["label_url"], use_container_width=True)
    if prod["poc_url"]:
        st.sidebar.link_button("🧪 mini POC 链接", prod["poc_url"], use_container_width=True)


# ─────────────────────────────────────────────
# 主界面
# ─────────────────────────────────────────────
def main():
    st.title("🔬 CI New SKU Dashboard")
    st.caption("竞品注册/备案月历 · 成分 · 功效 · Artwork PDF")

    records = load_data()
    translation_cache = load_translation_cache()

    if not records:
        st.error(
            "⚠️ 未读到数据。请先运行：\n\n"
            "```\npython src/main.py --days 730\n```"
        )
        st.stop()

    # ── 筛选 ──
    all_years = sorted({r["year"] for r in records}, reverse=True)
    all_brand_keys = list(BRANDS.keys())

    col1, col2, col3, col4 = st.columns([2, 1.5, 1.5, 1.5])
    with col1:
        selected_brands = st.multiselect(
            "品牌", all_brand_keys, default=all_brand_keys,
            placeholder="选择品牌"
        )
    with col2:
        selected_year = st.selectbox("年份", all_years)
    with col3:
        reg_type_filter = st.selectbox("类型", ["全部", "普通备案", "特殊注册"])
    with col4:
        show_en = st.toggle("显示英文名（需网络）", value=False)

    if not show_en:
        # Clear cache to avoid translation calls
        pass

    # Filter
    filtered = [
        r for r in records
        if r["brand_en"] in selected_brands
        and r["year"] == selected_year
        and (reg_type_filter == "全部" or r["reg_type"] == reg_type_filter)
    ]

    # 统计
    total_count = len(filtered)
    normal_count  = sum(1 for r in filtered if r["reg_type"] == "普通备案")
    special_count = sum(1 for r in filtered if r["reg_type"] == "特殊注册")

    m1, m2, m3 = st.columns(3)
    m1.metric("总产品数", total_count)
    m2.metric("🟢 普通备案", normal_count)
    m3.metric("🔴 特殊注册", special_count)

    st.markdown("---")

    # ── 月份生成 ──
    month_labels = [
        "1月", "2月", "3月", "4月", "5月", "6月",
        "7月", "8月", "9月", "10月", "11月", "12月",
    ]
    months = [
        (f"{selected_year}-{m:02d}", month_labels[m-1])
        for m in range(1, 13)
    ]

    # 如果不显示英文名，清空缓存避免网络调用
    if not show_en:
        fake_cache = {r["name"]: "" for r in filtered}
        calendar_html = build_calendar_html(filtered, selected_brands, months, fake_cache)
    else:
        calendar_html = build_calendar_html(filtered, selected_brands, months, translation_cache)

    # ── 月历表格 ──
    st.components.v1.html(calendar_html, height=800, scrolling=True)

    # ── 侧边栏：产品详情 ──
    show_product_detail(records, translation_cache if show_en else {r["name"]: "" for r in records})

    # ── 数据表格（可下载）──
    with st.expander("📊 查看原始数据表格 / 下载 CSV"):
        import pandas as pd
        df = pd.DataFrame(filtered).drop(columns=["year", "month"], errors="ignore")
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "⬇️ 下载 CSV",
            csv,
            file_name=f"ci_newsku_{selected_year}.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
