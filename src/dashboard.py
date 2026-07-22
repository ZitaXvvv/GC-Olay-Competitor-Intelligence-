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
    DOWNLOAD_BASE,
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
# 数据加载（读取当前 + 所有历史月份 Excel，跨文件去重）
# ─────────────────────────────────────────────
import re as _re_global
_REG_PAT_GLOBAL = _re_global.compile(
    r"(妆网备字|国妆备字|国妆特字|国妆特进字|国妆备进字|卫妆特字)"
)


def _excel_files_to_load() -> list[Path]:
    """返回所有 CI_List_Ada*.xlsx 路径，当前文件排最后（优先级最高）"""
    base_dir = Path(EXCEL_PATH).parent
    historical = sorted(base_dir.glob("CI_List_Ada *.xlsx"))
    current = Path(EXCEL_PATH)
    # 排除已在 historical 里的当前文件（名字不同）
    result = [f for f in historical if f != current]
    if current.exists():
        result.append(current)
    return result


def _parse_notif_date(notif_raw, upload_raw):
    from datetime import date as _date, datetime as _datetime
    for val in [notif_raw, upload_raw]:
        if val is None:
            continue
        if hasattr(val, "date") and callable(val.date):
            return val.date()
        if isinstance(val, _date):
            return val
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return _datetime.strptime(str(val).strip(), fmt).date()
            except ValueError:
                pass
    return None


def _load_one_excel(path: Path, seen_regs: set, records: list):
    """从单个 Excel 读取所有品牌数据，去重后追加到 records"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return
    for brand_en, brand_cn in BRANDS.items():
        if brand_en not in wb.sheetnames:
            continue
        ws = wb[brand_en]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 5:
                continue

            def g(col):
                return row[col - 1] if len(row) >= col else None

            name_raw   = g(COL_NAME)
            effect_raw = g(COL_EFFECT)
            notif_raw  = g(COL_DATE)
            upload_raw = g(COL_UPLOAD_DATE)
            ingr_raw   = g(COL_INGREDIENTS)
            pdf_raw    = g(COL_PDF_URL)
            label_raw  = g(COL_LABEL_URL)
            poc_raw    = g(COL_POC_URL)

            if not name_raw:
                continue

            notif_date = _parse_notif_date(notif_raw, upload_raw)
            if not notif_date:
                continue

            # 找备案号（任意列）
            reg_str = ""
            for cell in row:
                s = str(cell or "").strip()
                if s and _REG_PAT_GLOBAL.search(s):
                    reg_str = s
                    break

            # 跨文件去重（优先保留当前文件的数据，historical 已先加入）
            if reg_str and reg_str in seen_regs:
                continue
            if reg_str:
                seen_regs.add(reg_str)

            reg_type = ""
            if _re_global.search(r"国妆特字|国妆特进字|卫妆特字", reg_str):
                reg_type = "特殊注册"
            elif _re_global.search(r"妆网备字|国妆备字|国妆备进字", reg_str):
                reg_type = "普通备案"

            def clean(v):
                s = str(v or "").strip()
                return s if s not in ("None", "NA", "") else ""

            records.append({
                "brand_en":    brand_en,
                "brand_cn":    brand_cn,
                "name":        str(name_raw).strip(),
                "effect":      clean(effect_raw),
                "ingredients": clean(ingr_raw),
                "notif_date":  notif_date,
                "year":        notif_date.year,
                "month":       notif_date.month,
                "year_month":  notif_date.strftime("%Y-%m"),
                "reg_num":     reg_str,
                "reg_type":    reg_type,
                "pdf_url":     clean(pdf_raw),
                "label_url":   clean(label_raw),
                "poc_url":     clean(poc_raw),
                "source_file": path.name,
            })
    wb.close()


@st.cache_data(ttl=1800, show_spinner="⏳ 读取数据…")
def load_data() -> list[dict]:
    files = _excel_files_to_load()
    if not files:
        return []
    records: list[dict] = []
    seen_regs: set[str] = set()
    for f in files:
        _load_one_excel(f, seen_regs, records)
    return records
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

  /* ── multiselect 选中标签：去掉红色，改为中性蓝灰色 ── */
  [data-baseweb="tag"] {
    background-color: #e8edf5 !important;
    border: 1px solid #c5cfe0 !important;
    border-radius: 6px !important;
  }
  [data-baseweb="tag"] span { color: #1a2b4a !important; font-weight: 500; font-size: 12px; }
  [data-baseweb="tag"] button {
    color: #5f7089 !important;
    background: transparent !important;
  }
  [data-baseweb="tag"] button:hover { color: #1a2b4a !important; }
  /* multiselect 输入框 展开时高亮 */
  [data-baseweb="select"] [data-testid="stMultiSelectInput"]:focus-within {
    border-color: #3d7bd9 !important;
    box-shadow: 0 0 0 2px rgba(61,123,217,.15) !important;
  }

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
    min-width: 200px;
    background: white;
  }
  .cal-cell:hover { background: #fafbff; }

  /* 每格3列网格 */
  .cell-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 4px;
  }

  /* 产品卡片（默认折叠，点击展开） */
  .prod-card {
    background: white;
    border: 1px solid #e8ecf1;
    border-top: 3px solid #1565c0;
    border-radius: 7px;
    box-sizing: border-box;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
  }
  .prod-card:hover { box-shadow: 0 3px 10px rgba(0,0,0,.10); }
  .prod-card.special { border-top-color: #c62828; }

  /* <details> 折叠 */
  .prod-card details { margin: 0; }
  .prod-card summary {
    cursor: pointer; list-style: none;
    padding: 7px 8px;
    display: flex; flex-direction: column; gap: 3px;
  }
  .prod-card summary::-webkit-details-marker { display: none; }
  .prod-card summary:hover { background: #f5f8ff; }
  .sum-row { display: flex; align-items: flex-start; justify-content: space-between; gap: 4px; }
  .expand-hint { color: #aab4be; font-size: 9px; flex-shrink: 0; }
  details[open] .expand-hint::after { content: '▲'; }
  details:not([open]) .expand-hint::after { content: '▼'; }
  .prod-body { padding: 0 8px 8px; border-top: 1px solid #f0f2f5; }

  .prod-name { font-weight: 600; color: #1a2b4a; font-size: 11px;
               line-height: 1.3; flex: 1; }
  .prod-en   { color: #5f7089; font-size: 10px; margin: 5px 0 3px;
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
    badge = (f'<span class="badge badge-s">特殊</span>'
             if is_special else f'<span class="badge badge-n">备案</span>')
    eff   = (prod["effect"][:60] + "…") if len(prod["effect"]) > 60 else prod["effect"]
    ingr  = prod.get("ingredients", "") or ""
    ingr_s = (ingr[:120] + "…") if len(ingr) > 120 else ingr

    btns = ""
    if prod["pdf_url"]:
        btns += f'<a href="{prod["pdf_url"]}" target="_blank" class="btn btn-art">🖼 Artwork</a>'
    if prod["label_url"]:
        btns += f'<a href="{prod["label_url"]}" target="_blank" class="btn btn-lbl">🏷 Label</a>'
    if prod["poc_url"]:
        btns += f'<a href="{prod["poc_url"]}" target="_blank" class="btn btn-poc">🧪 POC</a>'

    en_block   = f'<div class="prod-en">{en_name[:50]}</div>' if en_name else ""
    eff_block  = f'<div class="prod-eff">{eff}</div>'  if eff else ""
    ingr_block = f'<div class="prod-ingr">{ingr_s}</div>' if ingr_s else ""
    btn_block  = f'<div class="btn-row">{btns}</div>' if btns else ""

    card_cls = "prod-card special" if is_special else "prod-card"
    name_disp = prod["name"][:32]
    pid = prod.get("_pid", "")
    drag_attrs = (f'draggable="true" data-pid="{pid}" ondragstart="cmpDragStart(event)"'
                  if pid else "")
    # Artwork/POC 按钮始终可见（放在 <details> 外部，卡片底部）
    always_btns = ""
    if prod["pdf_url"]:
        always_btns += f'<a href="{prod["pdf_url"]}" target="_blank" class="btn btn-art">&#x1F5BC; Artwork</a>'
    if prod["label_url"]:
        always_btns += f'<a href="{prod["label_url"]}" target="_blank" class="btn btn-lbl">&#x1F3F7; Label</a>'
    if prod["poc_url"]:
        always_btns += f'<a href="{prod["poc_url"]}" target="_blank" class="btn btn-poc">&#x1F9EA; POC</a>'
    always_btn_block = f'<div class="btn-row" style="padding:4px 8px 6px">{always_btns}</div>' if always_btns else ""
    return f"""<div class="{card_cls}" {drag_attrs}>
  <details>
    <summary>
      <div class="sum-row">
        <span class="prod-name">{name_disp}</span>
        <span class="expand-hint"></span>
      </div>
      {badge}
    </summary>
    <div class="prod-body">
      {en_block}{eff_block}{ingr_block}
    </div>
  </details>
  {always_btn_block}
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
# 成分中英文词典（INCI / 通用名）
# ─────────────────────────────────────────────
_INGR_ZH2EN: dict[str, str] = {
    # 基础
    "水": "Water (Aqua)", "去离子水": "Purified Water", "蒸馏水": "Distilled Water",
    "甘油": "Glycerin", "丙二醇": "Propylene Glycol", "丁二醇": "Butylene Glycol",
    "戊二醇": "Pentylene Glycol", "己二醇": "Hexylene Glycol",
    "1,3-丙二醇": "1,3-Propanediol", "双丙甘醇": "Dipropylene Glycol",
    # 活性成分
    "透明质酸钠": "Sodium Hyaluronate", "玻尿酸": "Hyaluronic Acid",
    "烟酰胺": "Niacinamide", "视黄醇": "Retinol", "视黄醇棕榈酸酯": "Retinyl Palmitate",
    "抗坏血酸": "Ascorbic Acid (Vit.C)", "抗坏血酸葡糖苷": "Ascorbyl Glucoside",
    "3-邻-乙基抗坏血酸": "Ethyl Ascorbic Acid", "维生素C": "Vitamin C",
    "生育酚": "Tocopherol (Vit.E)", "生育酚乙酸酯": "Tocopheryl Acetate",
    "神经酰胺 NP": "Ceramide NP", "神经酰胺 EOP": "Ceramide EOP",
    "神经酰胺 AP": "Ceramide AP", "神经酰胺 NS": "Ceramide NS",
    "神经酰胺NP": "Ceramide NP", "神经酰胺EOP": "Ceramide EOP",
    "角鲨烷": "Squalane", "角鲨烯": "Squalene",
    "泛醇": "Panthenol", "尿囊素": "Allantoin",
    "积雪草苷": "Asiaticoside", "积雪草酸": "Asiatic Acid",
    "羟基积雪草酸": "Madecassic Acid",
    "积雪草（CENTELLA ASIATICA）提取物": "Centella Asiatica Extract",
    "积雪草": "Centella Asiatica",
    "胶原": "Collagen", "水解胶原": "Hydrolyzed Collagen",
    "肌肽": "Carnosine", "谷胱甘肽": "Glutathione",
    "虾青素": "Astaxanthin", "辅酶Q10": "Coenzyme Q10",
    "麦角硫因": "Ergothioneine", "凝血酸": "Tranexamic Acid",
    # 保湿/调理
    "甜菜碱": "Betaine", "海藻糖": "Trehalose", "甘露糖醇": "Mannitol",
    "赤藓醇": "Erythritol", "木糖醇": "Xylitol", "山梨醇": "Sorbitol",
    "PEG-40氢化蓖麻油": "PEG-40 Hydrogenated Castor Oil",
    "卡波姆": "Carbomer", "丙烯酸钠": "Sodium Acrylate",
    "黄原胶": "Xanthan Gum", "卡拉胶": "Carrageenan",
    "羟乙基纤维素": "Hydroxyethylcellulose",
    "羟丙基甲基纤维素": "Hydroxypropyl Methylcellulose",
    # 防腐/稳定
    "苯氧乙醇": "Phenoxyethanol", "乙基己基甘油": "Ethylhexylglycerin",
    "1,2-己二醇": "1,2-Hexanediol", "辛酰羟肟酸": "Caprylyl Hydroxamic Acid",
    "对羟基苯乙酮": "p-Hydroxyacetophenone",
    "EDTA二钠": "Disodium EDTA", "EDTA": "EDTA",
    "苯甲酸钠": "Sodium Benzoate",
    # 防晒
    "甲氧基肉桂酸乙基己酯": "Ethylhexyl Methoxycinnamate",
    "二苯酮-3": "Benzophenone-3 (Oxybenzone)",
    "氧化锌": "Zinc Oxide", "二氧化钛": "Titanium Dioxide",
    # 乳化/表活
    "鲸蜡醇": "Cetyl Alcohol", "硬脂醇": "Stearyl Alcohol",
    "鲸蜡硬脂醇": "Cetearyl Alcohol",
    "聚二甲基硅氧烷": "Dimethicone", "环五聚二甲基硅氧烷": "Cyclopentasiloxane",
    "环甲硅油": "Cyclomethicone",
    "椰油酰两性基二乙酸二钠": "Disodium Cocoamphodiacetate",
    "月桂醇聚醚硫酸酯钠": "Sodium Laureth Sulfate",
    "烷基葡萄糖苷": "Alkyl Glucoside", "椰油葡糖苷": "Coco-Glucoside",
    # 植物提取
    "白池花（LIMNANTHES ALBA）籽油": "Meadowfoam Seed Oil",
    "霍霍巴（SIMMONDSIA CHINENSIS）籽油": "Jojoba Seed Oil",
    "向日葵（HELIANTHUS ANNUUS）籽油": "Sunflower Seed Oil",
    "山茶（CAMELLIA JAPONICA）籽油": "Camellia Japonica Seed Oil",
    "角鲨烷": "Squalane",
    "薰衣草（LAVANDULA ANGUSTIFOLIA）花提取物": "Lavender Flower Extract",
    "人参": "Panax Ginseng",
    # 其他常见
    "香精": "Fragrance (Parfum)", "香料": "Fragrance",
    "色素": "Pigment/Colorant",
    "氯化钠": "Sodium Chloride (Salt)",
    "柠檬酸": "Citric Acid", "苹果酸": "Malic Acid",
    "精氨酸": "Arginine", "赖氨酸": "Lysine",
}


def _translate_ingr(zh: str) -> str:
    """翻译单个成分名：精确匹配→局部匹配→原文"""
    key = zh.strip()
    if key in _INGR_ZH2EN:
        return _INGR_ZH2EN[key]
    # 局部匹配（取最长匹配的英文名）
    best = ""
    for zh_k, en_v in _INGR_ZH2EN.items():
        if zh_k in key and len(zh_k) > len(best):
            best = en_v
    # 如有括号内的INCI名，直接提取
    import re
    inci = re.findall(r"[A-Z][A-Z\-\s]+[A-Z]", key)
    if inci:
        return inci[0].title()
    return best


def _parse_ingr(s: str) -> list[str]:
    """解析成分列表（逗号/中文逗号分隔）"""
    import re
    items = re.split(r"[，,、；\[\]【】\n]", s or "")
    return [i.strip().strip("0123456789. ") for i in items if i.strip() and len(i.strip()) > 1]


# ─────────────────────────────────────────────
# 悬浮成分对比层：拖拽 SKU 卡片放入 → JS 端计算成分差异
# （整页月历 + 悬浮层渲染在同一个 components.v1.html 文档里，
#   这样拖拽事件才能在同一 DOM 内被捕获）
# ─────────────────────────────────────────────
_FLOATING_CMP_CSS = """
<style>
  .prod-card[draggable="true"] { cursor: grab; }
  .prod-card[draggable="true"]:active { cursor: grabbing; }
  .prod-card[draggable="true"] summary::before {
    content: "⠿"; color: #c5cfe0; font-size: 10px; margin-right: 4px;
  }

  #cmp-toggle-btn {
    position: fixed; right: 20px; top: 18px;
    background: #1565c0; color: #fff; border: none; border-radius: 30px;
    padding: 8px 16px; font-size: 13px; font-weight: 600; cursor: pointer;
    box-shadow: 0 4px 14px rgba(0,0,0,.25); z-index: 9999;
    display: flex; align-items: center; gap: 6px;
  }
  #cmp-toggle-btn:hover { background: #0d47a1; }

  #cmp-drawer {
    position: fixed; left: 0; right: 0; bottom: 0; height: 0; overflow: hidden;
    background: #fff; border-top: 2px solid #1565c0;
    box-shadow: 0 -6px 20px rgba(0,0,0,.18);
    transition: height .28s ease; z-index: 9998;
  }
  #cmp-drawer.open { height: 40vh; }
  #cmp-drawer-inner { padding: 14px 22px; height: 100%; box-sizing: border-box; overflow-y: auto; }
  #cmp-drawer-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
  #cmp-drawer-header h3 { margin: 0; font-size: 15px; color: #1a2b4a; font-family: 'Inter', sans-serif; }
  .cmp-close-btn { background: #f0f2f5; border: none; font-size: 13px; cursor: pointer;
                   color: #5f7089; border-radius: 6px; padding: 3px 9px; margin-left: 6px; }
  .cmp-close-btn:hover { background: #e2e8f0; }

  .cmp-slots { display: flex; gap: 14px; margin-bottom: 10px; }
  .cmp-slot {
    flex: 1; min-height: 66px; border: 2px dashed #c5cfe0; border-radius: 8px;
    display: flex; align-items: center; justify-content: center; padding: 8px;
    font-size: 12px; color: #9aa5b1; text-align: center; transition: .15s;
  }
  .cmp-slot.filled { border-style: solid; border-color: #1565c0; background: #f5f9ff;
                     color: #1a2b4a; flex-direction: column; align-items: flex-start; text-align: left; }
  .cmp-slot.dragover { background: #e3f2fd; border-color: #1565c0; }
  .cmp-slot .slot-brand { font-weight: 700; font-size: 12px; }
  .cmp-slot .slot-name  { font-size: 11px; margin: 2px 0; }
  .cmp-slot .slot-reg   { font-size: 9px; color: #666; font-family: monospace; }
  .cmp-slot .slot-clear { align-self: flex-end; background: #eee; border: none; border-radius: 4px;
                           font-size: 10px; padding: 2px 6px; cursor: pointer; margin-top: 4px; }

  .cmp-grid { display: grid; grid-template-columns: 70px 1fr 60px 60px; gap: 3px; font-size: 12px; }
  .cmp-head { font-weight: 700; color: #8898aa; font-size: 10px; text-transform: uppercase;
              padding: 4px 6px; background: #f8f9fb; border-radius: 4px; }
  .cmp-row  { display: contents; }
  .cmp-cell { padding: 4px 6px; border-bottom: 1px solid #f5f5f5; }
  .cmp-cell.same    { color: #374151; }
  .cmp-cell.moved-u { color: #1565c0; background: #e3f2fd; border-radius: 3px; }
  .cmp-cell.moved-d { color: #e65100; background: #fff3e0; border-radius: 3px; }
  .cmp-cell.added   { color: #1b5e20; background: #e8f5e9; border-radius: 3px; }
  .cmp-cell.removed { color: #b71c1c; background: #ffebee; border-radius: 3px; text-decoration: line-through; }
</style>
"""


def _build_products_map(filtered: list[dict]) -> str:
    """构建供悬浮层 JS 使用的产品字典（含已翻译成分列表），返回 JSON 字符串"""
    products: dict[str, dict] = {}
    for r in filtered:
        pid = r.get("_pid")
        if not pid:
            continue
        zh_list = _parse_ingr(r.get("ingredients", ""))
        en_list = [(_translate_ingr(x) or "") for x in zh_list]
        # 每条成分：用英文名做匹配键，同时携带中文原名用于双语显示
        ingr_items = [
            {"key": (en or zh).lower(), "label": (f"{zh} / {en}" if en and en != zh else zh)}
            for zh, en in zip(zh_list, en_list)
        ]
        products[pid] = {
            "name": r["name"],
            "brand": r["brand_en"],
            "reg": r["reg_num"],
            "date": str(r["notif_date"]),
            "ingr": ingr_items,
        }
    return json.dumps(products, ensure_ascii=False)


def _build_compare_widget_html(products_json: str) -> str:
    """悬浮开关按钮 + 可展开抽屉（两个拖拽槽 + JS 成分对比表）"""
    return f"""
<button id="cmp-toggle-btn" onclick="cmpToggleDrawer()">🔬 Compare <span id="cmp-badge">0/2</span></button>
<div id="cmp-drawer">
  <div id="cmp-drawer-inner">
    <div id="cmp-drawer-header">
      <h3>🔬 Ingredient Comparison — drag 2 SKU cards here</h3>
      <div>
        <button class="cmp-close-btn" title="Clear both" onclick="cmpClearAll()">🗑 Clear</button>
        <button class="cmp-close-btn" title="Close" onclick="cmpToggleDrawer(false)">✕ Close</button>
      </div>
    </div>
    <div class="cmp-slots">
      <div class="cmp-slot" id="cmp-slot-A"
           ondragover="cmpAllowDrop(event)" ondragleave="cmpDragLeave(event)" ondrop="cmpDropSlot(event,'A')">
        Drag SKU A here
      </div>
      <div class="cmp-slot" id="cmp-slot-B"
           ondragover="cmpAllowDrop(event)" ondragleave="cmpDragLeave(event)" ondrop="cmpDropSlot(event,'B')">
        Drag SKU B here
      </div>
    </div>
    <div id="cmp-diff-area">
      <p style="color:#9aa5b1;font-size:12px;margin-top:6px">
        Drop two SKU cards above to compare ingredients (🟢 added · 🔴 removed · 🔵 moved up ≥5 · 🟠 moved down ≥5).
      </p>
    </div>
  </div>
</div>
<script>
  var PRODUCTS = {products_json};
  var cmpSlots = {{A: null, B: null}};

  function cmpDragStart(ev) {{
    var pid = ev.currentTarget.getAttribute('data-pid');
    if (!pid) {{ ev.preventDefault(); return; }}
    ev.dataTransfer.setData('text/plain', pid);
    ev.dataTransfer.effectAllowed = 'copy';
  }}
  function cmpAllowDrop(ev) {{ ev.preventDefault(); ev.currentTarget.classList.add('dragover'); }}
  function cmpDragLeave(ev) {{ ev.currentTarget.classList.remove('dragover'); }}
  function cmpDropSlot(ev, key) {{
    ev.preventDefault();
    ev.currentTarget.classList.remove('dragover');
    var pid = ev.dataTransfer.getData('text/plain');
    if (!pid || !(pid in PRODUCTS)) return;
    cmpSlots[key] = pid;
    cmpRenderSlots(); cmpRenderDiff(); cmpUpdateBadge();
    var drawer = document.getElementById('cmp-drawer');
    if (!drawer.classList.contains('open')) drawer.classList.add('open');
  }}
  function cmpClearSlot(key) {{
    cmpSlots[key] = null;
    cmpRenderSlots(); cmpRenderDiff(); cmpUpdateBadge();
  }}
  function cmpClearAll() {{
    cmpSlots = {{A: null, B: null}};
    cmpRenderSlots(); cmpRenderDiff(); cmpUpdateBadge();
  }}
  function cmpToggleDrawer(force) {{
    var d = document.getElementById('cmp-drawer');
    if (force === true) d.classList.add('open');
    else if (force === false) d.classList.remove('open');
    else d.classList.toggle('open');
  }}
  function cmpUpdateBadge() {{
    var n = (cmpSlots.A ? 1 : 0) + (cmpSlots.B ? 1 : 0);
    document.getElementById('cmp-badge').innerText = n + '/2';
  }}
  function cmpRenderSlots() {{
    ['A', 'B'].forEach(function(key) {{
      var el = document.getElementById('cmp-slot-' + key);
      var pid = cmpSlots[key];
      if (!pid) {{
        el.className = 'cmp-slot';
        el.innerHTML = 'Drag SKU ' + key + ' here';
      }} else {{
        var p = PRODUCTS[pid];
        el.className = 'cmp-slot filled';
        el.innerHTML = '<div class="slot-brand">' + key + ' · ' + p.brand + '</div>' +
          '<div class="slot-name">' + p.name + '</div>' +
          '<div class="slot-reg">' + (p.reg || '') + '</div>' +
          '<button class="slot-clear" onclick="cmpClearSlot(\\'' + key + '\\')">✕ remove</button>';
      }}
    }});
  }}
  function cmpIngrDiff(listA, listB) {{
    // listA/listB 现在是 {{key, label}} 对象数组
    var posA = {{}}, posB = {{}};
    listA.forEach(function(x, i) {{ if (!(x.key in posA)) posA[x.key] = i; }});
    listB.forEach(function(x, i) {{ if (!(x.key in posB)) posB[x.key] = i; }});
    var rows = [];
    listA.forEach(function(item, i) {{
      if (item.key in posB) {{
        var j = posB[item.key], shift = i - j;
        rows.push({{label: item.label, status: Math.abs(shift) >= 5 ? 'moved' : 'same',
                    posA: i + 1, posB: j + 1, shift: shift}});
      }} else {{
        rows.push({{label: item.label, status: 'removed', posA: i + 1, posB: null, shift: null}});
      }}
    }});
    listB.forEach(function(item, j) {{
      if (!(item.key in posA)) rows.push({{label: item.label, status: 'added', posA: null, posB: j + 1, shift: null}});
    }});
    rows.sort(function(a, b) {{
      var sa = a.posB !== null ? a.posB : (a.posA || 999) + 5000;
      var sb = b.posB !== null ? b.posB : (b.posA || 999) + 5000;
      return sa - sb;
    }});
    return rows;
  }}
  function cmpRenderDiff() {{
    var container = document.getElementById('cmp-diff-area');
    var a = cmpSlots.A ? PRODUCTS[cmpSlots.A] : null;
    var b = cmpSlots.B ? PRODUCTS[cmpSlots.B] : null;
    if (!a || !b) {{
      container.innerHTML = '<p style="color:#9aa5b1;font-size:12px;margin-top:6px">' +
        'Drop two SKU cards above to compare ingredients (🟢 added · 🔴 removed · 🔵 moved up ≥5 · 🟠 moved down ≥5).</p>';
      return;
    }}
    if (!a.ingr.length || !b.ingr.length) {{
      container.innerHTML = '<p style="color:#e65100;font-size:12px;margin-top:6px">' +
        '⚠️ One of the selected products has no ingredient data.</p>';
      return;
    }}
    var diff = cmpIngrDiff(a.ingr, b.ingr);
    var added = diff.filter(function(d) {{ return d.status === 'added'; }}).length;
    var removed = diff.filter(function(d) {{ return d.status === 'removed'; }}).length;
    var up = diff.filter(function(d) {{ return d.status === 'moved' && d.shift > 0; }}).length;
    var down = diff.filter(function(d) {{ return d.status === 'moved' && d.shift < 0; }}).length;

    var rowsHtml = '';
    diff.forEach(function(d) {{
      var cls, arrow, pa, pb;
      if (d.status === 'added') {{ cls = 'added'; arrow = '🟢 NEW'; pa = '—'; pb = d.posB; }}
      else if (d.status === 'removed') {{ cls = 'removed'; arrow = '🔴 OUT'; pa = d.posA; pb = '—'; }}
      else if (d.status === 'moved') {{
        if (d.shift > 0) {{ cls = 'moved-u'; arrow = '🔵 ↑' + d.shift; }}
        else {{ cls = 'moved-d'; arrow = '🟠 ↓' + Math.abs(d.shift); }}
        pa = d.posA; pb = d.posB;
      }} else {{ cls = 'same'; arrow = '='; pa = d.posA; pb = d.posB; }}
      rowsHtml += '<div class="cmp-row">' +
        '<div class="cmp-cell ' + cls + '" style="text-align:center">' + arrow + '</div>' +
        '<div class="cmp-cell ' + cls + '">' + d.label + '</div>' +
        '<div class="cmp-cell ' + cls + '" style="text-align:center">' + pa + '</div>' +
        '<div class="cmp-cell ' + cls + '" style="text-align:center">' + pb + '</div>' +
        '</div>';
    }});

    container.innerHTML =
      '<div style="font-size:12px;margin:6px 0 8px">' +
      '<span style="color:#1b5e20"><b>+' + added + ' Added</b></span> &nbsp; ' +
      '<span style="color:#b71c1c"><b>-' + removed + ' Removed</b></span> &nbsp; ' +
      '<span style="color:#1565c0"><b>↑' + up + ' Moved Up ≥5</b></span> &nbsp; ' +
      '<span style="color:#e65100"><b>↓' + down + ' Moved Down ≥5</b></span>' +
      '</div>' +
      '<div class="cmp-grid">' +
      '<div class="cmp-head">Change</div><div class="cmp-head">Ingredient</div>' +
      '<div class="cmp-head">Pos A</div><div class="cmp-head">Pos B</div>' +
      rowsHtml + '</div>';
  }}

  cmpRenderSlots(); cmpUpdateBadge();
</script>
"""


def _render_pdf_download_section(records: list[dict]):
    """
    扫描本地已下载的 PDF 文件（DOWNLOAD_BASE/{brand}/{name}.pdf 或 特化--{name}.pdf），
    提供 Streamlit 原生 download_button 一键下载。
    也列出 Artwork (H列) 和 mini-POC (J列) 链接供直接访问。
    """
    base = Path(DOWNLOAD_BASE)
    if not base.exists():
        return

    # 收集本地已存在的 PDF（按品牌→产品名建立索引）
    local_pdfs: dict[str, Path] = {}  # key = "{brand_en}:{prod_name}" → Path
    for brand_en in BRANDS:
        folder = base / brand_en
        if not folder.exists():
            continue
        for pdf in folder.glob("*.pdf"):
            # 兼容 "特化--{name}.pdf" 和 "{name}.pdf"
            stem = pdf.stem
            if stem.startswith("特化--"):
                stem = stem[3:]
            local_pdfs[f"{brand_en}:{stem}"] = pdf

    if not local_pdfs and not any(r.get("pdf_url") or r.get("poc_url") for r in records):
        return

    with st.expander(f"⬇ PDF & Links  ({len(local_pdfs)} 个本地 PDF 可下载)"):
        # 按品牌分组显示
        by_brand: dict[str, list] = {}
        for r in records:
            by_brand.setdefault(r["brand_en"], []).append(r)

        for brand_en, prods in by_brand.items():
            # 只显示有链接或有本地文件的品牌
            brand_entries = []
            for p in prods:
                key_exact  = f"{brand_en}:{p['name']}"
                key_prefix = next(
                    (k for k in local_pdfs if k.startswith(f"{brand_en}:") and p["name"][:15] in k),
                    None
                )
                local_path = local_pdfs.get(key_exact) or (local_pdfs.get(key_prefix) if key_prefix else None)
                has_link = p.get("pdf_url") or p.get("poc_url") or local_path
                if has_link:
                    brand_entries.append((p, local_path))

            if not brand_entries:
                continue

            st.markdown(f"**{brand_en} / {BRANDS[brand_en]}** ({len(brand_entries)}件)")
            for p, local_path in brand_entries:
                cols = st.columns([4, 1, 1, 1])
                cols[0].markdown(
                    f"<span style='font-size:12px'>{p['name'][:40]}</span>",
                    unsafe_allow_html=True
                )
                # 本地 PDF 下载
                if local_path:
                    pdf_bytes = local_path.read_bytes()
                    cols[1].download_button(
                        "⬇ PDF",
                        data=pdf_bytes,
                        file_name=local_path.name,
                        mime="application/pdf",
                        key=f"dl_{brand_en}_{p['reg_num']}",
                        use_container_width=True
                    )
                else:
                    cols[1].markdown("<span style='color:#aaa;font-size:11px'>无本地文件</span>",
                                     unsafe_allow_html=True)
                # Artwork 链接（NMPA PDF viewer）
                if p.get("pdf_url"):
                    cols[2].markdown(
                        f"<a href='{p['pdf_url']}' target='_blank' "
                        f"style='font-size:11px;color:#1565c0'>🖼 Artwork</a>",
                        unsafe_allow_html=True
                    )
                # mini-POC 链接
                if p.get("poc_url"):
                    cols[3].markdown(
                        f"<a href='{p['poc_url']}' target='_blank' "
                        f"style='font-size:11px;color:#2e7d32'>🧪 POC</a>",
                        unsafe_allow_html=True
                    )


# ─────────────────────────────────────────────
# 主界面
# ─────────────────────────────────────────────
def main():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

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

    all_brand_keys = list(BRANDS.keys())
    all_years = sorted({r["year"] for r in records}, reverse=True)

    # 品牌显示标签："PROYA / 珀莱雅"
    brand_labels = [f"{k} / {v}" for k, v in BRANDS.items()]
    label_to_key = {f"{k} / {v}": k for k, v in BRANDS.items()}
    default_labels = brand_labels  # 默认全选

    # ── 筛选栏 ──
    c1, c2, c3, c4, c5 = st.columns([3, 2, 1.5, 0.5, 0.5])

    with c1:
        sel_brand_labels = st.multiselect(
            "品牌", brand_labels, default=default_labels,
            label_visibility="collapsed",
            placeholder="Select brands (all by default)…",
        )
        selected_brands = [label_to_key[l] for l in sel_brand_labels if l in label_to_key]
    with c2:
        selected_years = st.multiselect(
            "Year (max 3)", all_years,
            default=all_years[:min(3, len(all_years))],
            max_selections=3,
            label_visibility="collapsed",
            placeholder="Select up to 3 years…",
        )
    with c3:
        reg_type_filter = st.selectbox(
            "Type", ["全部 All", "普通备案 Filing", "特殊注册 Registration"],
            label_visibility="collapsed",
        )
    with c4:
        if st.button("✔ All", use_container_width=True, help="Select all brands"):
            st.session_state["展开品牌"] = True  # trigger rerun with all selected
    with c5:
        show_en = st.toggle("EN", value=False, help="Show English product name translation")

    if not selected_brands:
        selected_brands = all_brand_keys
    if not selected_years:
        selected_years = all_years[:1]

    # ── 过滤 ──
    def _type_match(r):
        if "特殊" in reg_type_filter: return r["reg_type"] == "特殊注册"
        if "备案" in reg_type_filter: return r["reg_type"] == "普通备案"
        return True

    filtered = [r for r in records
                if r["brand_en"] in selected_brands
                and r["year"] in selected_years
                and _type_match(r)]

    # ── 统计摘要 ──
    total   = len(filtered)
    normal  = sum(1 for r in filtered if r["reg_type"] == "普通备案")
    special = sum(1 for r in filtered if r["reg_type"] == "特殊注册")
    brands_n = len({r["brand_en"] for r in filtered})

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-card blue"><div class="label">Total Products</div><div class="value">{total}</div></div>
      <div class="stat-card green"><div class="label">🟢 普通备案</div><div class="value">{normal}</div></div>
      <div class="stat-card red"><div class="label">🔴 特殊注册</div><div class="value">{special}</div></div>
      <div class="stat-card"><div class="label">Active Brands</div><div class="value">{brands_n}</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ── 月历（每个选中年份一张）+ 悬浮成分对比层 ──
    # 翻译缓存（只在 EN 开关打开时填充）
    trans_cache: dict = {}
    if show_en:
        names = [r["name"] for r in filtered]
        trans_cache = translate_batch(names, get_translation_cache())

    # 悬浮层需要和卡片在同一个 DOM 内才能捕获拖拽事件，所以把所有年份的
    # 月历 + 悬浮开关按钮 + 抽屉一起放进一个 components.v1.html 文档里渲染。
    for i, r in enumerate(filtered):
        r["_pid"] = f"p{i}"

    month_labels = ["Jan","Feb","Mar","Apr","May","Jun",
                    "Jul","Aug","Sep","Oct","Nov","Dec"]

    year_sections = ""
    for yr in sorted(selected_years, reverse=True):
        yr_filtered = [r for r in filtered if r["year"] == yr]
        if not yr_filtered:
            continue
        months = [(f"{yr}-{m:02d}", month_labels[m-1]) for m in range(1, 13)]
        cal_html = build_calendar_html(yr_filtered, selected_brands, months, trans_cache)
        year_sections += f"""
        <h3 style="color:#1a2b4a;margin:18px 0 6px;font-family:'Inter',sans-serif">📅 {yr}</h3>
        {cal_html}"""

    products_json = _build_products_map(filtered)
    compare_widget_html = _build_compare_widget_html(products_json)

    full_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
{GLOBAL_CSS}
{_FLOATING_CMP_CSS}
</head><body style="margin:0;padding:8px 4px 70px;background:#f7f9fc">
{year_sections}
{compare_widget_html}
</body></html>"""

    # 固定高度 + 内部滚动条：悬浮按钮/抽屉的 position:fixed 是相对这个
    # iframe 自身可视区域的，只有高度固定、内容用内部滚动条滚动，才能让
    # 悬浮层在"滚动年份/品牌"时始终保持可见。
    FRAME_HEIGHT = 820
    st.iframe(full_html, height=FRAME_HEIGHT)

    # ── PDF 本地下载区 ──
    _render_pdf_download_section(filtered)

    # ── 数据下载 ──
    with st.expander("📊 原始数据 / 下载 CSV"):
        import pandas as pd
        df = pd.DataFrame(filtered).drop(columns=["year", "month", "source_file", "_pid"],
                                         errors="ignore")
        st.dataframe(df, use_container_width=True)
        st.download_button(
            "⬇️ 下载 CSV",
            df.to_csv(index=False, encoding="utf-8-sig"),
            file_name="ci_newsku.csv", mime="text/csv"
        )


if __name__ == "__main__":
    main()
