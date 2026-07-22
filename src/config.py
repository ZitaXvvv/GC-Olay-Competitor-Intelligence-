"""
CI New SKU Cross Brand — 全局配置
所有需要修改的设置都集中在这里，其他文件不需要动。
"""

import os
from pathlib import Path

# ===== 品牌配置：英文名 → 中文搜索名 =====
BRANDS = {
    "PROYA": "珀莱雅",
    "GUYU": "谷雨",
    "OSM": "欧诗漫",
    "Kiehls": "科颜氏",
    "LOREAL": "欧莱雅",
    "ESTEE LAUDER": "雅诗兰黛",
    "SKIN CEUTICALS": "修丽可",
    "Lancome": "兰蔻",
    "BQL": "百雀羚",
    "Chando": "自然堂",
    "Winona": "薇诺娜",
    "Kans": "韩束",
    "Clains": "娇韵诗",
}

# 搜索类目（美丽修行大数据）
SEARCH_CATEGORIES = ["防晒", "护肤"]

# 查找时间窗口（天）；可用 --days N 命令行参数覆盖，或直接修改这里
TIME_PERIOD_DAYS = int(os.environ.get("CI_DAYS", "28"))

# ===== 文件路径（按需修改） =====
EXCEL_PATH = r"C:\Users\xie.x.3\Documents\Olay CI\CI_List_Ada.xlsx"
DOWNLOAD_BASE = Path(r"C:\Users\xie.x.3\Documents\Olay CI")
COOKIES_FILE = Path(__file__).parent / "bebd_cookies.json"
LOG_DIR = Path(__file__).parent.parent / "log"
CHECKPOINT_FILE = Path(__file__).parent / "checkpoint.json"

# ===== 目标网站 =====
BEBD_URL = "https://bebd.bevol.com/"
NMPA_DATASEARCH_URL = "https://www.nmpa.gov.cn/datasearch/"
HZPBA_SEARCH_URL = "https://hzpba.nmpa.gov.cn/HZPBZCX/PTHZPBA-WEBUI/#/newskincare"         # 国产普通化妆品备案
HZPBA_IMPORT_URL  = "https://hzpba.nmpa.gov.cn/HZPBZCX/PTHZPBA-WEBUI/#/importcosmetics"    # 进口普通化妆品备案
HZPBA_PDF_BASE = (
    "https://hzpba.nmpa.gov.cn/HZPBZCX/PTHZPBA-SERVER/nmpafile/gsxxFilePreview"
)

# ===== SharePoint 配置 =====
SP_SITE_URL = "https://pgone.sharepoint.com/sites/ChinaOlayinnovationCI"
SP_UPLOAD_FOLDER = "Shared Documents/CI Hero SKU pic"   # 【确认上传目标文件夹名称】
SP_CLIENT_ID = "ed61bbd2-1948-4bbe-9849-7efb34bf9181"
SP_CLIENT_SECRET = "itI8Q~Vv8N5EPhr6HTc0vXcq08cPIkFTd-nbpbX3"
SP_TENANT_ID = "3596192b-fdf5-4e2c-a6fa-acb706c963d8"

# ===== 邮件配置 =====
EMAIL_FROM = "xie.x.3@pg.com"
EMAIL_TO = ["xie.x.3@pg.com", "iolay.im@pg.com"]
EMAIL_SUBJECT = "<Info Sharing>CI of new notification & Registration"
SP_DETAIL_LINK = (
    "https://pgone.sharepoint.com/sites/ChinaOlayinnovationCI/Shared%20Documents"
    "/Forms/AllItems.aspx?viewid=2ddb2cae%2D6f74%2D48c4%2Db8ce%2D409df8974861"
    "&noAuthRedirect=1"
)

# ===== Excel 列号定义（1=A, 2=B ...）不要修改 =====
COL_UPLOAD_DATE = 1   # A: upload time（脚本运行当天）
COL_NAME        = 2   # B: Name（产品名称）
COL_EFFECT      = 3   # C: English / Benefit（功效宣称）
COL_DATE        = 4   # D: Notification Time（备案时间）
COL_REG_NUM     = 5   # E: #（备案/注册号）—— 去重依据，与旧Excel格式一致
COL_CATEGORY    = 6   # F: Registration Ti（类目/其他）
COL_INGREDIENTS = 7   # G: Ingredient（成分列表）
COL_PDF_URL     = 8   # H: link（备案/注册链接 or PDF预览URL）← 流程块2 value[7]
COL_LABEL_URL   = 9   # I: 化妆品产品标签链接（artwork PDF）
COL_POC_URL     = 10  # J: mini POC（功效证明链接）

# 向后兼容别名
COL_NORMAL_URL  = COL_PDF_URL
COL_SPECIAL_URL = COL_PDF_URL
COL_EXTRA       = COL_POC_URL
