"""
模块4：将本地下载的 PDF 上传到 SharePoint
- 复用 extend/python/sharepoint.py 中已有的 upload_sharepoint_file 函数
- 按品牌名建子文件夹
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "extend" / "python"))

from sharepoint import upload_sharepoint_file
from config import DOWNLOAD_BASE, SP_SITE_URL, SP_UPLOAD_FOLDER

log = logging.getLogger(__name__)


def run():
    """扫描本地下载目录，将所有 PDF 上传到 SharePoint"""
    uploaded = failed = skipped = 0

    for brand_dir in sorted(DOWNLOAD_BASE.iterdir()):
        if not brand_dir.is_dir():
            continue
        brand = brand_dir.name

        pdf_files = list(brand_dir.glob("*.pdf"))
        if not pdf_files:
            continue

        log.info(f"\n上传品牌: {brand}（{len(pdf_files)} 个文件）")

        for pdf_file in pdf_files:
            sp_path = f"{SP_UPLOAD_FOLDER}/{brand}/{pdf_file.name}"
            log.info(f"  → {sp_path}")

            try:
                ok = upload_sharepoint_file(
                    site_url=SP_SITE_URL,
                    local_path=str(pdf_file),
                    sharepoint_path=sp_path,
                )
                if ok:
                    uploaded += 1
                    log.info(f"  ✅ 上传成功")
                else:
                    failed += 1
                    log.warning(f"  ❌ 上传返回 False")
            except Exception as exc:
                failed += 1
                log.error(f"  ❌ 上传异常: {exc}")

    log.info(f"\n✅ 模块4 完成 | 成功: {uploaded}  失败: {failed}  跳过: {skipped}")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    run()
