"""
CI New SKU Cross Brand — 主流程编排
替代原 UiBot 的 Python + Playwright 方案

用法：
  python src/main.py                   # 运行完整流程（步骤 1-5，默认查 28 天）
  python src/main.py --days 90         # 查过去 90 天的数据
  python src/main.py --step 1          # 只运行步骤1
  python src/main.py --from-step 2     # 从步骤 2 开始（用于手动中断后续跑）
  python src/main.py --resume          # 从上次中断处继续（自动读取 checkpoint）

步骤说明：

  步骤 1 — 美丽修行大数据搜索 + 特殊化妆品注册查询
    ① BEBD (美丽修行) 搜索各品牌新品 → 写入 Excel
    ② 对备案/注册号含“特”的产品：查 nmpa.gov.cn/datasearch
         - 含“特”且不含“进” → 国产特殊化妆品注册信息
         - 含“特”且含“进”    → 进口特殊化妆品注册信息

  步骤 2 — 普通化妆品备案 URL 查询回写
    ① 对备案号含“备”的产品：查 hzpba.nmpa.gov.cn
         - 含“备”且不含“进” → 国产普通化妆品备案（#/newskincare）
         - 含“备”且含“进”    → 进口普通化妆品备案（#/importcosmetics）
    ② 获取 PDF 预览 URL（gsxxFilePreview?attachmentId=xxx）写入 Excel H 列

  步骤 3 — 下载 PDF 到本地
    ① 读 Excel H 列 (link)，对每个 URL：
         - URL 含 hzpba.nmpa.gov.cn → 普通化妆品，保存为 {Name}.pdf
         - URL 含 nmpa.gov.cn/datasearch → 特殊化妆品，保存为 特化--{Name}.pdf
    ② 保存到本地文件夹 Documents/Olay CI/{brand_key}/

  步骤 4 — 上传 PDF 到 SharePoint

  步骤 5 — 发送汇总邮件

备案号关键字分类规则：
  含“特”  → 特殊化妆品注册 → nmpa.gov.cn/datasearch
  含“备”  → 普通化妆品备案 → hzpba.nmpa.gov.cn
  含“进”  → 进口（在以上两个网站内分别选进口页面）
"""

# ── 必须在所有其他 import 之前解析 --days，以便 config.py 读到正确的环境变量 ──
import os, sys
_days_default = 28
for _i, _arg in enumerate(sys.argv):
    if _arg in ("--days", "-d") and _i + 1 < len(sys.argv):
        try:
            _days_default = int(sys.argv[_i + 1])
        except ValueError:
            pass
os.environ["CI_DAYS"] = str(_days_default)
# ─────────────────────────────────────────────────────────────────────────────

import argparse
import importlib
import json
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import CHECKPOINT_FILE, LOG_DIR

# ─────────────────────────────────────────────
# 日志配置
# ─────────────────────────────────────────────
LOG_DIR.mkdir(parents=True, exist_ok=True)
_log_file = LOG_DIR / f"ci_bot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(_log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("main")

# ─────────────────────────────────────────────
# 步骤定义
# ─────────────────────────────────────────────
STEPS = {
    1: ("BEBD抓取 + 特殊化妆品注册查询 (nmpa.gov.cn/datasearch)", "module1_bebd"),
    2: ("普通化妆品备案 URL 查询 (hzpba.nmpa.gov.cn)",          "module2_nmpa"),
    3: ("PDF 下载到本地 (hzpba→普通 / datasearch→特殊)",      "module3_download"),
    4: ("上传 SharePoint",                                              "module4_upload"),
    5: ("发送汇总邮件",                                              "module5_email"),
}


# ─────────────────────────────────────────────
# Checkpoint（断点续跑）
# ─────────────────────────────────────────────

def save_checkpoint(step: int, status: str):
    data = {
        "last_completed_step": step,
        "status": status,
        "timestamp": datetime.now().isoformat(),
    }
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_checkpoint() -> int:
    if not CHECKPOINT_FILE.exists():
        return 0
    try:
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("last_completed_step", 0)
    except Exception:
        return 0


# ─────────────────────────────────────────────
# 运行单个步骤
# ─────────────────────────────────────────────

def run_step(step_num: int):
    desc, module_name = STEPS[step_num]
    log.info(f"\n{'='*60}")
    log.info(f"▶ 步骤 {step_num}/{len(STEPS)}: {desc}")
    log.info(f"{'='*60}")

    try:
        module = importlib.import_module(module_name)
        module.run()
        save_checkpoint(step_num, "completed")
        log.info(f"✅ 步骤 {step_num} 完成\n")
    except Exception:
        save_checkpoint(step_num, "failed")
        log.error(f"❌ 步骤 {step_num} 失败:\n{traceback.format_exc()}")
        raise


# ─────────────────────────────────────────────
# 主函数
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="CI New SKU Cross Brand 自动化流程")
    parser.add_argument("--days", "-d", type=int, default=_days_default, metavar="N",
                        help="查询过去 N 天的数据（默认 28）")
    parser.add_argument("--step",      type=int, metavar="N",
                        help="只运行步骤 N (1-5)")
    parser.add_argument("--from-step", type=int, metavar="N", dest="from_step",
                        help="从步骤 N 开始运行")
    parser.add_argument("--resume",    action="store_true",
                        help="从上次中断的步骤继续")
    args = parser.parse_args()
    # 确保 config 拿到最终值（argparse 可能覆盖了预解析值）
    os.environ["CI_DAYS"] = str(args.days)

    log.info(f"CI New SKU Cross Brand Bot 启动")
    log.info(f"日志文件: {_log_file}")
    log.info(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"查询时间窗口: 过去 {args.days} 天")
    log.info("")
    log.info("备案号分类规则：")
    log.info("  合 '特'  → 步骤 1 查 nmpa.gov.cn/datasearch  (特殊化妆品注册)")
    log.info("         进一步: 含'进' → 进口 | 不含'进' → 国产")
    log.info("  含 '备'  → 步骤 2 查 hzpba.nmpa.gov.cn      (普通化妆品备案)")
    log.info("         进一步: 含'进' → 进口 | 不含'进' → 国产")
    log.info("  两类 PDF 均写入 Excel H 列 (link)，步骤 3 按 URL 含有的域名判断并下载")
    log.info("")

    if args.step:
        # 只跑单步
        if args.step not in STEPS:
            log.error(f"无效步骤: {args.step}（有效范围 1-5）")
            sys.exit(1)
        run_step(args.step)

    else:
        # 确定起始步骤
        start = 1
        if args.resume:
            last = load_checkpoint()
            start = last + 1
            if start > len(STEPS):
                log.info("所有步骤已完成，无需重跑")
                return
            log.info(f"断点续跑 → 从步骤 {start} 开始（上次完成: {last}）")
        elif args.from_step:
            start = args.from_step

        for n in range(start, len(STEPS) + 1):
            try:
                run_step(n)
            except Exception:
                log.error(f"流程在步骤 {n} 中止")
                log.error(f"修复问题后运行 --from-step {n} 或 --resume 可重新继续")
                sys.exit(1)

    log.info(f"\n{'='*60}")
    log.info("🎉 全部步骤完成！")
    log.info(f"{'='*60}")


if __name__ == "__main__":
    main()
