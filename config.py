# 配置文件：Excel 路径、数据库路径等
import os
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent

# 默认 Excel 文件路径（可改为你的实际路径，支持无扩展名或 .xlsx / .xls）
_DEFAULT_BASE = r"C:\Users\Administrator\Desktop\9.0网页计划\9.9911.xlsx"
for cand in (_DEFAULT_BASE, _DEFAULT_BASE + ".xlsx", _DEFAULT_BASE + ".xls"):
    if os.path.isfile(cand):
        _DEFAULT_EXCEL = cand
        break
else:
    _DEFAULT_EXCEL = _DEFAULT_BASE + ".xlsx"  # 默认假定 .xlsx
EXCEL_PATH = os.environ.get("EXCEL_PATH", _DEFAULT_EXCEL)

# SQLite 数据库路径（Render 上不设置则用项目内 excel_data.db）
DB_PATH = os.environ.get("DB_PATH", str(BASE_DIR / "excel_data.db"))

# 导入的图片保存目录
IMAGES_DIR = BASE_DIR / "static" / "images"
IMAGES_DIR.mkdir(parents=True, exist_ok=True)

# 每批写入数据库的行数（导入时）
IMPORT_BATCH_SIZE = 5000

# 部署用：只导入前 N 行，使 excel_data.db 控制在 100MB 内以便推送到 GitHub（不设则导入全部）
# 用法（仅本地生成“小库”时使用）：IMPORT_MAX_ROWS=5000 python import_excel.py
IMPORT_MAX_ROWS = int(os.environ.get("IMPORT_MAX_ROWS", "0") or "0")

# Render 部署用：若设置 DB_DOWNLOAD_URL，则后端在启动时会在本地不存在 excel_data.db 时，
# 自动从该 URL 下载数据库文件到 DB_PATH。适合将完整 excel_data.db 放到 GitHub Release 等位置。
DB_DOWNLOAD_URL = os.environ.get("DB_DOWNLOAD_URL", "").strip()
