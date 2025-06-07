"""
配置文件模块
包含程序的所有配置常量和设置
"""

import os
from pathlib import Path
from typing import Dict, Any

# 程序基本信息
APP_NAME = "收入证据管理程序"
APP_VERSION = "1.0.0"

# 文件路径配置
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"

# 数据文件路径
SETTINGS_FILE = DATA_DIR / "settings.pkl"
DATABASE_FILE = DATA_DIR / "database.pkl"
BACKUP_DIR = DATA_DIR / "backups"

# 确保数据目录存在
DATA_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)

# GUI配置
WINDOW_CONFIG = {
    "title": APP_NAME,
    "geometry": "1200x800",
    "min_width": 1000,
    "min_height": 600,
    "icon": None  # 可以后续添加图标文件路径
}

# 主题配置
THEME_CONFIG = {
    "appearance_mode": "light",  # "system", "dark", "light"
    "default_color_theme": "blue",  # "blue", "green", "dark-blue"
    "scaling": 1.0
}

# 表格列配置
TABLE_COLUMNS = {
    "合同号": {"width": 120, "required": True},
    "客户名": {"width": 150, "required": True},
    "本年确认的收入": {"width": 120, "required": True, "type": "number"},
    "附件确认的收入": {"width": 120, "required": False, "type": "number"},
    "差异": {"width": 100, "required": False, "type": "number"},
    "附件数量": {"width": 80, "required": False, "type": "number"},
    "差异备注": {"width": 200, "required": False},
    "导入时间": {"width": 120, "required": False, "type": "datetime"},
    "变化标识": {"width": 80, "required": False}
}

# 支持的文件格式
SUPPORTED_EXCEL_FORMATS = [".xlsx", ".xls"]
SUPPORTED_ATTACHMENT_FORMATS = [
    ".pdf", ".png", ".jpg", ".jpeg", ".gif", ".bmp", 
    ".zip", ".rar", ".7z", ".doc", ".docx", ".txt",
    ".eml", ".msg", ".tiff", ".tif"
]

# 导入配置
IMPORT_CONFIG = {
    "max_rows": 10000,  # 最大导入行数
    "required_columns": ["合同号", "客户名", "本年确认的收入"],
    "unique_column": "合同号"
}

# 备份配置
BACKUP_CONFIG = {
    "auto_backup": True,
    "backup_interval_days": 7,
    "max_backup_files": 10
}

# 默认设置
DEFAULT_SETTINGS: Dict[str, Any] = {
    "attachment_root_path": str(Path.home() / "Documents" / "收入证据管理" / "附件"),
    "theme": THEME_CONFIG,
    "window": WINDOW_CONFIG,
    "import": IMPORT_CONFIG,
    "backup": BACKUP_CONFIG,
    "last_import_path": "",
    "last_export_path": "",
    "first_run": True
}

# 日志配置
LOGGING_CONFIG = {
    "level": "INFO",
    "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    "file": DATA_DIR / "app.log",
    "max_size": 10 * 1024 * 1024,  # 10MB
    "backup_count": 5
} 