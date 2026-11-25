"""Cấu hình ứng dụng và đường dẫn dữ liệu local."""

from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "ExcelGradeConverter"
APP_TITLE = "Excel Grade Converter"
APP_VERSION = "1.0.0"

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}

# Giới hạn an toàn khi quét file bảng điểm.
MAX_SCAN_COLUMNS = 120
MAX_DATA_ROWS = 20000
EMPTY_ROW_STREAK = 100
PREVIEW_ROWS = 20
PREVIEW_COLUMNS = 10

# Mapping điểm mặc định: mẫu A (nguồn) từ dòng 14 cột E -> mẫu B (đích) từ dòng 2 cột E.
DEFAULT_SOURCE_HEADER_ROW = 13
DEFAULT_SOURCE_DATA_START = 14
DEFAULT_SOURCE_SCORE_COLUMN = "E"

DEFAULT_DEST_HEADER_ROW = 1
DEFAULT_DEST_DATA_START = 2
DEFAULT_DEST_SCORE_COLUMN = "E"


def app_data_dir() -> Path:
    """Thư mục dữ liệu trong AppData, không lưu bảng điểm gốc."""
    appdata = os.environ.get("APPDATA")
    if appdata:
        path = Path(appdata) / APP_NAME
    else:
        path = Path.home() / f".{APP_NAME}"
    for folder in ("database", "logs", "temp"):
        (path / folder).mkdir(parents=True, exist_ok=True)
    return path


def database_path() -> Path:
    return app_data_dir() / "database" / "app.db"


def settings_path() -> Path:
    return app_data_dir() / "settings.json"


def log_path() -> Path:
    return app_data_dir() / "logs" / "app.log"


def temp_dir() -> Path:
    return app_data_dir() / "temp"


def project_root() -> Path:
    return Path(__file__).resolve().parent
