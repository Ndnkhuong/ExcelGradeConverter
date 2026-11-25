"""Thao tác file an toàn: không ghi đè file gốc, dọn file tạm."""

from __future__ import annotations

import os
import re
from pathlib import Path

from app.core.errors import AppError
from config import SUPPORTED_EXTENSIONS, temp_dir

_INVALID_FILENAME = re.compile(r'[<>:"/\\|?*]')


def ensure_supported_excel(path: Path) -> None:
    if not path.exists() or not path.is_file():
        raise AppError("Không tìm thấy file đã chọn.")
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise AppError(
            "Phiên bản này chưa hỗ trợ file .xls. Hãy mở file bằng Excel và lưu lại dưới dạng .xlsx."
        )
    if suffix not in SUPPORTED_EXTENSIONS:
        raise AppError("Chỉ hỗ trợ file Excel .xlsx và .xlsm.")
    if not os.access(path, os.R_OK):
        raise AppError("Không đọc được file. File có thể đang bị khóa hoặc bạn không có quyền đọc.")


def same_file(left: str | Path, right: str | Path) -> bool:
    try:
        return Path(left).resolve() == Path(right).resolve()
    except OSError:
        return False


def safe_filename(name: str, default_suffix: str = ".xlsx") -> str:
    cleaned = _INVALID_FILENAME.sub("_", name).strip().rstrip(".")
    if not cleaned:
        cleaned = "BangDiem_DaChuyenDoi" + default_suffix
    suffix = Path(cleaned).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        cleaned += default_suffix
    return cleaned


def assert_output_is_new(output: Path, source: Path, dest: Path) -> None:
    if same_file(output, source) or same_file(output, dest):
        raise AppError(
            "Không ghi đè file nguồn hoặc file đích gốc. Hãy chọn một tên file kết quả mới."
        )


def cleanup_temp() -> None:
    folder = temp_dir()
    for item in folder.glob("*"):
        if item.is_file():
            try:
                item.unlink()
            except OSError:
                continue


def open_in_explorer(path: Path) -> None:
    target = path if path.is_dir() else path.parent
    if not target.exists():
        raise AppError("Không tìm thấy thư mục cần mở.")
    os.startfile(target)  # type: ignore[attr-defined]
