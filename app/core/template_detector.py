"""Nhận diện dòng tiêu đề và cột họ tên."""

from __future__ import annotations

from app.models.excel_file import ColumnInfo
from app.utils.name_normalizer import normalize_header

NAME_HEADERS = {
    "họ và tên",
    "họ tên",
    "họ và tên học sinh",
    "họ tên học sinh",
    "tên học sinh",
    "ho va ten",
    "ho ten",
    "ho va ten hoc sinh",
    "ten hoc sinh",
    "full name",
    "student name",
    "name",
}


def detect_header_row(grid: list[list[object]], max_rows: int = 30) -> int:
    """Chọn dòng có nhiều ô chữ nhất trong các dòng đầu. Trả về số dòng 1-based."""
    best_row = 1
    best_score = -1
    limit = min(len(grid), max_rows)
    for offset in range(limit):
        row = grid[offset]
        text_cells = 0
        for value in row:
            if value is None or str(value).strip() == "":
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                continue
            text_cells += 1
        if text_cells > best_score:
            best_score = text_cells
            best_row = offset + 1
    return best_row


def is_name_header(header: str) -> bool:
    normalized = normalize_header(header)
    if normalized in NAME_HEADERS:
        return True
    return "họ" in normalized and "tên" in normalized


def detect_name_column(columns: list[ColumnInfo]) -> str | None:
    for column in columns:
        if is_name_header(column.header):
            return column.key
    return None
