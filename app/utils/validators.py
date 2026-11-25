"""Kiểm tra giá trị điểm. Không sửa và không làm tròn."""

from __future__ import annotations


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def is_number(value: object) -> bool:
    """Số thật của Excel. Không coi chữ số hay giá trị đúng/sai là điểm."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def score_out_of_range(value: float, score_min: float, score_max: float) -> bool:
    return value < score_min or value > score_max
