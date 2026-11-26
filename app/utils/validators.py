"""Kiểm tra và chuyển đổi kiểu dữ liệu điểm. Không làm tròn và không đổi ô trống thành 0."""

from __future__ import annotations

import math
import re

_NUMERIC_SCORE_RE = re.compile(r"^[+-]?\d+(?:[.,]\d+)?$")

ASSESSMENT_CODE_MAP: dict[str, str] = {
    "T": "Hoàn thành tốt",
    "H": "Hoàn thành",
    "C": "Chưa hoàn thành",
}


def convert_assessment_code(value: object, enabled: bool) -> object:
    """Chuyển mã đánh giá đúng 1 ký tự T, H, C sang mô tả đầy đủ khi enabled=True.

    Không chuyển ô trống, None, giá trị số, bool hoặc các chuỗi nhiều ký tự khác.
    """
    if not enabled:
        return value
    if not isinstance(value, str):
        return value
    cleaned = value.replace("\u00a0", " ").strip().upper()
    if cleaned in ASSESSMENT_CODE_MAP:
        return ASSESSMENT_CODE_MAP[cleaned]
    return value


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def is_number(value: object) -> bool:
    """Số thật của Excel (int/float hữu hạn, không phải bool)."""
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return True
    if isinstance(value, float):
        return math.isfinite(value)
    return False


def parse_numeric_value(value: object) -> int | float | None:
    """Chuyển chuỗi số (hỗ trợ '.' hoặc ',') thành int/float.

    Trả về None nếu là ô trống hoặc dữ liệu không phải số.
    Không làm tròn, không chuyển bool hay ô trống thành số.
    """
    if is_blank(value) or isinstance(value, bool):
        return None
    if is_number(value):
        return value  # type: ignore[return-value]
    if not isinstance(value, str):
        return None
    cleaned = value.replace("\u00a0", " ").strip()
    if not _NUMERIC_SCORE_RE.match(cleaned):
        return None
    if "." in cleaned or "," in cleaned:
        parsed = float(cleaned.replace(",", "."))
        return parsed if math.isfinite(parsed) else None
    return int(cleaned)


def coerce_cell_value(value: object) -> object:
    """Nếu là chuỗi số thì trả về int/float, ngược lại giữ nguyên giá trị gốc."""
    parsed = parse_numeric_value(value)
    if parsed is not None:
        return parsed
    return value


def score_out_of_range(value: float, score_min: float, score_max: float) -> bool:
    return value < score_min or value > score_max
