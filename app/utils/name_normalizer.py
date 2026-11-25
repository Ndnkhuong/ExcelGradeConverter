"""Chuẩn hóa họ tên và tiêu đề cột để so sánh."""

from __future__ import annotations

import re
import unicodedata

_WHITESPACE = re.compile(r"\s+")


def normalize_person_name(value: object) -> str:
    """Chuẩn hóa họ tên theo quy tắc bắt buộc.

    NFC, cắt khoảng trắng, gộp khoảng trắng, không phân biệt hoa thường.
    Giữ nguyên dấu tiếng Việt.
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFC", str(value))
    text = text.replace("\u00a0", " ").strip()
    text = _WHITESPACE.sub(" ", text)
    return text.casefold()


def display_name(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFC", str(value))
    text = text.replace("\u00a0", " ").strip()
    return _WHITESPACE.sub(" ", text)


def normalize_header(value: object) -> str:
    return normalize_person_name(value)
