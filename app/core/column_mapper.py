"""Gợi ý mapping cột theo tên tiêu đề. Không tự ghép nếu không khớp tên."""

from __future__ import annotations

from dataclasses import dataclass

from app.models.column_mapping import COLUMN_TYPE_ASSESSMENT, COLUMN_TYPE_NUMBER
from app.models.excel_file import ColumnInfo, SheetData
from app.utils.name_normalizer import normalize_header
from app.utils.validators import ASSESSMENT_CODE_MAP, is_blank, parse_numeric_value
from config import DEFAULT_DEST_SCORE_COLUMN, DEFAULT_SOURCE_SCORE_COLUMN

_ASSESSMENT_HEADER_KEYWORDS = (
    "mức đạt",
    "muc dat",
    "đánh giá",
    "danh gia",
    "xếp loại",
    "xep loai",
    "nhận xét",
    "nhan xet",
    "mức độ",
    "muc do",
    "năng lực",
    "nang luc",
    "phẩm chất",
    "pham chat",
)

ALIAS_GROUPS = [
    {"giữa kỳ", "giữa kì", "giua ky", "giua ki", "gk", "điểm giữa kỳ", "điểm giữa kì", "diem giua ky"},
    {"cuối kỳ", "cuối kì", "cuoi ky", "cuoi ki", "ck", "điểm cuối kỳ", "điểm cuối kì", "diem cuoi ky"},
    {"tx1", "tx 1", "thường xuyên 1", "thuong xuyen 1", "điểm tx1", "diem tx1"},
    {"tx2", "tx 2", "thường xuyên 2", "thuong xuyen 2", "điểm tx2", "diem tx2"},
    {"tx3", "tx 3", "thường xuyên 3", "thuong xuyen 3", "điểm tx3", "diem tx3"},
    {"tx4", "tx 4", "thường xuyên 4", "thuong xuyen 4", "điểm tx4", "diem tx4"},
    {"mười lăm phút", "15 phút", "điểm 15 phút", "diem 15 phut"},
    {"một tiết", "1 tiết", "điểm 1 tiết", "diem 1 tiet"},
]

_ALIAS_INDEX: dict[str, int] = {}
for _group_id, _group in enumerate(ALIAS_GROUPS):
    for _name in _group:
        _ALIAS_INDEX[normalize_header(_name)] = _group_id


@dataclass
class MappingSuggestion:
    source_name_key: str | None
    dest_name_key: str | None
    pairs: list[tuple[str, str]]


def headers_match(left: str, right: str) -> bool:
    normalized_left = normalize_header(left)
    normalized_right = normalize_header(right)
    if not normalized_left or not normalized_right:
        return False
    if normalized_left == normalized_right:
        return True
    left_group = _ALIAS_INDEX.get(normalized_left)
    right_group = _ALIAS_INDEX.get(normalized_right)
    return left_group is not None and left_group == right_group


def find_column_by_letter(columns: list[ColumnInfo], letter: str) -> ColumnInfo | None:
    target = letter.strip().upper()
    for column in columns:
        if column.letter.upper() == target:
            return column
    return None


def suggest_default_mapping(
    source: SheetData,
    dest: SheetData,
    source_name_key: str | None = None,
    dest_name_key: str | None = None,
    source_score_letter: str = DEFAULT_SOURCE_SCORE_COLUMN,
    dest_score_letter: str = DEFAULT_DEST_SCORE_COLUMN,
) -> MappingSuggestion:
    """Mapping mặc định: chỉ ghép cột điểm E (mẫu A) sang cột điểm E (mẫu B).

    Không tự ghép các cột khác để đảm bảo không ghi đè ngoài phạm vi cột E.
    """
    from app.core.template_detector import detect_name_column

    if source_name_key is None:
        source_name_key = detect_name_column(source.columns)
    if dest_name_key is None:
        dest_name_key = detect_name_column(dest.columns)

    pairs: list[tuple[str, str]] = []
    source_score = find_column_by_letter(source.columns, source_score_letter)
    dest_score = find_column_by_letter(dest.columns, dest_score_letter)
    if (
        source_score is not None
        and dest_score is not None
        and source_score.key != source_name_key
        and dest_score.key != dest_name_key
    ):
        pairs.append((source_score.key, dest_score.key))
    return MappingSuggestion(source_name_key, dest_name_key, pairs)


def suggest_mapping(
    source: SheetData,
    dest: SheetData,
    source_name_key: str | None = None,
    dest_name_key: str | None = None,
) -> MappingSuggestion:
    """Áp dụng mapping mặc định cột E nguồn -> cột E đích."""
    return suggest_default_mapping(source, dest, source_name_key, dest_name_key)


def suggest_mapping_by_header(
    source: SheetData,
    dest: SheetData,
    source_name_key: str | None = None,
    dest_name_key: str | None = None,
) -> MappingSuggestion:
    """Gợi ý ghép cột theo tên tiêu đề trùng hoặc tương đồng."""
    from app.core.template_detector import detect_name_column

    if source_name_key is None:
        source_name_key = detect_name_column(source.columns)
    if dest_name_key is None:
        dest_name_key = detect_name_column(dest.columns)

    used_sources: set[str] = set()
    pairs: list[tuple[str, str]] = []
    for dest_column in dest.columns:
        if dest_column.key == dest_name_key:
            continue
        match = _find_source(dest_column, source.columns, source_name_key, used_sources)
        if match is None:
            continue
        used_sources.add(match.key)
        pairs.append((match.key, dest_column.key))
    return MappingSuggestion(source_name_key, dest_name_key, pairs)


def _find_source(
    dest_column: ColumnInfo,
    source_columns: list[ColumnInfo],
    source_name_key: str | None,
    used: set[str],
) -> ColumnInfo | None:
    for source_column in source_columns:
        if source_column.key in used or source_column.key == source_name_key:
            continue
        if headers_match(source_column.header, dest_column.header):
            return source_column
    return None


def find_column_by_header(columns: list[ColumnInfo], header: str) -> ColumnInfo | None:
    matches = [column for column in columns if headers_match(column.header, header)]
    if not matches:
        return None
    return matches[0]


def detect_column_type(
    source: SheetData,
    source_key: str,
    source_header: str = "",
    dest_header: str = "",
) -> str:
    """Tự nhận diện cột là điểm số (number) hay đánh giá bằng chữ (assessment)."""
    numeric_count = 0
    assessment_code_count = 0
    other_text_count = 0
    sampled = 0
    for row in source.rows:
        cell = row.get(source_key)
        if cell is None or is_blank(cell.value):
            continue
        sampled += 1
        if parse_numeric_value(cell.value) is not None:
            numeric_count += 1
        elif isinstance(cell.value, str):
            code = cell.value.replace("\u00a0", " ").strip().upper()
            if code in ASSESSMENT_CODE_MAP:
                assessment_code_count += 1
            else:
                other_text_count += 1
        if sampled >= 50:
            break

    if assessment_code_count > 0 and numeric_count == 0:
        return COLUMN_TYPE_ASSESSMENT
    if numeric_count > 0 and assessment_code_count == 0 and other_text_count == 0:
        return COLUMN_TYPE_NUMBER
    if other_text_count > 0 and numeric_count == 0:
        return COLUMN_TYPE_ASSESSMENT

    combined = f"{normalize_header(source_header)} {normalize_header(dest_header)}"
    if any(keyword in combined for keyword in _ASSESSMENT_HEADER_KEYWORDS):
        return COLUMN_TYPE_ASSESSMENT
    if assessment_code_count + other_text_count > numeric_count:
        return COLUMN_TYPE_ASSESSMENT
    return COLUMN_TYPE_NUMBER
