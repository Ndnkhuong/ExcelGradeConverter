"""Phân loại ô điểm và các lỗi cấu trúc. Không sửa giá trị điểm."""

from __future__ import annotations

from app.models.column_mapping import COLUMN_TYPE_ASSESSMENT, COLUMN_TYPE_NUMBER, ColumnMappingItem
from app.models.excel_file import CellData
from app.models.student_record import MatchResult
from app.models.validation_result import (
    SEVERITY_ERROR,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    ValidationIssue,
)
from app.utils.validators import (
    convert_assessment_code,
    is_blank,
    parse_numeric_value,
    score_out_of_range,
)


def validate_structure(
    source_name_key: str,
    dest_name_key: str,
    mappings: list[ColumnMappingItem],
    match: MatchResult,
) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    if not source_name_key or not dest_name_key:
        issues.append(
            ValidationIssue(
                SEVERITY_ERROR,
                "REQUIRED_NAME",
                "Chưa chọn cột họ tên ở cả file nguồn và file đích.",
            )
        )
    if not mappings:
        issues.append(
            ValidationIssue(
                SEVERITY_ERROR,
                "REQUIRED_MAPPING",
                "Chưa có cột điểm nào được ghép.",
            )
        )
    dest_keys = [item.dest_key for item in mappings]
    if len(dest_keys) != len(set(dest_keys)):
        issues.append(
            ValidationIssue(
                SEVERITY_ERROR,
                "DUPLICATE_DEST",
                "Một cột đích đang được ghép nhiều hơn một lần.",
            )
        )
    source_rows = [pair.source.row_index for pair in match.matched]
    dest_rows = [pair.dest.row_index for pair in match.matched]
    if len(source_rows) != len(set(source_rows)) or len(dest_rows) != len(set(dest_rows)):
        issues.append(
            ValidationIssue(
                SEVERITY_ERROR,
                "DUPLICATE_MATCH",
                "Một học sinh đang được ghép nhiều lần.",
            )
        )
    if not match.matched:
        issues.append(
            ValidationIssue(
                SEVERITY_ERROR,
                "NO_MATCH",
                "Không có học sinh nào được ghép. Điểm sẽ không được ghi.",
            )
        )
    for student in match.unmatched_source:
        issues.append(
            ValidationIssue(
                SEVERITY_WARNING,
                "UNMATCHED_SOURCE",
                f"“{student.name_raw}” chưa có trong file đích. Điểm của học sinh này sẽ không được ghi.",
                row=student.row_index,
                column="Họ và tên",
            )
        )
    for group in match.ambiguous:
        issues.append(
            ValidationIssue(
                SEVERITY_WARNING,
                "AMBIGUOUS_NAME",
                (
                    f"Tên “{group.display_name}” xuất hiện {len(group.sources)} lần ở file nguồn "
                    f"và {len(group.dests)} lần ở file đích. Cần ghép thủ công. "
                    "Điểm của các học sinh này sẽ không được ghi cho đến khi được xác nhận."
                ),
            )
        )
    for student in match.dest_only:
        issues.append(
            ValidationIssue(
                SEVERITY_INFO,
                "DEST_ONLY",
                f"“{student.name_raw}” chỉ có trong file đích và được giữ nguyên.",
                row=student.row_index,
                column="Họ và tên",
            )
        )
    return issues


def classify_score(
    cell: CellData | None,
    score_min: float,
    score_max: float,
    student_name: str,
    header: str,
    sheet: str,
    column_type: str = COLUMN_TYPE_NUMBER,
    auto_convert_assessment: bool = True,
) -> tuple[str, object, ValidationIssue | None]:
    """Trả về (hành động, giá trị ghi, lỗi nếu có).

    Hành động là "update" hoặc "skip".
    - Giá trị số không bao giờ bị chuyển mã đánh giá.
    - Mã đánh giá T/H/C chỉ chuyển khi column_type == COLUMN_TYPE_ASSESSMENT và auto_convert_assessment=True.
    """
    if cell is None or is_blank(cell.value):
        if cell is not None and cell.unresolved_formula:
            return (
                "skip",
                None,
                ValidationIssue(
                    SEVERITY_WARNING,
                    "FORMULA_NO_VALUE",
                    f"Ô điểm “{header}” của {student_name} là công thức nhưng không đọc được giá trị đã tính. Ô này được bỏ qua.",
                    sheet=sheet,
                    row=cell.row,
                    column=header,
                ),
            )
        row = None if cell is None else cell.row
        return (
            "skip",
            None,
            ValidationIssue(
                SEVERITY_WARNING,
                "SCORE_MISSING",
                f"Học sinh {student_name} chưa có điểm ở cột “{header}”. Ô đích được giữ nguyên.",
                sheet=sheet,
                row=row,
                column=header,
            ),
        )
    numeric_value = parse_numeric_value(cell.value)
    if numeric_value is not None:
        if score_out_of_range(float(numeric_value), score_min, score_max):
            return (
                "skip",
                None,
                ValidationIssue(
                    SEVERITY_ERROR,
                    "SCORE_OUT_OF_RANGE",
                    (
                        f"Điểm “{header}” của {student_name} là {cell.value}, "
                        f"nằm ngoài thang {score_min:g}–{score_max:g}."
                    ),
                    sheet=sheet,
                    row=cell.row,
                    column=header,
                ),
            )
        return "update", numeric_value, None
    if column_type == COLUMN_TYPE_ASSESSMENT:
        return "update", convert_assessment_code(cell.value, auto_convert_assessment), None
    return "update", cell.value, None
