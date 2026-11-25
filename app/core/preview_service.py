"""Dựng bảng xem trước. Chỉ các ô hợp lệ mới được đánh dấu cập nhật."""

from __future__ import annotations

from app.models.column_mapping import ColumnMappingItem
from app.models.conversion_result import ACTION_SKIP, ACTION_UPDATE, PreviewCell, ReviewResult
from app.models.excel_file import SheetData
from app.models.student_record import MatchResult
from app.models.validation_result import SEVERITY_WARNING, ValidationIssue
from app.core.data_validator import classify_score, validate_structure
from app.utils.validators import is_blank, is_number


def build_review(
    source: SheetData,
    dest: SheetData,
    source_name_key: str,
    dest_name_key: str,
    mappings: list[ColumnMappingItem],
    match: MatchResult,
    score_min: float,
    score_max: float,
) -> ReviewResult:
    issues = validate_structure(source_name_key, dest_name_key, mappings, match)
    cells: list[PreviewCell] = []
    for pair in match.matched:
        for item in mappings:
            source_cell = pair.source.scores.get(item.source_key)
            dest_cell = pair.dest.scores.get(item.dest_key)
            action, value, issue = classify_score(
                source_cell,
                score_min,
                score_max,
                pair.source.name_raw,
                item.dest_header,
                source.sheet_name,
            )
            if issue is not None:
                issues.append(issue)
            if dest_cell is not None and dest_cell.is_formula:
                issues.append(
                    ValidationIssue(
                        SEVERITY_WARNING,
                        "DEST_FORMULA",
                        (
                            f"Ô “{item.dest_header}” của {pair.dest.name_raw} trên file đích đang là công thức. "
                            "Ứng dụng giữ nguyên công thức và không ghi đè."
                        ),
                        sheet=dest.sheet_name,
                        row=dest_cell.row,
                        column=item.dest_header,
                    )
                )
                action = ACTION_SKIP
                reason = "Giữ công thức của file đích"
            elif action == ACTION_UPDATE:
                reason = _overwrite_reason(dest_cell, value, pair.dest.name_raw, item.dest_header, dest.sheet_name, issues)
            else:
                reason = issue.message if issue is not None else "Bỏ qua"
            cells.append(
                PreviewCell(
                    dest_row=pair.dest.row_index,
                    dest_column=int(item.dest_key),
                    dest_header=item.dest_header,
                    student_name=pair.dest.name_raw,
                    source_value=None if source_cell is None else source_cell.value,
                    dest_value=None if dest_cell is None else dest_cell.value,
                    action=action if action == ACTION_UPDATE else ACTION_SKIP,
                    reason=reason,
                )
            )
    cells.sort(key=lambda cell: (cell.dest_row, cell.dest_column))
    return ReviewResult(issues=issues, cells=cells)


def _overwrite_reason(dest_cell, new_value, student_name, header, sheet, issues) -> str:
    if dest_cell is None or is_blank(dest_cell.value):
        return "Ghi vào ô trống"
    if is_number(dest_cell.value) and is_number(new_value) and float(dest_cell.value) == float(new_value):
        return "Giá trị trùng với ô đích"
    if isinstance(dest_cell.value, str) and isinstance(new_value, str) and dest_cell.value.strip() == new_value.strip():
        return "Giá trị trùng với ô đích"
    issues.append(
        ValidationIssue(
            SEVERITY_WARNING,
            "OVERWRITE",
            (
                f"Ô “{header}” của {student_name} trên file đích đang có dữ liệu "
                f"và sẽ được thay bằng dữ liệu nguồn nếu bạn xác nhận."
            ),
            sheet=sheet,
            row=dest_cell.row,
            column=header,
        )
    )
    return "Sẽ ghi đè dữ liệu đang có trên file đích"
