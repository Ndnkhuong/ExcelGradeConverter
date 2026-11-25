"""Ô xem trước và kết quả xuất file."""

from __future__ import annotations

from dataclasses import dataclass, field

from app.models.validation_result import ValidationIssue


ACTION_UPDATE = "update"
ACTION_SKIP = "skip"


@dataclass
class PreviewCell:
    dest_row: int
    dest_column: int
    dest_header: str
    student_name: str
    source_value: object
    dest_value: object
    action: str
    reason: str


@dataclass
class ReviewResult:
    issues: list[ValidationIssue] = field(default_factory=list)
    cells: list[PreviewCell] = field(default_factory=list)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "error"]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "warning"]

    def can_export(self, acknowledge_warnings: bool) -> tuple[bool, str]:
        if self.errors:
            return False, "Còn lỗi nghiêm trọng. Hãy xử lý hoặc bỏ cột liên quan trước khi xuất."
        if self.warnings and not acknowledge_warnings:
            return False, "Hãy xác nhận các cảnh báo trước khi xuất file."
        updates = [cell for cell in self.cells if cell.action == ACTION_UPDATE]
        if not updates:
            return False, "Không có ô điểm hợp lệ nào để ghi vào file kết quả."
        return True, ""


@dataclass
class ConversionResult:
    output_path: str
    matched_count: int
    unmatched_count: int
    updated_cells: int
    skipped_cells: int
    mapping_name: str = ""
