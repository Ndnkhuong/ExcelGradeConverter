"""Kết quả kiểm tra dữ liệu trước khi xuất."""

from __future__ import annotations

from dataclasses import dataclass


SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"
SEVERITY_INFO = "info"


@dataclass
class ValidationIssue:
    severity: str
    code: str
    message: str
    sheet: str = ""
    row: int | None = None
    column: str = ""

    @property
    def location(self) -> str:
        parts = []
        if self.sheet:
            parts.append(self.sheet)
        if self.row:
            parts.append(f"dòng {self.row}")
        if self.column:
            parts.append(self.column)
        return ", ".join(parts)
