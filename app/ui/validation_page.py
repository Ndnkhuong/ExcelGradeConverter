"""Hiển thị lỗi và cảnh báo theo dòng, cột."""

from __future__ import annotations

from PySide6.QtGui import QColor
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QMessageBox, QVBoxLayout

from app.models.validation_result import SEVERITY_ERROR, SEVERITY_INFO, SEVERITY_WARNING
from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block

_LABELS = {
    SEVERITY_ERROR: "Lỗi nghiêm trọng",
    SEVERITY_WARNING: "Cảnh báo",
    SEVERITY_INFO: "Thông tin",
}
_COLORS = {
    SEVERITY_ERROR: QColor("#fdecec"),
    SEVERITY_WARNING: QColor("#fff6e5"),
    SEVERITY_INFO: QColor("#ffffff"),
}


class ValidationPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Kiểm tra dữ liệu",
                "Lỗi nghiêm trọng cần được xử lý trước khi xuất. "
                "Bạn có thể quay lại để bỏ ghép học sinh hoặc bỏ cột điểm liên quan. "
                "Ứng dụng không tự sửa điểm.",
            )
        )
        row = QHBoxLayout()
        row.addWidget(QLabel("Lọc"))
        self.filter = QComboBox()
        self.filter.addItems(["Tất cả", "Lỗi nghiêm trọng", "Cảnh báo", "Thông tin"])
        self.filter.currentIndexChanged.connect(self._fill)
        row.addWidget(self.filter)
        row.addStretch(1)
        layout.addLayout(row)
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.table = readonly_table(["Mức", "Nội dung", "Vị trí"])
        layout.addWidget(self.table, 1)

    def on_enter(self) -> None:
        self.ctx.engine.review(self.ctx.session)
        self._fill()

    def on_next(self) -> bool:
        review = self.ctx.session.review
        if review is None or review.errors:
            QMessageBox.warning(
                self,
                "Còn lỗi nghiêm trọng",
                "Hãy xử lý các lỗi nghiêm trọng trước khi xem trước và xuất file.",
            )
            return False
        return True

    def _fill(self) -> None:
        review = self.ctx.session.review
        issues = [] if review is None else review.issues
        selected = self.filter.currentText()
        wanted = {
            "Lỗi nghiêm trọng": SEVERITY_ERROR,
            "Cảnh báo": SEVERITY_WARNING,
            "Thông tin": SEVERITY_INFO,
        }.get(selected)
        visible = [issue for issue in issues if wanted is None or issue.severity == wanted]
        error_count = sum(1 for issue in issues if issue.severity == SEVERITY_ERROR)
        warning_count = sum(1 for issue in issues if issue.severity == SEVERITY_WARNING)
        self.summary.setText(f"Lỗi nghiêm trọng: {error_count}. Cảnh báo: {warning_count}.")
        self.table.setRowCount(len(visible))
        for row, issue in enumerate(visible):
            values = [_LABELS.get(issue.severity, issue.severity), issue.message, issue.location or "—"]
            for column, value in enumerate(values):
                item = text_item(value)
                item.setBackground(_COLORS.get(issue.severity, QColor("#ffffff")))
                self.table.setItem(row, column, item)
            self.table.resizeRowToContents(row)
