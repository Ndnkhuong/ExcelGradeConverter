"""Xem các ô sẽ ghi, các ô bỏ qua và cảnh báo trước khi xuất."""

from __future__ import annotations

from PySide6.QtGui import QColor
from PySide6.QtWidgets import QCheckBox, QLabel, QMessageBox, QTabWidget, QVBoxLayout

from app.models.conversion_result import ACTION_UPDATE
from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block


class PreviewPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Xem trước kết quả",
                "Thứ tự học sinh của file đích được giữ nguyên. "
                "Chỉ những ô trong danh sách “Sẽ cập nhật” mới được ghi vào file mới.",
            )
        )
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.auto_convert_checkbox = QCheckBox("Tự động chuyển mã đánh giá")
        self.auto_convert_checkbox.setChecked(True)
        self.auto_convert_checkbox.setToolTip(
            "Khi bật: T → Hoàn thành tốt, H → Hoàn thành, C → Chưa hoàn thành (chỉ áp dụng cho cột đánh giá bằng chữ)."
        )
        self.auto_convert_checkbox.toggled.connect(self._toggle_auto_convert)
        layout.addWidget(self.auto_convert_checkbox)
        self.tabs = QTabWidget()
        self.updates = readonly_table(["Dòng đích", "Học sinh", "Cột", "Giá trị sẽ ghi", "Giá trị đích hiện tại", "Ghi chú"])
        self.skipped = readonly_table(["Dòng đích", "Học sinh", "Cột", "Giá trị nguồn", "Lý do bỏ qua"])
        self.warnings = readonly_table(["Mức", "Nội dung", "Vị trí"])
        self.tabs.addTab(self.updates, "Sẽ cập nhật")
        self.tabs.addTab(self.skipped, "Bỏ qua")
        self.tabs.addTab(self.warnings, "Cảnh báo")
        layout.addWidget(self.tabs, 1)
        self.confirm = QCheckBox("Tôi đã xem các cảnh báo và đồng ý bỏ qua những ô không được ghi")
        layout.addWidget(self.confirm)

    def on_enter(self) -> None:
        self.auto_convert_checkbox.blockSignals(True)
        self.auto_convert_checkbox.setChecked(self.ctx.session.auto_convert_assessment)
        self.auto_convert_checkbox.blockSignals(False)
        self.ctx.engine.review(self.ctx.session)
        self.confirm.setChecked(self.ctx.session.acknowledge_warnings)
        self._fill()

    def _toggle_auto_convert(self, checked: bool) -> None:
        self.ctx.session.auto_convert_assessment = bool(checked)
        self.ctx.session.acknowledge_warnings = False
        self.ctx.engine.review(self.ctx.session)
        self.confirm.setChecked(False)
        self._fill()

    def on_next(self) -> bool:
        review = self.ctx.session.review
        if review is None:
            return False
        allowed, message = review.can_export(self.confirm.isChecked())
        if not allowed:
            QMessageBox.warning(self, "Chưa thể xuất", message)
            return False
        self.ctx.session.acknowledge_warnings = self.confirm.isChecked()
        return True

    def _fill(self) -> None:
        review = self.ctx.session.review
        cells = [] if review is None else review.cells
        updates = [cell for cell in cells if cell.action == ACTION_UPDATE]
        skipped = [cell for cell in cells if cell.action != ACTION_UPDATE]
        warnings = [] if review is None else review.warnings
        self.summary.setText(
            f"Sẽ cập nhật {len(updates)} ô. Bỏ qua {len(skipped)} ô. Cảnh báo: {len(warnings)}."
        )
        self.confirm.setVisible(bool(warnings))
        self.updates.setRowCount(len(updates))
        for row, cell in enumerate(updates):
            values = [
                cell.dest_row,
                cell.student_name,
                cell.dest_header,
                _show(cell.source_value),
                _show(cell.dest_value),
                cell.reason,
            ]
            for column, value in enumerate(values):
                item = text_item(value)
                item.setBackground(QColor("#eefbf3"))
                self.updates.setItem(row, column, item)
        self.skipped.setRowCount(len(skipped))
        for row, cell in enumerate(skipped):
            values = [cell.dest_row, cell.student_name, cell.dest_header, _show(cell.source_value), cell.reason]
            for column, value in enumerate(values):
                self.skipped.setItem(row, column, text_item(value))
        self.warnings.setRowCount(len(warnings))
        for row, issue in enumerate(warnings):
            values = ["Cảnh báo", issue.message, issue.location or "—"]
            for column, value in enumerate(values):
                item = text_item(value)
                item.setBackground(QColor("#fff6e5"))
                self.warnings.setItem(row, column, item)


def _show(value: object) -> str:
    if value is None or value == "":
        return ""
    return str(value)
