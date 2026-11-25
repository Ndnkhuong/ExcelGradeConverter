"""Lịch sử chuyển đổi. Không hiển thị điểm học sinh."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QHBoxLayout, QMessageBox, QPushButton, QVBoxLayout

from app.core.errors import AppError
from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block
from app.utils.file_utils import open_in_explorer


class HistoryPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Lịch sử chuyển đổi",
                "Chỉ lưu thời gian, tên file và số học sinh. Nội dung điểm không được lưu trong lịch sử.",
            )
        )
        tools = QHBoxLayout()
        refresh = QPushButton("Làm mới")
        open_folder = QPushButton("Mở thư mục file kết quả")
        refresh.clicked.connect(self.on_enter)
        open_folder.clicked.connect(self._open)
        tools.addWidget(refresh)
        tools.addWidget(open_folder)
        tools.addStretch(1)
        layout.addLayout(tools)
        self.table = readonly_table(
            ["Thời gian", "File nguồn", "File đích", "Cấu hình", "Đã khớp", "Chưa khớp", "File kết quả"]
        )
        layout.addWidget(self.table, 1)

    def on_enter(self) -> None:
        entries = self.ctx.history_repo.list_recent(100)
        self.table.setRowCount(len(entries))
        for row, entry in enumerate(entries):
            values = [
                _format_time(entry.converted_at),
                entry.source_name,
                entry.dest_name,
                entry.mapping_name or "—",
                entry.matched_count,
                entry.unmatched_count,
                entry.result_path,
            ]
            for column, value in enumerate(values):
                self.table.setItem(row, column, text_item(value))

    def _open(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Chưa chọn dòng", "Hãy chọn một lần chuyển đổi.")
            return
        item = self.table.item(row, 6)
        if item is None or not item.text():
            return
        path = Path(item.text())
        try:
            open_in_explorer(path if path.exists() else path.parent)
        except AppError as exc:
            QMessageBox.warning(self, "Không mở được", exc.message)


def _format_time(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value
