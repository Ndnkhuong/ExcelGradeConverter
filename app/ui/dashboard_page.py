"""Màn hình chính: bắt đầu chuyển đổi, lịch sử gần đây và hướng dẫn."""

from __future__ import annotations

from datetime import datetime

from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block


class DashboardPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)
        layout.addLayout(
            title_block(
                "Excel Grade Converter",
                "Chuyển điểm từ file Excel nguồn sang file mẫu đích trên máy này. "
                "File gốc không bị sửa và dữ liệu không được gửi lên mạng.",
            )
        )

        buttons = QHBoxLayout()
        self.start_button = QPushButton("Bắt đầu chuyển đổi")
        self.start_button.setObjectName("primary")
        self.profile_button = QPushButton("Mở cấu hình đã lưu")
        self.history_button = QPushButton("Lịch sử")
        self.settings_button = QPushButton("Cài đặt")
        for button in (
            self.start_button,
            self.profile_button,
            self.history_button,
            self.settings_button,
        ):
            buttons.addWidget(button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        layout.addWidget(QLabel("Các lần chuyển đổi gần đây"))
        self.table = readonly_table(
            ["Thời gian", "File nguồn", "File đích", "Cấu hình", "Đã khớp", "Chưa khớp"]
        )
        layout.addWidget(self.table, 1)

        guide = QLabel(
            "Quy trình: chọn hai file Excel → ghép cột điểm → đối chiếu họ tên → "
            "kiểm tra → xem trước → xuất file mới.\n"
            "Học sinh chỉ được ghép khi họ tên trùng nhau và không bị trùng trong từng file. "
            "Tên trùng phải được bạn xác nhận thủ công. Ứng dụng không tự tính lại điểm."
        )
        guide.setObjectName("hint")
        guide.setWordWrap(True)
        layout.addWidget(guide)

    def on_enter(self) -> None:
        entries = self.ctx.history_repo.list_recent(8)
        self.table.setRowCount(len(entries))
        for row, entry in enumerate(entries):
            when = _format_time(entry.converted_at)
            values = [
                when,
                entry.source_name,
                entry.dest_name,
                entry.mapping_name or "—",
                entry.matched_count,
                entry.unmatched_count,
            ]
            for column, value in enumerate(values):
                self.table.setItem(row, column, text_item(value))


def _format_time(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value
