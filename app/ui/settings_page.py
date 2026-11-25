"""Cài đặt thang điểm và thư mục dữ liệu local."""

from __future__ import annotations

from PySide6.QtWidgets import QFormLayout, QGroupBox, QLabel, QMessageBox, QPushButton, QDoubleSpinBox, QVBoxLayout

from app.core.errors import AppError
from app.ui.context import AppContext
from app.ui.widgets import Page, title_block
from app.utils.file_utils import open_in_explorer
from config import APP_VERSION, app_data_dir


class SettingsPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Cài đặt",
                "Thang điểm dùng để phát hiện điểm nằm ngoài phạm vi. Ứng dụng không tự đổi thang và không tự làm tròn.",
            )
        )
        box = QGroupBox("Thang điểm hợp lệ")
        form = QFormLayout(box)
        self.minimum = QDoubleSpinBox()
        self.maximum = QDoubleSpinBox()
        for spin in (self.minimum, self.maximum):
            spin.setRange(-1000, 1000)
            spin.setDecimals(2)
            spin.setSingleStep(0.5)
        form.addRow("Điểm nhỏ nhất", self.minimum)
        form.addRow("Điểm lớn nhất", self.maximum)
        layout.addWidget(box)

        save = QPushButton("Lưu cài đặt")
        save.setObjectName("primary")
        save.clicked.connect(self._save)
        layout.addWidget(save)

        info = QLabel(
            f"Phiên bản {APP_VERSION}. Dữ liệu ứng dụng nằm tại:\n{app_data_dir()}\n"
            "Nhật ký không ghi điểm hoặc họ tên học sinh."
        )
        info.setWordWrap(True)
        info.setObjectName("hint")
        layout.addWidget(info)
        open_data = QPushButton("Mở thư mục dữ liệu")
        open_data.clicked.connect(self._open_data)
        layout.addWidget(open_data)
        layout.addStretch(1)

    def on_enter(self) -> None:
        self.minimum.setValue(self.ctx.settings.score_min)
        self.maximum.setValue(self.ctx.settings.score_max)

    def _save(self) -> None:
        low = float(self.minimum.value())
        high = float(self.maximum.value())
        if low >= high:
            QMessageBox.warning(self, "Thang điểm chưa đúng", "Điểm nhỏ nhất phải nhỏ hơn điểm lớn nhất.")
            return
        self.ctx.settings.score_min = low
        self.ctx.settings.score_max = high
        self.ctx.settings.save()
        QMessageBox.information(self, "Đã lưu", f"Thang điểm hiện là {low:g}–{high:g}.")

    def _open_data(self) -> None:
        try:
            open_in_explorer(app_data_dir())
        except AppError as exc:
            QMessageBox.warning(self, "Không mở được", exc.message)
