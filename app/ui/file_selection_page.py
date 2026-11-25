"""Chọn file nguồn, file đích, sheet và dòng bắt đầu dữ liệu."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QComboBox,
    QVBoxLayout,
)

from app.core.errors import AppError
from app.ui.context import AppContext
from app.ui.widgets import Page, fill_preview, readonly_table, title_block
from app.utils.file_utils import same_file
from config import (
    DEFAULT_DEST_DATA_START,
    DEFAULT_DEST_HEADER_ROW,
    DEFAULT_SOURCE_DATA_START,
    DEFAULT_SOURCE_HEADER_ROW,
    PREVIEW_ROWS,
)


class FileSelectionPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        self._signature: tuple | None = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Chọn file Excel",
                "Chọn file nguồn (mẫu A, mặc định dữ liệu từ dòng 14) và file đích (mẫu B, mặc định dữ liệu từ dòng 2). "
                "Bạn vẫn có thể chỉnh lại dòng tiêu đề và dòng bắt đầu dữ liệu nếu cần.",
            )
        )

        columns = QHBoxLayout()
        self.source_path, self.source_sheet, self.source_header, self.source_start, source_box = (
            self._file_box(
                "File nguồn (Mẫu A)",
                self._browse_source,
                DEFAULT_SOURCE_HEADER_ROW,
                DEFAULT_SOURCE_DATA_START,
            )
        )
        self.dest_path, self.dest_sheet, self.dest_header, self.dest_start, dest_box = self._file_box(
            "File đích (Mẫu B)",
            self._browse_dest,
            DEFAULT_DEST_HEADER_ROW,
            DEFAULT_DEST_DATA_START,
        )
        columns.addWidget(source_box)
        columns.addWidget(dest_box)
        layout.addLayout(columns)

        self.notice = QLabel("")
        self.notice.setObjectName("hint")
        self.notice.setWordWrap(True)
        layout.addWidget(self.notice)

        layout.addWidget(QLabel("Xem nhanh file nguồn"))
        self.source_preview = readonly_table([])
        layout.addWidget(self.source_preview, 1)
        layout.addWidget(QLabel("Xem nhanh file đích"))
        self.dest_preview = readonly_table([])
        layout.addWidget(self.dest_preview, 1)

        for spin in (self.source_header, self.dest_header, self.source_start, self.dest_start):
            spin.valueChanged.connect(self._refresh_previews)
        self.source_sheet.currentIndexChanged.connect(lambda: self._sheet_changed("source"))
        self.dest_sheet.currentIndexChanged.connect(lambda: self._sheet_changed("dest"))

    def _file_box(self, title: str, browse, default_header: int, default_start: int):
        box = QGroupBox(title)
        form = QFormLayout(box)
        path = QLineEdit()
        path.setReadOnly(True)
        path_row = QHBoxLayout()
        path_row.addWidget(path, 1)
        button = QPushButton("Chọn file")
        button.clicked.connect(browse)
        path_row.addWidget(button)
        sheet = QComboBox()
        header = QSpinBox()
        header.setRange(1, 5000)
        header.setValue(default_header)
        start = QSpinBox()
        start.setRange(2, 5001)
        start.setValue(default_start)
        form.addRow("Đường dẫn", path_row)
        form.addRow("Sheet", sheet)
        form.addRow("Dòng tiêu đề", header)
        form.addRow("Dòng bắt đầu dữ liệu", start)
        return path, sheet, header, start, box

    def on_enter(self) -> None:
        session = self.ctx.session
        if self._signature is None and not session.source_path and not session.dest_path:
            source_header, source_start = self._default_rows("source")
            dest_header, dest_start = self._default_rows("dest")
            self._set_spin_values(self.source_header, self.source_start, source_header, source_start)
            self._set_spin_values(self.dest_header, self.dest_start, dest_header, dest_start)

    def _default_rows(self, side: str) -> tuple[int, int]:
        if side == "source":
            header, start = DEFAULT_SOURCE_HEADER_ROW, DEFAULT_SOURCE_DATA_START
        else:
            header, start = DEFAULT_DEST_HEADER_ROW, DEFAULT_DEST_DATA_START
        profile_name = self.ctx.session.pending_profile_name
        if profile_name:
            profile = self.ctx.mapping_repo.get(profile_name)
            if profile is not None:
                if side == "source":
                    header = profile.source_header_row or header
                    start = profile.source_data_start or start
                else:
                    header = profile.dest_header_row or header
                    start = profile.dest_data_start or start
        return header, start

    def _set_spin_values(self, header_spin: QSpinBox, start_spin: QSpinBox, header: int, start: int) -> None:
        header_spin.blockSignals(True)
        start_spin.blockSignals(True)
        header_spin.setValue(header)
        start_spin.setValue(start)
        header_spin.blockSignals(False)
        start_spin.blockSignals(False)

    def _browse_source(self) -> None:
        default_header, default_start = self._default_rows("source")
        self._browse(
            self.source_path,
            self.source_sheet,
            self.source_header,
            self.source_start,
            default_header,
            default_start,
        )

    def _browse_dest(self) -> None:
        default_header, default_start = self._default_rows("dest")
        self._browse(
            self.dest_path,
            self.dest_sheet,
            self.dest_header,
            self.dest_start,
            default_header,
            default_start,
        )

    def _browse(
        self,
        path_edit,
        sheet_box,
        header_spin,
        start_spin,
        default_header: int,
        default_start: int,
    ) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel",
            str(Path.home()),
            "Excel (*.xlsx *.xlsm)",
        )
        if not selected:
            return
        try:
            names = self.ctx.engine.list_sheets(selected)
        except AppError as exc:
            QMessageBox.warning(self, "Không đọc được file", exc.message)
            return
        if not names:
            QMessageBox.warning(self, "Không đọc được file", "File không có sheet nào.")
            return
        path_edit.setText(selected)
        sheet_box.blockSignals(True)
        sheet_box.clear()
        sheet_box.addItems(names)
        sheet_box.blockSignals(False)
        self._set_spin_values(header_spin, start_spin, default_header, default_start)
        self.notice.setText("")
        self._refresh_previews()

    def _sheet_changed(self, side: str) -> None:
        if side == "source":
            path = self.source_path.text()
            sheet = self.source_sheet.currentText()
            header = self.source_header
            start = self.source_start
        else:
            path = self.dest_path.text()
            sheet = self.dest_sheet.currentText()
            header = self.dest_header
            start = self.dest_start
        default_header, default_start = self._default_rows(side)
        if path and sheet:
            self._set_spin_values(header, start, default_header, default_start)
        self._refresh_previews()

    def _refresh_previews(self) -> None:
        self._show_preview(
            self.source_path.text(),
            self.source_sheet.currentText(),
            self.source_header.value(),
            self.source_start.value(),
            self.source_preview,
        )
        self._show_preview(
            self.dest_path.text(),
            self.dest_sheet.currentText(),
            self.dest_header.value(),
            self.dest_start.value(),
            self.dest_preview,
        )

    def _show_preview(self, path: str, sheet: str, header_row: int, data_start: int, table) -> None:
        if not path or not sheet:
            table.setRowCount(0)
            return
        try:
            max_rows = max(PREVIEW_ROWS, data_start + 5)
            grid = self.ctx.engine.preview(path, sheet, max_rows=max_rows)
        except AppError as exc:
            self.notice.setText(exc.message)
            return
        fill_preview(table, grid, header_row, data_start)

    def on_next(self) -> bool:
        source = self.source_path.text().strip()
        dest = self.dest_path.text().strip()
        if not source or not dest or not self.source_sheet.currentText() or not self.dest_sheet.currentText():
            QMessageBox.warning(self, "Thiếu file", "Hãy chọn cả file nguồn và file đích.")
            return False
        if same_file(source, dest):
            QMessageBox.warning(
                self,
                "Trùng file",
                "Hãy chọn hai file khác nhau. Ứng dụng không sửa file gốc.",
            )
            return False
        if self.source_start.value() <= self.source_header.value() or self.dest_start.value() <= self.dest_header.value():
            QMessageBox.warning(
                self,
                "Dòng dữ liệu chưa đúng",
                "Dòng bắt đầu dữ liệu phải nằm bên dưới dòng tiêu đề.",
            )
            return False
        signature = (
            source,
            dest,
            self.source_sheet.currentText(),
            self.dest_sheet.currentText(),
            self.source_header.value(),
            self.dest_header.value(),
            self.source_start.value(),
            self.dest_start.value(),
        )
        session = self.ctx.session
        if signature != self._signature:
            session.mappings.clear()
            session.manual_pairs.clear()
            session.blocked_pairs.clear()
            session.source_name_key = ""
            session.dest_name_key = ""
            session.acknowledge_warnings = False
            session.review = None
            session.match = None
        self._signature = signature
        session.source_path = source
        session.dest_path = dest
        session.source_sheet = self.source_sheet.currentText()
        session.dest_sheet = self.dest_sheet.currentText()
        session.source_header_row = self.source_header.value()
        session.dest_header_row = self.dest_header.value()
        session.source_data_start = self.source_start.value()
        session.dest_data_start = self.dest_start.value()
        return True
