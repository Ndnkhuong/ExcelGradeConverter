"""Xuất file Excel mới. Không ghi đè file nguồn hoặc file đích gốc."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from app.core.errors import AppError
from app.ui.context import AppContext
from app.ui.widgets import Page, title_block
from app.utils.file_utils import assert_output_is_new, open_in_explorer, safe_filename
from config import SUPPORTED_EXTENSIONS


class ExportPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        self._output: Path | None = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Xuất file Excel",
                "File kết quả là một bản sao của file đích, chỉ các ô đã xem trước được cập nhật. "
                "Định dạng, thứ tự học sinh và các cột không ghép được giữ nguyên.",
            )
        )
        layout.addWidget(QLabel("Thư mục lưu"))
        folder_row = QHBoxLayout()
        self.folder = QLineEdit()
        self.folder.setReadOnly(True)
        browse = QPushButton("Chọn thư mục")
        browse.clicked.connect(self._browse)
        folder_row.addWidget(self.folder, 1)
        folder_row.addWidget(browse)
        layout.addLayout(folder_row)

        layout.addWidget(QLabel("Tên file"))
        self.filename = QLineEdit()
        layout.addWidget(self.filename)

        self.export_button = QPushButton("Xuất file Excel")
        self.export_button.setObjectName("primary")
        self.export_button.clicked.connect(self._export)
        layout.addWidget(self.export_button)

        self.open_button = QPushButton("Mở thư mục chứa file kết quả")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self._open_folder)
        layout.addWidget(self.open_button)

        self.result = QLabel("")
        self.result.setWordWrap(True)
        layout.addWidget(self.result)
        layout.addStretch(1)

    def on_enter(self) -> None:
        session = self.ctx.session
        suffix = Path(session.dest_path).suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            suffix = ".xlsx"
        if not session.output_dir:
            session.output_dir = str(Path(session.dest_path).parent)
        if not session.output_name:
            session.output_name = f"BangDiem_DaChuyenDoi_{datetime.now():%Y%m%d_%H%M}{suffix}"
        self.folder.setText(session.output_dir)
        self.filename.setText(session.output_name)
        self.result.setText("")
        self.open_button.setEnabled(self._output is not None and self._output.exists())

    def _browse(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu", self.folder.text() or str(Path.home()))
        if selected:
            self.folder.setText(selected)
            self.ctx.session.output_dir = selected

    def _export(self) -> None:
        session = self.ctx.session
        folder = self.folder.text().strip()
        suffix = Path(session.dest_path).suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            suffix = ".xlsx"
        filename = safe_filename(self.filename.text().strip(), suffix)
        if Path(filename).suffix.lower() != suffix:
            filename = f"{Path(filename).stem}{suffix}"
        self.filename.setText(filename)
        if not folder:
            QMessageBox.warning(self, "Thiếu thư mục", "Hãy chọn thư mục lưu file kết quả.")
            return
        output = Path(folder) / filename
        try:
            assert_output_is_new(output, Path(session.source_path), Path(session.dest_path))
        except AppError as exc:
            QMessageBox.warning(self, "Không xuất được", exc.message)
            return
        if output.exists():
            answer = QMessageBox.question(
                self,
                "File đã tồn tại",
                f"“{output.name}” đã có trong thư mục. Bạn có muốn ghi đè file kết quả này không?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            result = self.ctx.engine.export(session, output, session.acknowledge_warnings)
        except AppError as exc:
            QMessageBox.warning(self, "Không xuất được", exc.message)
            return
        self.ctx.history_repo.add(
            source_name=Path(session.source_path).name,
            dest_name=Path(session.dest_path).name,
            mapping_name=session.profile_name,
            matched_count=result.matched_count,
            unmatched_count=result.unmatched_count,
            result_path=result.output_path,
        )
        self._output = output
        self.open_button.setEnabled(True)
        numeric_note = (
            f", xác minh {result.numeric_cells_verified} ô điểm kiểu số (Number)"
            if result.numeric_cells_verified
            else ""
        )
        region_note = (
            f" Đã chuẩn hóa toàn bộ sheet đích về General bằng Microsoft Excel ({result.region_cells_formatted} ô{numeric_note})."
            if result.region_cells_formatted
            else ""
        )
        hcm_note = (
            f" Cột Mã HCM ({result.hcm_cells_formatted} ô) giữ nguyên giá trị gốc."
            if result.hcm_column_found
            else " Lưu ý: Không tìm thấy cột Mã HCM trong sheet đích."
        )
        self.result.setText(
            f"Đã xuất {output.name}. Đã ghi {result.updated_cells} ô cho {result.matched_count} học sinh."
            f"{region_note}{hcm_note} "
            f"Bỏ qua {result.skipped_cells} ô. File nguồn và file đích gốc không bị thay đổi."
        )
        QMessageBox.information(self, "Xuất thành công", self.result.text())

    def _open_folder(self) -> None:
        if self._output is None:
            return
        try:
            open_in_explorer(self._output)
        except AppError as exc:
            QMessageBox.warning(self, "Không mở được thư mục", exc.message)
