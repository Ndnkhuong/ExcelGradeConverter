"""Cửa sổ chính và điều hướng quy trình 7 bước."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from app.core.conversion_engine import ConversionEngine
from app.core.errors import AppError
from app.database.database import Database
from app.database.history_repository import HistoryRepository
from app.database.mapping_repository import MappingRepository
from app.database.settings_store import SettingsStore
from app.models.session import ConversionSession
from app.ui.context import AppContext
from app.ui.dashboard_page import DashboardPage
from app.ui.export_page import ExportPage
from app.ui.file_selection_page import FileSelectionPage
from app.ui.history_page import HistoryPage
from app.ui.mapping_page import MappingPage
from app.ui.matching_page import MatchingPage
from app.ui.preview_page import PreviewPage
from app.ui.settings_page import SettingsPage
from app.ui.validation_page import ValidationPage
from app.utils.logger import get_logger
from config import APP_TITLE, APP_VERSION, database_path, settings_path

DASHBOARD = "dashboard"
FILES = "files"
MAPPING = "mapping"
MATCHING = "matching"
VALIDATION = "validation"
PREVIEW = "preview"
EXPORT = "export"
HISTORY = "history"
SETTINGS = "settings"
WIZARD = [FILES, MAPPING, MATCHING, VALIDATION, PREVIEW, EXPORT]
NEXT_PAGES = {FILES, MAPPING, MATCHING, VALIDATION, PREVIEW}


def load_styles() -> str:
    if getattr(sys, "frozen", False):
        path = Path(sys._MEIPASS) / "app" / "resources" / "styles" / "app.qss"  # type: ignore[attr-defined]
    else:
        path = Path(__file__).resolve().parents[1] / "resources" / "styles" / "app.qss"
    if path.exists():
        return path.read_text(encoding="utf-8")
    return ""


class MainWindow(QMainWindow):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        ctx.window = self
        self.current = DASHBOARD
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION}")
        self.resize(1180, 780)
        self.setMinimumSize(1024, 680)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.stepper = QHBoxLayout()
        self.step_labels: list[QLabel] = []
        self.step_bar = QWidget()
        self.step_bar.setLayout(self.stepper)
        self.stepper.setContentsMargins(20, 12, 20, 0)
        for text in [
            "1. Chọn file",
            "2. Ghép cột",
            "3. Đối chiếu",
            "4. Kiểm tra",
            "5. Xem trước",
            "6. Xuất file",
        ]:
            label = QLabel(text)
            label.setObjectName("step")
            self.step_labels.append(label)
            self.stepper.addWidget(label)
        self.stepper.addStretch(1)
        layout.addWidget(self.step_bar)

        self.stack = QStackedWidget()
        self.pages = {
            DASHBOARD: DashboardPage(ctx),
            FILES: FileSelectionPage(ctx),
            MAPPING: MappingPage(ctx),
            MATCHING: MatchingPage(ctx),
            VALIDATION: ValidationPage(ctx),
            PREVIEW: PreviewPage(ctx),
            EXPORT: ExportPage(ctx),
            HISTORY: HistoryPage(ctx),
            SETTINGS: SettingsPage(ctx),
        }
        for page in self.pages.values():
            self.stack.addWidget(page)
        layout.addWidget(self.stack, 1)

        footer = QWidget()
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(28, 8, 28, 16)
        self.back_button = QPushButton("Quay lại")
        self.next_button = QPushButton("Tiếp tục")
        self.next_button.setObjectName("primary")
        self.back_button.clicked.connect(self._back)
        self.next_button.clicked.connect(self._next)
        footer_layout.addWidget(self.back_button)
        footer_layout.addStretch(1)
        footer_layout.addWidget(self.next_button)
        self.footer = footer
        layout.addWidget(footer)

        author_footer = QFrame()
        author_footer.setObjectName("authorFooter")
        author_layout = QVBoxLayout(author_footer)
        author_layout.setContentsMargins(16, 6, 16, 8)
        author_layout.setSpacing(0)
        self.author_label = QLabel(
            '<b>Excel Grade Converter</b><br>'
            'Developed by Ndnkhuong<br>'
            'Link mã nguồn: <a href="https://github.com/Ndnkhuong/ExcelGradeConverter" '
            'style="color:#1f6feb; text-decoration:none;">'
            'https://github.com/Ndnkhuong/ExcelGradeConverter</a><br>'
            'Liên hệ: 0962 439 389'
        )
        self.author_label.setObjectName("authorInfo")
        self.author_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.author_label.setOpenExternalLinks(True)
        self.author_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        author_layout.addWidget(self.author_label)
        self.author_footer = author_footer
        layout.addWidget(author_footer)

        self.pages[DASHBOARD].start_button.clicked.connect(self._start)
        self.pages[DASHBOARD].profile_button.clicked.connect(self._start_from_profile)
        self.pages[DASHBOARD].history_button.clicked.connect(lambda: self.go(HISTORY))
        self.pages[DASHBOARD].settings_button.clicked.connect(lambda: self.go(SETTINGS))
        self.go(DASHBOARD)

    def go(self, page_id: str) -> None:
        page = self.pages[page_id]
        try:
            page.on_enter()
        except AppError as exc:
            QMessageBox.warning(self, "Không thực hiện được", exc.message)
            return
        except Exception:
            get_logger().exception("Lỗi khi mở màn hình %s", page_id)
            QMessageBox.critical(
                self,
                "Có lỗi",
                "Không mở được màn hình này. Chi tiết đã ghi vào nhật ký local, không gồm điểm học sinh.",
            )
            return
        self.stack.setCurrentWidget(page)
        self.current = page_id
        self._update_chrome()

    def _start(self) -> None:
        self.ctx.session.reset()
        self.pages[FILES]._signature = None
        self.go(FILES)

    def _start_from_profile(self) -> None:
        names = self.ctx.mapping_repo.list_names()
        if not names:
            QMessageBox.information(self, "Chưa có cấu hình", "Bạn chưa lưu cấu hình mapping nào.")
            return
        from app.ui.mapping_page import _ProfileDialog
        from PySide6.QtWidgets import QDialog

        dialog = _ProfileDialog(names, self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.selected:
            return
        self.ctx.session.reset()
        self.ctx.session.pending_profile_name = dialog.selected
        self.ctx.session.profile_name = dialog.selected
        self.pages[FILES]._signature = None
        self.go(FILES)

    def _next(self) -> None:
        page = self.pages[self.current]
        try:
            if page.on_next() is False:
                return
        except AppError as exc:
            QMessageBox.warning(self, "Không thực hiện được", exc.message)
            return
        except Exception:
            get_logger().exception("Lỗi khi chuyển bước từ %s", self.current)
            QMessageBox.critical(self, "Có lỗi", "Không chuyển được bước tiếp theo.")
            return
        if self.current in WIZARD:
            index = WIZARD.index(self.current)
            if index + 1 < len(WIZARD):
                self.go(WIZARD[index + 1])

    def _back(self) -> None:
        if self.current in (HISTORY, SETTINGS):
            self.go(DASHBOARD)
            return
        if self.current in WIZARD:
            index = WIZARD.index(self.current)
            self.go(DASHBOARD if index == 0 else WIZARD[index - 1])

    def _update_chrome(self) -> None:
        self.footer.setVisible(self.current != DASHBOARD)
        self.next_button.setVisible(self.current in NEXT_PAGES)
        self.step_bar.setVisible(self.current in WIZARD)
        for index, page_id in enumerate(WIZARD):
            self.step_labels[index].setObjectName("stepCurrent" if page_id == self.current else "step")
            self.step_labels[index].style().unpolish(self.step_labels[index])
            self.step_labels[index].style().polish(self.step_labels[index])


def run() -> None:
    database = Database(database_path())
    settings = SettingsStore(settings_path(), database)
    ctx = AppContext(
        session=ConversionSession(),
        settings=settings,
        mapping_repo=MappingRepository(database),
        history_repo=HistoryRepository(database),
        engine=ConversionEngine(settings),
    )
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(load_styles())
    window = MainWindow(ctx)
    window.show()
    app.exec()
