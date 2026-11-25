"""Ghép cột nguồn với cột đích và lưu cấu hình để dùng lại."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from app.core.column_mapper import (
    find_column_by_header,
    suggest_default_mapping,
    suggest_mapping_by_header,
)
from app.models.column_mapping import ColumnMappingItem, MappingProfile
from app.models.excel_file import SheetData
from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block
from config import DEFAULT_DEST_SCORE_COLUMN, DEFAULT_SOURCE_SCORE_COLUMN


class MappingPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        self.source: SheetData | None = None
        self.dest: SheetData | None = None
        self._loading = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Ghép cột điểm",
                f"Mặc định ghép cột điểm {DEFAULT_SOURCE_SCORE_COLUMN} (mẫu A) sang cột {DEFAULT_DEST_SCORE_COLUMN} (mẫu B) "
                "và bỏ qua các cột khác để không ghi đè dữ liệu ngoài phạm vi.",
            )
        )

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Họ tên nguồn"))
        self.source_name = QComboBox()
        name_row.addWidget(self.source_name, 1)
        name_row.addWidget(QLabel("Họ tên đích"))
        self.dest_name = QComboBox()
        name_row.addWidget(self.dest_name, 1)
        layout.addLayout(name_row)

        tools = QHBoxLayout()
        default_btn = QPushButton(
            f"Mapping mặc định (Cột {DEFAULT_SOURCE_SCORE_COLUMN} → {DEFAULT_DEST_SCORE_COLUMN})"
        )
        suggest = QPushButton("Gợi ý theo tên cột")
        save = QPushButton("Lưu cấu hình")
        load = QPushButton("Tải cấu hình")
        default_btn.clicked.connect(self._apply_default_mapping)
        suggest.clicked.connect(self._apply_header_suggestion)
        save.clicked.connect(self._save_profile)
        load.clicked.connect(self._load_profile)
        tools.addWidget(default_btn)
        tools.addWidget(suggest)
        tools.addWidget(save)
        tools.addWidget(load)
        tools.addStretch(1)
        layout.addLayout(tools)

        self.table = readonly_table(["Cột đích", "Cột nguồn", "Trạng thái"])
        self.table.setEditTriggers(self.table.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table, 1)
        self.notice = QLabel("")
        self.notice.setObjectName("hint")
        self.notice.setWordWrap(True)
        layout.addWidget(self.notice)

        self.source_name.currentIndexChanged.connect(self._names_changed)
        self.dest_name.currentIndexChanged.connect(self._names_changed)

    def on_enter(self) -> None:
        session = self.ctx.session
        self.source, self.dest = self.ctx.engine.load(session)
        notices = self.source.notices + self.dest.notices
        self.notice.setText(" ".join(notices))
        self._fill_name_combos()
        if session.pending_profile_name and not session.mappings:
            self._apply_named_profile(session.pending_profile_name)
            session.pending_profile_name = ""
        elif not session.mappings:
            self._apply_default_mapping()
        else:
            self._fill_table()

    def on_next(self) -> bool:
        self._capture()
        session = self.ctx.session
        if not session.source_name_key or not session.dest_name_key:
            QMessageBox.warning(self, "Thiếu cột họ tên", "Hãy chọn cột họ tên của cả hai file.")
            return False
        if not session.mappings:
            QMessageBox.warning(self, "Chưa ghép cột điểm", "Hãy ghép ít nhất một cột điểm.")
            return False
        dest_keys = [item.dest_key for item in session.mappings]
        if len(dest_keys) != len(set(dest_keys)):
            QMessageBox.warning(self, "Ghép trùng", "Mỗi cột đích chỉ được nhận một cột nguồn.")
            return False
        session.acknowledge_warnings = False
        return True

    def _fill_name_combos(self) -> None:
        assert self.source is not None and self.dest is not None
        session = self.ctx.session
        suggestion = suggest_default_mapping(self.source, self.dest)
        self._fill_combo(self.source_name, self.source, session.source_name_key or suggestion.source_name_key)
        self._fill_combo(self.dest_name, self.dest, session.dest_name_key or suggestion.dest_name_key)

    def _fill_combo(self, combo: QComboBox, sheet: SheetData, selected: str | None) -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("— Chọn cột —", "")
        for column in sheet.columns:
            combo.addItem(column.label, column.key)
        index = combo.findData(selected or "")
        combo.setCurrentIndex(max(index, 0))
        combo.blockSignals(False)

    def _names_changed(self) -> None:
        if self.source is None:
            return
        self._capture()
        self._fill_table()

    def _apply_default_mapping(self) -> None:
        if self.source is None or self.dest is None:
            return
        suggestion = suggest_default_mapping(self.source, self.dest)
        self._fill_combo(self.source_name, self.source, suggestion.source_name_key)
        self._fill_combo(self.dest_name, self.dest, suggestion.dest_name_key)
        self.ctx.session.source_name_key = str(self.source_name.currentData() or "")
        self.ctx.session.dest_name_key = str(self.dest_name.currentData() or "")
        self.ctx.session.mappings = self._items_from_pairs(suggestion.pairs)
        self.ctx.session.profile_name = ""
        self._fill_table()
        notices = list(self.source.notices) + list(self.dest.notices)
        if suggestion.pairs:
            notices.append(
                f"Đã áp dụng mapping mặc định: cột {DEFAULT_SOURCE_SCORE_COLUMN} (nguồn) → "
                f"cột {DEFAULT_DEST_SCORE_COLUMN} (đích). Các cột khác được giữ nguyên."
            )
        else:
            notices.append(
                f"Không tìm thấy cặp cột mặc định {DEFAULT_SOURCE_SCORE_COLUMN} → {DEFAULT_DEST_SCORE_COLUMN} "
                "(hoặc cột đang trùng cột họ tên). Hãy chọn thủ công."
            )
        self.notice.setText(" ".join(notices))

    def _apply_header_suggestion(self) -> None:
        if self.source is None or self.dest is None:
            return
        suggestion = suggest_mapping_by_header(self.source, self.dest)
        self._fill_combo(self.source_name, self.source, suggestion.source_name_key)
        self._fill_combo(self.dest_name, self.dest, suggestion.dest_name_key)
        self.ctx.session.source_name_key = str(self.source_name.currentData() or "")
        self.ctx.session.dest_name_key = str(self.dest_name.currentData() or "")
        self.ctx.session.mappings = self._items_from_pairs(suggestion.pairs)
        self.ctx.session.profile_name = ""
        self._fill_table()

    def _items_from_pairs(self, pairs: list[tuple[str, str]]) -> list[ColumnMappingItem]:
        assert self.source is not None and self.dest is not None
        items: list[ColumnMappingItem] = []
        used_dest: set[str] = set()
        name_source = self.source_name.currentData()
        name_dest = self.dest_name.currentData()
        for source_key, dest_key in pairs:
            if source_key == name_source or dest_key == name_dest or dest_key in used_dest:
                continue
            source_column = self.source.column_by_key(source_key)
            dest_column = self.dest.column_by_key(dest_key)
            if source_column is None or dest_column is None:
                continue
            used_dest.add(dest_key)
            items.append(
                ColumnMappingItem(
                    source_key=source_key,
                    dest_key=dest_key,
                    source_header=source_column.header,
                    dest_header=dest_column.header,
                )
            )
        return items

    def _fill_table(self) -> None:
        assert self.source is not None and self.dest is not None
        self._loading = True
        try:
            self._fill_table_rows()
        finally:
            self._loading = False
        self._refresh_status()

    def _fill_table_rows(self) -> None:
        assert self.source is not None and self.dest is not None
        selected = {item.dest_key: item.source_key for item in self.ctx.session.mappings}
        name_dest = self.dest_name.currentData()
        name_source = self.source_name.currentData()
        self.table.setRowCount(len(self.dest.columns))
        for row, column in enumerate(self.dest.columns):
            self.table.setItem(row, 0, text_item(column.label))
            if column.key == name_dest:
                self.table.setCellWidget(row, 1, QLabel("Cột họ tên, không chuyển điểm"))
                self.table.setItem(row, 2, text_item("Họ tên"))
                continue
            combo = QComboBox()
            combo.addItem("— Bỏ qua —", "")
            for source_column in self.source.columns:
                if source_column.key == name_source:
                    continue
                combo.addItem(source_column.label, source_column.key)
            current = selected.get(column.key, "")
            index = combo.findData(current)
            combo.blockSignals(True)
            combo.setCurrentIndex(index if index >= 0 else 0)
            combo.blockSignals(False)
            combo.currentIndexChanged.connect(self._capture)
            self.table.setCellWidget(row, 1, combo)
            self.table.setItem(row, 2, text_item("Đã ghép" if current else "Bỏ qua"))

    def _refresh_status(self) -> None:
        name_dest = self.dest_name.currentData()
        if self.dest is None:
            return
        for row, column in enumerate(self.dest.columns):
            if column.key == name_dest:
                continue
            widget = self.table.cellWidget(row, 1)
            if isinstance(widget, QComboBox):
                self.table.setItem(row, 2, text_item("Đã ghép" if widget.currentData() else "Bỏ qua"))

    def _capture(self) -> None:
        if self._loading or self.source is None or self.dest is None:
            return
        session = self.ctx.session
        session.source_name_key = str(self.source_name.currentData() or "")
        session.dest_name_key = str(self.dest_name.currentData() or "")
        items: list[ColumnMappingItem] = []
        used_dest: set[str] = set()
        for row, column in enumerate(self.dest.columns):
            if column.key == session.dest_name_key:
                continue
            widget = self.table.cellWidget(row, 1)
            if not isinstance(widget, QComboBox):
                continue
            source_key = str(widget.currentData() or "")
            if not source_key or source_key == session.source_name_key or column.key in used_dest:
                continue
            source_column = self.source.column_by_key(source_key)
            if source_column is None:
                continue
            used_dest.add(column.key)
            items.append(
                ColumnMappingItem(
                    source_key=source_key,
                    dest_key=column.key,
                    source_header=source_column.header,
                    dest_header=column.header,
                )
            )
        session.mappings = items
        self._refresh_status()

    def _save_profile(self) -> None:
        self._capture()
        session = self.ctx.session
        if not session.mappings or not session.source_name_key or not session.dest_name_key:
            QMessageBox.warning(self, "Chưa đủ thông tin", "Hãy chọn họ tên và ít nhất một cột điểm trước khi lưu.")
            return
        name, accepted = QInputDialog.getText(self, "Lưu cấu hình", "Tên cấu hình:")
        name = name.strip()
        if not accepted or not name:
            return
        if name in self.ctx.mapping_repo.list_names():
            answer = QMessageBox.question(
                self,
                "Ghi đè cấu hình",
                f"Cấu hình “{name}” đã tồn tại. Bạn có muốn ghi đè không?",
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        assert self.source is not None and self.dest is not None
        source_column = self.source.column_by_key(session.source_name_key)
        dest_column = self.dest.column_by_key(session.dest_name_key)
        profile = MappingProfile(
            name=name,
            source_name_header="" if source_column is None else source_column.header,
            dest_name_header="" if dest_column is None else dest_column.header,
            items=list(session.mappings),
            source_header_row=session.source_header_row,
            dest_header_row=session.dest_header_row,
            source_data_start=session.source_data_start,
            dest_data_start=session.dest_data_start,
        )
        self.ctx.mapping_repo.save(profile)
        session.profile_name = name
        QMessageBox.information(self, "Đã lưu", f"Đã lưu cấu hình “{name}”.")

    def _load_profile(self) -> None:
        names = self.ctx.mapping_repo.list_names()
        if not names:
            QMessageBox.information(self, "Chưa có cấu hình", "Bạn chưa lưu cấu hình mapping nào.")
            return
        dialog = _ProfileDialog(names, self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.selected:
            return
        self._apply_named_profile(dialog.selected)

    def _apply_named_profile(self, name: str) -> None:
        profile = self.ctx.mapping_repo.get(name)
        if profile is None or self.source is None or self.dest is None:
            QMessageBox.warning(self, "Không tải được", "Không tìm thấy cấu hình đã chọn.")
            return
        missing: list[str] = []
        source_name = find_column_by_header(self.source.columns, profile.source_name_header)
        dest_name = find_column_by_header(self.dest.columns, profile.dest_name_header)
        if source_name is None:
            missing.append(f"không thấy cột họ tên nguồn “{profile.source_name_header}”")
        if dest_name is None:
            missing.append(f"không thấy cột họ tên đích “{profile.dest_name_header}”")
        self._fill_combo(self.source_name, self.source, None if source_name is None else source_name.key)
        self._fill_combo(self.dest_name, self.dest, None if dest_name is None else dest_name.key)
        pairs: list[tuple[str, str]] = []
        for item in profile.items:
            source_column = find_column_by_header(self.source.columns, item.source_header)
            dest_column = find_column_by_header(self.dest.columns, item.dest_header)
            if source_column is None or dest_column is None:
                missing.append(f"bỏ qua “{item.source_header}” → “{item.dest_header}”")
                continue
            pairs.append((source_column.key, dest_column.key))
        self.ctx.session.mappings = self._items_from_pairs(pairs)
        self.ctx.session.profile_name = name
        self._fill_table()
        if missing:
            self.notice.setText("Đã tải một phần cấu hình: " + "; ".join(missing) + ".")
        else:
            self.notice.setText(f"Đã tải cấu hình “{name}”.")


class _ProfileDialog(QDialog):
    def __init__(self, names: list[str], parent):
        super().__init__(parent)
        self.selected = ""
        self.setWindowTitle("Cấu hình đã lưu")
        self.resize(420, 360)
        layout = QVBoxLayout(self)
        self.list = QListWidget()
        self.list.addItems(names)
        layout.addWidget(self.list)
        buttons = QHBoxLayout()
        use = QPushButton("Dùng cấu hình này")
        use.setObjectName("primary")
        close = QPushButton("Đóng")
        use.clicked.connect(self._accept)
        close.clicked.connect(self.reject)
        buttons.addStretch(1)
        buttons.addWidget(close)
        buttons.addWidget(use)
        layout.addLayout(buttons)

    def _accept(self) -> None:
        item = self.list.currentItem()
        if item is None:
            QMessageBox.information(self, "Chưa chọn", "Hãy chọn một cấu hình.")
            return
        self.selected = item.text()
        self.accept()
