"""Đối chiếu học sinh theo họ tên. Tên trùng không được ghép tự động."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from app.core.errors import AppError
from app.models.student_record import MatchStatus, StudentRecord
from app.ui.context import AppContext
from app.ui.widgets import Page, readonly_table, text_item, title_block

_STATUS = {
    "auto": "Đã khớp",
    "manual": "Ghép thủ công",
    "unmatched": "Chưa khớp",
    "dest": "Chỉ có ở đích",
    "ambiguous": "Cần xác nhận",
}


class MatchingPage(Page):
    def __init__(self, ctx: AppContext):
        super().__init__()
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(12)
        layout.addLayout(
            title_block(
                "Đối chiếu học sinh",
                "Chỉ ghép theo họ tên đã chuẩn hóa. Tên phải xuất hiện đúng một lần ở mỗi file "
                "thì mới được ghép tự động. Tên trùng hoặc không có trong file đích cần bạn xử lý.",
            )
        )
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Hiển thị"))
        self.filter = QComboBox()
        self.filter.addItems(["Tất cả", "Đã khớp", "Chưa khớp", "Chỉ có ở đích", "Cần xác nhận"])
        self.filter.currentIndexChanged.connect(self._fill_table)
        filters.addWidget(self.filter)
        filters.addStretch(1)
        clear = QPushButton("Xóa ghép thủ công")
        clear.clicked.connect(self._clear_manual)
        filters.addWidget(clear)
        layout.addLayout(filters)

        self.table = readonly_table(["Trạng thái", "Dòng nguồn", "Tên nguồn", "Dòng đích", "Tên đích", "Ghi chú"])
        layout.addWidget(self.table, 1)

        manual = QHBoxLayout()
        self.source_combo = QComboBox()
        self.dest_combo = QComboBox()
        pair = QPushButton("Ghép thủ công")
        pair.setObjectName("primary")
        unpair = QPushButton("Bỏ ghép dòng đang chọn")
        pair.clicked.connect(self._pair)
        unpair.clicked.connect(self._unpair)
        manual.addWidget(self.source_combo, 1)
        manual.addWidget(self.dest_combo, 1)
        manual.addWidget(pair)
        manual.addWidget(unpair)
        layout.addLayout(manual)

    def on_enter(self) -> None:
        self._reload()

    def on_next(self) -> bool:
        match = self.ctx.session.match
        if match is None:
            return False
        if match.ambiguous:
            answer = QMessageBox.question(
                self,
                "Còn tên trùng",
                "Một số học sinh trùng tên chưa được ghép. Điểm của những học sinh này sẽ không được ghi. Tiếp tục?",
            )
            return answer == QMessageBox.StandardButton.Yes
        return True

    def _reload(self) -> None:
        try:
            _source, _dest, match, _review = self.ctx.engine.review(self.ctx.session)
        except AppError:
            raise
        self.summary.setText(
            f"Đã khớp: {match.matched_count}. "
            f"Chưa có trong file đích: {len(match.unmatched_source)}. "
            f"Chỉ có ở file đích: {len(match.dest_only)}. "
            f"Nhóm trùng tên: {len(match.ambiguous)}."
        )
        self._fill_combos()
        self._fill_table()

    def _fill_combos(self) -> None:
        match = self.ctx.session.match
        self.source_combo.clear()
        self.dest_combo.clear()
        if match is None:
            return
        sources = list(match.unmatched_source)
        dests = list(match.dest_only)
        for group in match.ambiguous:
            sources.extend(group.sources)
            dests.extend(group.dests)
        self.source_combo.addItem("— Học sinh nguồn chưa khớp —", -1)
        self.dest_combo.addItem("— Học sinh đích chưa khớp —", -1)
        for student in sources:
            self.source_combo.addItem(f"Dòng {student.row_index} — {student.name_raw}", student.row_index)
        for student in dests:
            self.dest_combo.addItem(f"Dòng {student.row_index} — {student.name_raw}", student.row_index)

    def _fill_table(self) -> None:
        match = self.ctx.session.match
        self.table.setRowCount(0)
        if match is None:
            return
        rows: list[tuple] = []
        for pair in match.matched:
            kind = "auto" if pair.status == MatchStatus.AUTO else "manual"
            rows.append(
                (
                    kind,
                    pair.source.row_index,
                    pair.source.name_raw,
                    pair.dest.row_index,
                    pair.dest.name_raw,
                    pair.note,
                    {"kind": "matched", "source": pair.source.row_index, "dest": pair.dest.row_index, "status": kind},
                )
            )
        for student in match.unmatched_source:
            rows.append(self._one_side("unmatched", student, None, "Không có trong file đích"))
        for student in match.dest_only:
            rows.append(self._one_side("dest", None, student, "Chỉ có trong file đích, giữ nguyên"))
        for group in match.ambiguous:
            for student in group.sources:
                rows.append(self._one_side("ambiguous", student, None, "Trùng tên, cần ghép thủ công"))
            for student in group.dests:
                rows.append(self._one_side("ambiguous", None, student, "Trùng tên ở file đích"))

        selected_filter = self.filter.currentText()
        visible = [row for row in rows if _filter_matches(selected_filter, row[0])]
        self.table.setRowCount(len(visible))
        for index, row in enumerate(visible):
            values = [_STATUS.get(row[0], row[0]), row[1], row[2], row[3], row[4], row[5]]
            for column, value in enumerate(values):
                item = text_item("" if value in ("", None) else value)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row[6])
                self.table.setItem(index, column, item)

    def _one_side(self, kind: str, source: StudentRecord | None, dest: StudentRecord | None, note: str):
        return (
            kind,
            "" if source is None else source.row_index,
            "" if source is None else source.name_raw,
            "" if dest is None else dest.row_index,
            "" if dest is None else dest.name_raw,
            note,
            {"kind": kind},
        )

    def _pair(self) -> None:
        source_row = self.source_combo.currentData()
        dest_row = self.dest_combo.currentData()
        if source_row in (None, -1) or dest_row in (None, -1):
            QMessageBox.information(self, "Chưa chọn đủ", "Hãy chọn một học sinh nguồn và một học sinh đích.")
            return
        pairs = self.ctx.session.manual_pairs
        if any(source_row == left or dest_row == right for left, right in pairs):
            QMessageBox.warning(self, "Đã được ghép", "Mỗi học sinh chỉ được ghép với một học sinh.")
            return
        pairs.append((int(source_row), int(dest_row)))
        self.ctx.session.blocked_pairs = [
            pair for pair in self.ctx.session.blocked_pairs if pair != (int(source_row), int(dest_row))
        ]
        self._reload()

    def _unpair(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Chưa chọn dòng", "Hãy chọn một dòng đã ghép.")
            return
        item = self.table.item(row, 0)
        data = None if item is None else item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(data, dict) or data.get("kind") != "matched":
            QMessageBox.information(self, "Không bỏ được", "Chỉ bỏ được dòng đã ghép.")
            return
        pair = (int(data["source"]), int(data["dest"]))
        if data.get("status") == "manual":
            self.ctx.session.manual_pairs = [item for item in self.ctx.session.manual_pairs if item != pair]
        else:
            if pair not in self.ctx.session.blocked_pairs:
                self.ctx.session.blocked_pairs.append(pair)
        self._reload()

    def _clear_manual(self) -> None:
        if not self.ctx.session.manual_pairs and not self.ctx.session.blocked_pairs:
            return
        answer = QMessageBox.question(self, "Xóa ghép thủ công", "Xóa toàn bộ ghép thủ công và các lần bỏ ghép?")
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.ctx.session.manual_pairs.clear()
        self.ctx.session.blocked_pairs.clear()
        self._reload()


def _filter_matches(selected: str, kind: str) -> bool:
    if selected == "Tất cả":
        return True
    return {
        "Đã khớp": kind in {"auto", "manual"},
        "Chưa khớp": kind == "unmatched",
        "Chỉ có ở đích": kind == "dest",
        "Cần xác nhận": kind == "ambiguous",
    }.get(selected, True)
