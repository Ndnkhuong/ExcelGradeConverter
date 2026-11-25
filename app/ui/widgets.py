"""Thành phần giao diện dùng lại."""

from __future__ import annotations

from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QLabel, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget


def title_block(title: str, subtitle: str) -> QVBoxLayout:
    layout = QVBoxLayout()
    heading = QLabel(title)
    heading.setObjectName("pageTitle")
    description = QLabel(subtitle)
    description.setObjectName("pageSubtitle")
    description.setWordWrap(True)
    layout.addWidget(heading)
    layout.addWidget(description)
    layout.setSpacing(4)
    return layout


def readonly_table(headers: list[str]) -> QTableWidget:
    table = QTableWidget(0, len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    table.setAlternatingRowColors(True)
    table.verticalHeader().setVisible(False)
    table.horizontalHeader().setStretchLastSection(True)
    table.setWordWrap(False)
    return table


def text_item(value: object) -> QTableWidgetItem:
    item = QTableWidgetItem("" if value is None else str(value))
    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
    return item


def fill_preview(
    table: QTableWidget,
    grid: list[list[str]],
    header_row: int,
    data_start_row: int | None = None,
) -> None:
    if not grid:
        table.clear()
        table.setRowCount(0)
        table.setColumnCount(0)
        return
    columns = max(len(row) for row in grid)
    table.setColumnCount(columns)
    table.setRowCount(len(grid))
    table.verticalHeader().setVisible(True)
    table.setVerticalHeaderLabels([str(index + 1) for index in range(len(grid))])
    table.setHorizontalHeaderLabels([get_column_letter(index + 1) for index in range(columns)])
    for row_index, row in enumerate(grid):
        for column_index in range(columns):
            value = row[column_index] if column_index < len(row) else ""
            item = text_item(value)
            if row_index + 1 == header_row:
                item.setBackground(QColor("#dbeafe"))
                item.setToolTip("Dòng tiêu đề đang chọn")
            elif data_start_row is not None and row_index + 1 == data_start_row:
                item.setBackground(QColor("#dcfce7"))
                item.setToolTip("Dòng bắt đầu dữ liệu học sinh")
            table.setItem(row_index, column_index, item)


class Page(QWidget):
    def on_enter(self) -> None:
        return None

    def on_next(self) -> bool:
        return True
