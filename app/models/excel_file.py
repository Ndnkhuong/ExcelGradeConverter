"""Thông tin file Excel đã đọc."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class CellData:
    row: int
    column: int
    value: object
    is_formula: bool = False
    unresolved_formula: bool = False


@dataclass
class ColumnInfo:
    index: int
    letter: str
    header: str
    key: str

    @property
    def label(self) -> str:
        return f"{self.letter} - {self.header}"


@dataclass
class RowData:
    row_index: int
    cells: dict[str, CellData] = field(default_factory=dict)

    def get(self, key: str) -> CellData | None:
        return self.cells.get(key)


@dataclass
class SheetData:
    path: str
    sheet_name: str
    sheet_names: list[str]
    header_row: int
    data_start_row: int
    columns: list[ColumnInfo]
    rows: list[RowData]
    preview: list[list[str]] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)

    def column_by_key(self, key: str) -> ColumnInfo | None:
        for column in self.columns:
            if column.key == key:
                return column
        return None
