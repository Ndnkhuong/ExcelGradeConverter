"""Đọc workbook Excel. Không giả định dữ liệu bắt đầu ở dòng 1."""

from __future__ import annotations

import warnings
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from zipfile import BadZipFile

from app.core.errors import AppError
from app.core.template_detector import detect_header_row
from app.models.excel_file import CellData, ColumnInfo, RowData, SheetData
from app.utils.file_utils import ensure_supported_excel
from app.utils.name_normalizer import display_name
from config import EMPTY_ROW_STREAK, MAX_DATA_ROWS, MAX_SCAN_COLUMNS, PREVIEW_COLUMNS, PREVIEW_ROWS

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def list_sheet_names(path: str | Path) -> list[str]:
    workbook = _open_workbook(path, data_only=False)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def suggest_header_row(path: str | Path, sheet_name: str) -> int:
    workbook = _open_workbook(path, data_only=False)
    try:
        worksheet = _get_sheet(workbook, sheet_name)
        grid = _preview_grid(worksheet, PREVIEW_ROWS, PREVIEW_COLUMNS)
        values = [[_parse_preview(item) for item in row] for row in grid]
        return detect_header_row(values)
    finally:
        workbook.close()


def read_preview(path: str | Path, sheet_name: str, max_rows: int = PREVIEW_ROWS) -> list[list[str]]:
    workbook = _open_workbook(path, data_only=False)
    try:
        worksheet = _get_sheet(workbook, sheet_name)
        return _preview_grid(worksheet, max_rows, PREVIEW_COLUMNS)
    finally:
        workbook.close()


def read_sheet(
    path: str | Path,
    sheet_name: str,
    header_row: int,
    data_start_row: int,
) -> SheetData:
    file_path = Path(path)
    ensure_supported_excel(file_path)
    if header_row < 1 or data_start_row <= header_row:
        raise AppError("Dòng dữ liệu phải nằm bên dưới dòng tiêu đề.")

    workbook = _open_workbook(file_path, data_only=False)
    value_book = _open_workbook(file_path, data_only=True)
    try:
        worksheet = _get_sheet(workbook, sheet_name)
        value_sheet = _get_sheet(value_book, sheet_name)
        columns = _read_columns(worksheet, header_row, data_start_row)
        if not columns:
            raise AppError("Không tìm thấy cột nào trên dòng tiêu đề đã chọn.")
        notices: list[str] = []
        rows = _read_rows(worksheet, value_sheet, columns, data_start_row, notices)
        preview = _preview_grid(worksheet, PREVIEW_ROWS, min(PREVIEW_COLUMNS, len(columns)))
        return SheetData(
            path=str(file_path),
            sheet_name=sheet_name,
            sheet_names=list(workbook.sheetnames),
            header_row=header_row,
            data_start_row=data_start_row,
            columns=columns,
            rows=rows,
            preview=preview,
            notices=notices,
        )
    finally:
        workbook.close()
        value_book.close()


def _open_workbook(path: str | Path, data_only: bool) -> Workbook:
    from openpyxl import load_workbook

    file_path = Path(path)
    try:
        ensure_supported_excel(file_path)
    except AppError:
        raise
    try:
        return load_workbook(file_path, data_only=data_only, read_only=False, keep_vba=False)
    except PermissionError as exc:
        raise AppError("File đang được chương trình khác giữ. Hãy đóng file rồi thử lại.") from exc
    except (BadZipFile, InvalidFileException, OSError, KeyError) as exc:
        raise AppError("File Excel bị hỏng, bị khóa hoặc không đọc được.") from exc


def _get_sheet(workbook: Workbook, sheet_name: str):
    if sheet_name not in workbook.sheetnames:
        raise AppError(f"Không tìm thấy sheet “{sheet_name}”.")
    return workbook[sheet_name]


def _read_columns(worksheet, header_row: int, data_start_row: int | None = None) -> list[ColumnInfo]:
    merged_headers: dict[int, object] = {}
    for merged in worksheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged))
        if min_row <= header_row <= max_row:
            anchor = worksheet.cell(min_row, min_col).value
            if anchor is not None and str(anchor).strip() != "":
                for col_idx in range(min_col, min(max_col, MAX_SCAN_COLUMNS) + 1):
                    merged_headers.setdefault(col_idx, anchor)

    last_data_col = 0
    if data_start_row is not None:
        sample_end = data_start_row + 20
        for row_idx in range(data_start_row, sample_end):
            for col_idx in range(1, MAX_SCAN_COLUMNS + 1):
                cell_val = worksheet.cell(row_idx, col_idx).value
                if cell_val is not None and str(cell_val).strip() != "":
                    if col_idx > last_data_col:
                        last_data_col = col_idx

    raw_headers: list[tuple[int, object]] = []
    for index in range(1, MAX_SCAN_COLUMNS + 1):
        value = worksheet.cell(header_row, index).value
        if value is None or str(value).strip() == "":
            value = merged_headers.get(index)
        if value is None or str(value).strip() == "":
            raw_headers.append((index, None))
            continue
        raw_headers.append((index, value))
    while raw_headers and raw_headers[-1][1] is None and raw_headers[-1][0] > last_data_col:
        raw_headers.pop()

    seen: dict[str, int] = {}
    columns: list[ColumnInfo] = []
    for index, value in raw_headers:
        letter = get_column_letter(index)
        header = display_name(value) if value is not None else f"Cột {letter}"
        if not header:
            header = f"Cột {letter}"
        count = seen.get(header, 0) + 1
        seen[header] = count
        if count > 1:
            header = f"{header} ({letter})"
        columns.append(ColumnInfo(index=index, letter=letter, header=header, key=str(index)))
    return columns


def _read_rows(worksheet, value_sheet, columns: list[ColumnInfo], data_start: int, notices: list[str]) -> list[RowData]:
    rows: list[RowData] = []
    empty_streak = 0
    row_index = data_start
    while row_index < data_start + MAX_DATA_ROWS:
        cells: dict[str, CellData] = {}
        empty = True
        for column in columns:
            formula_cell = worksheet.cell(row_index, column.index)
            cached_cell = value_sheet.cell(row_index, column.index)
            is_formula = formula_cell.data_type == "f"
            if is_formula:
                cached = cached_cell.value
                unresolved = cached is None
                value = cached
            else:
                unresolved = False
                value = formula_cell.value
            if value not in (None, ""):
                empty = False
            cells[column.key] = CellData(
                row=row_index,
                column=column.index,
                value=value,
                is_formula=is_formula,
                unresolved_formula=unresolved,
            )
        if empty:
            empty_streak += 1
            if empty_streak >= EMPTY_ROW_STREAK:
                break
        else:
            empty_streak = 0
            rows.append(RowData(row_index=row_index, cells=cells))
        row_index += 1
    else:
        notices.append(
            f"Đã dừng ở {MAX_DATA_ROWS} dòng dữ liệu. File dài hơn giới hạn xử lý an toàn."
        )
    return rows


def _preview_grid(worksheet, max_rows: int, max_cols: int) -> list[list[str]]:
    grid: list[list[str]] = []
    for row_index in range(1, max_rows + 1):
        row = []
        for column_index in range(1, max_cols + 1):
            value = worksheet.cell(row_index, column_index).value
            row.append("" if value is None else str(value))
        grid.append(row)
    return grid


def _parse_preview(value: str) -> object:
    if value == "":
        return None
    try:
        number = float(value)
    except ValueError:
        return value
    if number.is_integer() and "." not in value:
        return int(number)
    return number
