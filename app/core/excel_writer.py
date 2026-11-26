"""Ghi file kết quả bằng cách sao chép file đích rồi chỉ sửa các ô được mapping."""

from __future__ import annotations

import re
import shutil
import uuid
from dataclasses import dataclass
from pathlib import Path

from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter, range_boundaries

from app.core.errors import AppError
from app.core.excel_com import normalize_sheet_with_excel_com
from app.utils.file_utils import assert_output_is_new, cleanup_temp
from app.utils.logger import get_logger
from app.utils.validators import coerce_cell_value, is_blank, is_number
from config import temp_dir

_QUOTED_OR_BRACKETED_RE = re.compile(r'"[^"]*"|\[[^\]]*\]')


@dataclass
class CellUpdate:
    row: int
    column: int
    value: object


@dataclass
class HcmColumnSpec:
    column: int
    start_row: int
    end_row: int


@dataclass
class DataRegionSpec:
    start_row: int
    end_row: int
    start_col: int
    end_col: int


def is_valid_number_format(fmt: str | None) -> bool:
    """Kiểm tra number_format có phải định dạng số hợp lệ của Excel (không phải Text '@')."""
    if not fmt:
        return False
    cleaned = fmt.strip()
    if not cleaned:
        return False
    first_section = _QUOTED_OR_BRACKETED_RE.sub("", cleaned.split(";")[0]).strip()
    if first_section.lower() == "general":
        return True
    if "@" in first_section:
        return False
    return any(ch in first_section for ch in ("0", "#", "?"))


def write_updates(
    dest_path: str | Path,
    output_path: str | Path,
    sheet_name: str,
    updates: list[CellUpdate],
    source_path: str | Path | None = None,
    hcm_column: HcmColumnSpec | None = None,
    data_region: DataRegionSpec | None = None,
) -> int:
    """Tạo file mới từ bản sao của file đích và xác minh kết quả. Trả về số ô đã ghi."""
    written, _verified, _hcm, _region = write_updates_with_report(
        dest_path, output_path, sheet_name, updates, source_path, hcm_column, data_region
    )
    return written


def write_updates_with_report(
    dest_path: str | Path,
    output_path: str | Path,
    sheet_name: str,
    updates: list[CellUpdate],
    source_path: str | Path | None = None,
    hcm_column: HcmColumnSpec | None = None,
    data_region: DataRegionSpec | None = None,
) -> tuple[int, int, int, int]:
    """Tạo file mới từ bản sao của file đích.

    Trả về (số ô đã ghi, số ô điểm Number đã xác minh, số ô Mã HCM đã xác minh, số ô vùng dữ liệu đã chuẩn hóa General).
    Chỉ thay đổi number_format trong vùng dữ liệu về General; giữ nguyên tiêu đề, style, công thức và các sheet khác.
    """
    destination = Path(dest_path)
    output = Path(output_path)
    if source_path is not None:
        assert_output_is_new(output, Path(source_path), destination)
    else:
        if output.resolve() == destination.resolve():
            raise AppError("Không ghi đè file đích gốc.")
    if not updates:
        raise AppError("Không có ô nào được phép ghi.")

    output.parent.mkdir(parents=True, exist_ok=True)
    temp_path = temp_dir() / f"export_{uuid.uuid4().hex}{destination.suffix}"
    written = 0
    written_cells: list[tuple[int, int, object]] = []
    hcm_snapshots: list[tuple[int, int, object, str]] = []
    sheet_snapshots: list[tuple[int, int, object, str]] = []
    try:
        shutil.copy2(destination, temp_path)
        written, written_cells, hcm_snapshots, sheet_snapshots = _apply_updates(
            temp_path, sheet_name, updates, hcm_column, data_region
        )
        # Bước cuối cùng xử lý và lưu file: chuẩn hóa toàn bộ worksheet đích bằng Microsoft Excel COM.
        normalize_sheet_with_excel_com(temp_path, sheet_name)
        shutil.copy2(temp_path, output)
    except PermissionError as exc:
        raise AppError(
            "Không ghi được file kết quả. Hãy đóng file nếu đang mở và chọn thư mục bạn có quyền ghi."
        ) from exc
    except AppError:
        raise
    except OSError as exc:
        raise AppError("Không tạo được file kết quả.") from exc
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                cleanup_temp()
    # Chỉ đọc (không lưu lại bằng openpyxl) để xác minh kết quả sau khi Excel COM đã lưu.
    total_verified, verified_numeric, verified_hcm, verified_sheet = verify_exported_cells(
        output,
        sheet_name,
        written_cells,
        hcm_snapshots,
        sheet_snapshots,
        verify_entire_sheet=True,
    )
    get_logger().info(
        "Đã ghi %s ô vào file kết quả %s qua Excel COM (General toàn bộ sheet đích %s ô, %s ô đã map, %s ô điểm Number, %s ô Mã HCM)",
        written,
        output.name,
        verified_sheet,
        total_verified,
        verified_numeric,
        verified_hcm,
    )
    return written, verified_numeric, verified_hcm, verified_sheet


def verify_exported_cells(
    path: str | Path,
    sheet_name: str,
    expected_cells: list[tuple[int, int, object]],
    hcm_snapshots: list[tuple[int, int, object, str]] | None = None,
    sheet_snapshots: list[tuple[int, int, object, str]] | None = None,
    verify_entire_sheet: bool = False,
) -> tuple[int, int, int, int]:
    """Mở lại file kết quả sau khi xuất để kiểm tra:
    - Toàn bộ ô trên sheet đích đều có number_format == "General" và giữ nguyên giá trị & kiểu thực tế.
    - Tất cả ô đã map đều có number_format == "General" và đúng kiểu số/văn bản.
    - Toàn bộ ô dữ liệu cột Mã HCM có number_format == "General" và giữ nguyên chính xác giá trị & kiểu gốc.
    """
    hcm_list = list(hcm_snapshots or [])
    sheet_list = list(sheet_snapshots or [])
    if not expected_cells and not hcm_list and not sheet_list and not verify_entire_sheet:
        return 0, 0, 0, 0
    from openpyxl import load_workbook

    file_path = Path(path)
    workbook = load_workbook(file_path, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise AppError(f"Sheet “{sheet_name}” không tồn tại trong file kết quả.")
        worksheet = workbook[sheet_name]
        invalid_format: list[str] = []
        invalid_numeric: list[str] = []
        invalid_text: list[str] = []
        invalid_hcm: list[str] = []
        invalid_sheet_format: list[str] = []
        invalid_sheet_value: list[str] = []
        numeric_count = 0
        for row, column, expected_value in expected_cells:
            cell = worksheet.cell(row, column)
            coord = f"{get_column_letter(column)}{row}"
            if cell.number_format != "General":
                invalid_format.append(coord)
            if is_number(expected_value):
                numeric_count += 1
                if (
                    cell.data_type != "n"
                    or not is_number(cell.value)
                    or float(cell.value) != float(expected_value)
                ):
                    invalid_numeric.append(coord)
            elif isinstance(expected_value, str):
                if (
                    cell.data_type != "s"
                    or not isinstance(cell.value, str)
                    or cell.value != expected_value
                ):
                    invalid_text.append(coord)
        for row, column, original_value, original_data_type in hcm_list:
            cell = worksheet.cell(row, column)
            coord = f"{get_column_letter(column)}{row}"
            if (
                cell.number_format != "General"
                or not _values_match(cell.value, cell.data_type, original_value, original_data_type)
            ):
                invalid_hcm.append(coord)

        checked_sheet_cells = 0
        if verify_entire_sheet:
            for row in worksheet.iter_rows():
                for cell in row:
                    checked_sheet_cells += 1
                    if cell.number_format != "General":
                        invalid_sheet_format.append(f"{get_column_letter(cell.column)}{cell.row}")
        for row, column, expected_value, expected_data_type in sheet_list:
            cell = worksheet.cell(row, column)
            coord = f"{get_column_letter(column)}{row}"
            if not verify_entire_sheet and cell.number_format != "General":
                invalid_sheet_format.append(coord)
            if not _values_match(cell.value, cell.data_type, expected_value, expected_data_type):
                invalid_sheet_value.append(coord)
        if not verify_entire_sheet:
            checked_sheet_cells = len(sheet_list)

        if invalid_format:
            preview = ", ".join(invalid_format[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_format)} ô đã map chưa có Number Format 'General': {preview}."
            )
        if invalid_numeric:
            preview = ", ".join(invalid_numeric[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_numeric)} ô điểm chưa đúng kiểu dữ liệu số (Number): {preview}."
            )
        if invalid_text:
            preview = ", ".join(invalid_text[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_text)} ô văn bản chưa đúng kiểu dữ liệu chữ (Text): {preview}."
            )
        if invalid_hcm:
            preview = ", ".join(invalid_hcm[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_hcm)} ô cột Mã HCM sai định dạng General hoặc bị thay đổi giá trị: {preview}."
            )
        if invalid_sheet_format:
            preview = ", ".join(invalid_sheet_format[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_sheet_format)} ô trên sheet đích chưa có Number Format 'General': {preview}."
            )
        if invalid_sheet_value:
            preview = ", ".join(invalid_sheet_value[:10])
            raise AppError(
                f"Kiểm tra sau khi xuất phát hiện {len(invalid_sheet_value)} ô trên sheet đích bị thay đổi giá trị hoặc kiểu dữ liệu: {preview}."
            )
        return len(expected_cells), numeric_count, len(hcm_list), checked_sheet_cells
    finally:
        workbook.close()


def verify_numeric_cells(
    path: str | Path,
    sheet_name: str,
    expected_cells: list[tuple[int, int, int | float]],
) -> int:
    """Kiểm tra các ô điểm đã ghi có đúng kiểu Number của Excel và định dạng General."""
    _total, numeric_count, _hcm, _sheet = verify_exported_cells(
        path, sheet_name, [(row, col, val) for row, col, val in expected_cells]
    )
    return numeric_count


def _apply_updates(
    path: Path,
    sheet_name: str,
    updates: list[CellUpdate],
    hcm_column: HcmColumnSpec | None = None,
    data_region: DataRegionSpec | None = None,
) -> tuple[
    int,
    list[tuple[int, int, object]],
    list[tuple[int, int, object, str]],
    list[tuple[int, int, object, str]],
]:
    from openpyxl import load_workbook

    keep_vba = path.suffix.lower() == ".xlsm"
    workbook = load_workbook(path, keep_vba=keep_vba)
    try:
        if sheet_name not in workbook.sheetnames:
            raise AppError(f"Sheet “{sheet_name}” không còn trong file đích.")
        worksheet = workbook[sheet_name]
        protected: set[tuple[int, int]] = set()
        horizontal_anchors: dict[tuple[int, int], tuple[int, int]] = {}
        for merged in worksheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged))
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    if not (row == min_row and column == min_col):
                        if min_row == max_row:
                            horizontal_anchors[(row, column)] = (min_row, min_col)
                        else:
                            protected.add((row, column))
        written = 0
        written_cells: list[tuple[int, int, object]] = []
        for update in updates:
            if is_blank(update.value):
                continue
            target = horizontal_anchors.get((update.row, update.column), (update.row, update.column))
            if target in protected:
                continue
            cell = worksheet.cell(target[0], target[1])
            if _is_formula(cell):
                continue
            value = coerce_cell_value(update.value)
            cell.value = value
            cell.number_format = "General"
            written_cells.append((target[0], target[1], value))
            written += 1

        hcm_snapshots: list[tuple[int, int, object, str]] = []
        if hcm_column is not None and hcm_column.end_row >= hcm_column.start_row:
            for row in range(hcm_column.start_row, hcm_column.end_row + 1):
                target = horizontal_anchors.get((row, hcm_column.column), (row, hcm_column.column))
                if target in protected:
                    continue
                cell = worksheet.cell(target[0], target[1])
                hcm_snapshots.append((target[0], target[1], cell.value, cell.data_type))

        # Chuẩn hóa Number Format của TOÀN BỘ sheet đích về General (tương đương Ctrl+A -> General).
        # Chỉ thay đổi number_format, giữ nguyên tuyệt đối giá trị, kiểu dữ liệu, tiêu đề, style,
        # độ rộng cột, chiều cao dòng, merge cells, công thức và các sheet khác.
        sheet_snapshots: list[tuple[int, int, object, str]] = []
        for row in worksheet.iter_rows():
            for cell in row:
                sheet_snapshots.append((cell.row, cell.column, cell.value, cell.data_type))
                cell.number_format = "General"

        # Đồng bộ number_format mức cột/dòng nếu mẫu đích đặt sẵn định dạng Text '@' ở mức cột/dòng.
        for col_dim in worksheet.column_dimensions.values():
            if col_dim.number_format and col_dim.number_format != "General":
                col_dim.number_format = "General"
        for row_dim in worksheet.row_dimensions.values():
            if row_dim.number_format and row_dim.number_format != "General":
                row_dim.number_format = "General"

        workbook.save(path)
        return written, written_cells, hcm_snapshots, sheet_snapshots
    finally:
        workbook.close()


def _values_match(
    actual_value: object,
    actual_type: str,
    expected_value: object,
    expected_type: str,
) -> bool:
    """So khớp giá trị và kiểu dữ liệu thật (chấp nhận 9 vs 9.0 sau khi Excel COM lưu XML)."""
    if expected_value is None:
        return actual_value is None
    if is_number(expected_value):
        return actual_type == "n" and is_number(actual_value) and float(actual_value) == float(expected_value)
    if expected_type == "f":
        return actual_type == "f" and actual_value == expected_value
    return (
        actual_type == expected_type
        and type(actual_value) is type(expected_value)
        and actual_value == expected_value
    )


def _is_formula(cell: Cell) -> bool:
    return cell.data_type == "f"
