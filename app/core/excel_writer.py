"""Ghi file kết quả bằng cách sao chép file đích rồi chỉ sửa các ô được mapping."""

from __future__ import annotations

import shutil
import uuid
from dataclasses import dataclass
from pathlib import Path

from openpyxl.cell.cell import Cell
from openpyxl.utils import range_boundaries

from app.core.errors import AppError
from app.utils.file_utils import assert_output_is_new, cleanup_temp
from app.utils.logger import get_logger
from config import temp_dir


@dataclass
class CellUpdate:
    row: int
    column: int
    value: object


def write_updates(
    dest_path: str | Path,
    output_path: str | Path,
    sheet_name: str,
    updates: list[CellUpdate],
    source_path: str | Path | None = None,
) -> int:
    """Tạo file mới từ bản sao của file đích. Trả về số ô đã ghi.

    Định dạng, công thức ngoài danh sách cập nhật và các cột khác được giữ nguyên
    vì openpyxl chỉ thay giá trị của các ô được chỉ định trên bản sao.
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
    try:
        shutil.copy2(destination, temp_path)
        written = _apply_updates(temp_path, sheet_name, updates)
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
    get_logger().info("Đã ghi %s ô vào file kết quả %s", written, output.name)
    return written


def _apply_updates(path: Path, sheet_name: str, updates: list[CellUpdate]) -> int:
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
        for update in updates:
            target = horizontal_anchors.get((update.row, update.column), (update.row, update.column))
            if target in protected:
                continue
            cell = worksheet.cell(target[0], target[1])
            if _is_formula(cell):
                continue
            cell.value = update.value
            written += 1
        workbook.save(path)
        return written
    finally:
        workbook.close()


def _is_formula(cell: Cell) -> bool:
    return cell.data_type == "f"
