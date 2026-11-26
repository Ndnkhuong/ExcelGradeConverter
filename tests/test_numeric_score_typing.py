"""Kiểm thử tự động chuyển kiểu dữ liệu điểm sang Number của Excel và xác minh sau khi xuất."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.core.conversion_engine import ConversionEngine
from app.core.errors import AppError
from app.core.excel_writer import verify_exported_cells, verify_numeric_cells
from app.database.settings_store import SettingsStore
from app.models.column_mapping import COLUMN_TYPE_ASSESSMENT, COLUMN_TYPE_NUMBER, ColumnMappingItem
from app.models.session import ConversionSession
from app.utils.validators import coerce_cell_value, parse_numeric_value


class TestNumericScoreTyping(unittest.TestCase):
    def setUp(self) -> None:
        patcher = patch("app.core.excel_writer.normalize_sheet_with_excel_com")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_parse_numeric_value_cases(self) -> None:
        # 1 & 2: Chuỗi số nguyên, dấu '.' và ',' -> int/float chính xác.
        self.assertIs(type(parse_numeric_value("9")), int)
        self.assertEqual(parse_numeric_value("9"), 9)
        self.assertIs(type(parse_numeric_value("0")), int)
        self.assertEqual(parse_numeric_value("0"), 0)

        self.assertIs(type(parse_numeric_value("8.5")), float)
        self.assertEqual(parse_numeric_value("8.5"), 8.5)
        self.assertIs(type(parse_numeric_value("8,5")), float)
        self.assertEqual(parse_numeric_value("8,5"), 8.5)
        self.assertEqual(parse_numeric_value(" 7,25 "), 7.25)
        self.assertEqual(parse_numeric_value("9.125"), 9.125)  # 7: Không làm tròn.

        # 5: Không chuyển dữ liệu không phải số thành số.
        for non_numeric in ("H", "T", "Có tiến bộ", "10A1", "01/05/2026", "1.2.3", "nan", "inf", True, False):
            self.assertIsNone(parse_numeric_value(non_numeric))
            self.assertEqual(coerce_cell_value(non_numeric), non_numeric)

        # 6: Không chuyển ô trống thành 0.
        for blank in (None, "", "   "):
            self.assertIsNone(parse_numeric_value(blank))

    def test_export_converts_score_strings_to_excel_number_and_preserves_other_formats(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "source_str_scores.xlsx"
                file_b = tmp_dir / "dest_text_fmt.xlsx"
                output = tmp_dir / "exported.xlsx"

                # File nguồn chứa điểm dạng chuỗi ("8.5", "8,5", "9", "7,25"), ô trống, và cột văn bản.
                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Sheet1"
                ws_a.append(["Họ và tên", "Điểm GK", "Nhận xét"])
                ws_a.append(["Nguyễn Văn A", "8.5", "Tốt"])
                ws_a.append(["Trần Thị B", "8,5", "Khá"])
                ws_a.append(["Lê Văn C", "9", "Giỏi"])
                ws_a.append(["Phạm Thị D", "7,25", "Tiến bộ"])
                ws_a.append(["Hoàng Văn E", "   ", "Chưa có điểm"])
                # Ép các ô điểm nguồn lưu dạng chuỗi Text trong Excel.
                for r in range(2, 7):
                    ws_a.cell(r, 2).number_format = "@"
                wb_a.save(file_a)

                # File đích có cột mã ngoài phạm vi mapping (format "@"), cột điểm (ô r=2..3 là "@", r=4..6 là "0.00").
                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "Sheet1"
                ws_b.append(["Họ và tên", "Mã lớp", "Điểm GK", "Nhận xét"])
                students_b = [
                    ("Nguyễn Văn A", "001", "@"),
                    ("Trần Thị B", "002", "@"),
                    ("Lê Văn C", "003", "0.00"),
                    ("Phạm Thị D", "004", "0.00"),
                    ("Hoàng Văn E", "005", "@"),
                ]
                for idx, (name, code, score_fmt) in enumerate(students_b, start=2):
                    ws_b.cell(idx, 1, name)
                    code_cell = ws_b.cell(idx, 2, code)
                    code_cell.number_format = "@"  # Cột ngoài mapping, phải giữ nguyên "@".
                    score_cell = ws_b.cell(idx, 3, None)
                    score_cell.number_format = score_fmt
                    note_cell = ws_b.cell(idx, 4, None)
                    note_cell.number_format = "@"
                wb_b.save(file_b)

                engine = ConversionEngine(SettingsStore(tmp_dir / "settings.json"))
                session = ConversionSession(
                    source_path=str(file_a),
                    dest_path=str(file_b),
                    source_sheet="Sheet1",
                    dest_sheet="Sheet1",
                    source_header_row=1,
                    source_data_start=2,
                    dest_header_row=1,
                    dest_data_start=2,
                    source_name_key="1",
                    dest_name_key="1",
                    mappings=[
                        ColumnMappingItem("2", "3", "Điểm GK", "Điểm GK"),
                        ColumnMappingItem("3", "4", "Nhận xét", "Nhận xét"),
                    ],
                )

                result = engine.export(session, output, acknowledge_warnings=True)
                # 4 điểm số + 5 nhận xét = 9 ô ghi; 1 ô điểm trống bỏ qua; 4 ô điểm Number đã xác minh.
                self.assertEqual(result.updated_cells, 9)
                self.assertEqual(result.skipped_cells, 1)
                self.assertEqual(result.numeric_cells_verified, 4)

                # Kiểm tra file kết quả bằng openpyxl.
                out_wb = load_workbook(output, data_only=False)
                out_ws = out_wb["Sheet1"]

                # "8.5" -> float 8.5, data_type 'n', đổi '@' thành 'General'.
                c2 = out_ws["C2"]
                self.assertEqual(c2.data_type, "n")
                self.assertIs(type(c2.value), float)
                self.assertEqual(c2.value, 8.5)
                self.assertEqual(c2.number_format, "General")

                # "8,5" -> float 8.5, data_type 'n', đổi '@' thành 'General'.
                c3 = out_ws["C3"]
                self.assertEqual(c3.data_type, "n")
                self.assertIs(type(c3.value), float)
                self.assertEqual(c3.value, 8.5)
                self.assertEqual(c3.number_format, "General")

                # "9" -> int 9, data_type 'n', chuẩn hóa number_format về 'General'.
                c4 = out_ws["C4"]
                self.assertEqual(c4.data_type, "n")
                self.assertIs(type(c4.value), int)
                self.assertEqual(c4.value, 9)
                self.assertEqual(c4.number_format, "General")

                # "7,25" -> float 7.25 (không làm tròn), chuẩn hóa number_format về 'General'.
                c5 = out_ws["C5"]
                self.assertEqual(c5.data_type, "n")
                self.assertIs(type(c5.value), float)
                self.assertEqual(c5.value, 7.25)
                self.assertEqual(c5.number_format, "General")

                # Ô trống không bị chuyển thành 0; thuộc vùng dữ liệu nên có number_format = General.
                c6 = out_ws["C6"]
                self.assertIsNone(c6.value)
                self.assertEqual(c6.number_format, "General")

                # Cột văn bản được map vẫn là chuỗi ('s') và chuẩn hóa number_format về 'General'.
                self.assertEqual(out_ws["D2"].data_type, "s")
                self.assertEqual(out_ws["D2"].value, "Tốt")
                self.assertEqual(out_ws["D2"].number_format, "General")

                # Cột Mã lớp trong vùng dữ liệu được chuẩn hóa General nhưng giữ nguyên chuỗi "001" (không chuyển thành số 1).
                self.assertEqual(out_ws["B2"].value, "001")
                self.assertEqual(out_ws["B2"].data_type, "s")
                self.assertEqual(out_ws["B2"].number_format, "General")
                out_wb.close()

    def test_all_mapped_columns_general_format_and_preserve_visual_styles(self) -> None:
        """Kiểm tra điểm số, mức đánh giá và nhận xét đều có General, đúng kiểu dữ liệu và giữ nguyên style."""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "src_full.xlsx"
                file_b = tmp_dir / "dst_full.xlsx"
                output = tmp_dir / "out_full.xlsx"

                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Sheet1"
                ws_a.append(["Họ và tên", "Điểm GK", "Mức đạt GK", "Nhận xét GK"])
                ws_a.append(["Nguyễn Văn An", "8,5", "T", "Em đọc bài tốt, chữ viết đều, đẹp."])
                wb_a.save(file_a)

                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "Sheet1"
                ws_b.append(["Mã HS", "Họ và tên", "Điểm GK", "Mức đạt GK", "Nhận xét GK"])
                ws_b.append(["HS01", "Nguyễn Văn An", None, None, None])

                # Style & kích thước dòng/cột trên mẫu đích phải được giữ nguyên.
                ws_b.row_dimensions[1].height = 26
                ws_b.row_dimensions[2].height = 30
                ws_b.column_dimensions["E"].width = 40
                header_font = Font(name="Arial", size=11, bold=True)
                ws_b["C1"].font = header_font
                ws_b["C1"].number_format = "@"

                ws_b["A2"].number_format = "@"  # Ngoài vùng map -> phải giữ "@".
                fill = PatternFill(fill_type="solid", start_color="E2EFDA", end_color="E2EFDA")
                thin = Side(style="thin", color="000000")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)
                align = Alignment(horizontal="left", vertical="center", wrap_text=True)
                for col_idx, fmt in ((3, "0.00"), (4, "@"), (5, "@")):
                    c = ws_b.cell(2, col_idx)
                    c.number_format = fmt
                    c.fill = fill
                    c.border = border
                    c.alignment = align
                wb_b.save(file_b)

                engine = ConversionEngine(SettingsStore(tmp_dir / "settings.json"))
                session = ConversionSession(
                    source_path=str(file_a),
                    dest_path=str(file_b),
                    source_sheet="Sheet1",
                    dest_sheet="Sheet1",
                    source_header_row=1,
                    source_data_start=2,
                    dest_header_row=1,
                    dest_data_start=2,
                    source_name_key="1",
                    dest_name_key="2",
                    auto_convert_assessment=True,
                    mappings=[
                        ColumnMappingItem("2", "3", "Điểm GK", "Điểm GK", COLUMN_TYPE_NUMBER),
                        ColumnMappingItem("3", "4", "Mức đạt GK", "Mức đạt GK", COLUMN_TYPE_ASSESSMENT),
                        ColumnMappingItem("4", "5", "Nhận xét GK", "Nhận xét GK", COLUMN_TYPE_ASSESSMENT),
                    ],
                )
                result = engine.export(session, output, acknowledge_warnings=True)
                self.assertEqual(result.updated_cells, 3)
                self.assertEqual(result.numeric_cells_verified, 1)

                out_wb = load_workbook(output, data_only=False)
                out_ws = out_wb["Sheet1"]

                # 1. Điểm số -> float 8.5, data_type 'n', General.
                self.assertEqual(out_ws["C2"].value, 8.5)
                self.assertIs(type(out_ws["C2"].value), float)
                self.assertEqual(out_ws["C2"].data_type, "n")
                self.assertEqual(out_ws["C2"].number_format, "General")

                # 2. Mức đánh giá -> 'Hoàn thành tốt', data_type 's', General.
                self.assertEqual(out_ws["D2"].value, "Hoàn thành tốt")
                self.assertEqual(out_ws["D2"].data_type, "s")
                self.assertEqual(out_ws["D2"].number_format, "General")

                # 3. Nhận xét -> giữ nguyên nội dung, data_type 's', General.
                self.assertEqual(out_ws["E2"].value, "Em đọc bài tốt, chữ viết đều, đẹp.")
                self.assertEqual(out_ws["E2"].data_type, "s")
                self.assertEqual(out_ws["E2"].number_format, "General")

                # Toàn bộ ô trên sheet đích (kể cả dòng tiêu đề) có General;
                # giữ nguyên nội dung tiêu đề, font, màu nền, viền, căn lề, chiều cao dòng & độ rộng cột.
                for row in out_ws.iter_rows():
                    for cell in row:
                        self.assertEqual(cell.number_format, "General")
                self.assertEqual(out_ws["C1"].value, "Điểm GK")
                self.assertTrue(out_ws["C1"].font.bold)
                self.assertEqual(out_ws["A2"].value, "HS01")
                self.assertEqual(out_ws["A2"].data_type, "s")
                self.assertEqual(out_ws["E2"].fill.start_color.rgb[-6:], "E2EFDA")
                self.assertEqual(out_ws["E2"].border.left.style, "thin")
                self.assertTrue(out_ws["E2"].alignment.wrap_text)
                self.assertEqual(out_ws.row_dimensions[1].height, 26)
                self.assertEqual(out_ws.row_dimensions[2].height, 30)
                self.assertEqual(out_ws.column_dimensions["E"].width, 40)
                out_wb.close()

    def test_verify_exported_cells_rejects_non_general_format(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad_fmt.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["B2"] = "Hoàn thành tốt"
            ws["B2"].number_format = "@"
            wb.save(path)

            with self.assertRaises(AppError):
                verify_exported_cells(path, "Sheet1", [(2, 2, "Hoàn thành tốt")])

    def test_out_of_range_numeric_string_is_caught(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "src_bad.xlsx"
                file_b = tmp_dir / "dst_bad.xlsx"
                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Sheet1"
                ws_a.append(["Họ và tên", "Điểm"])
                ws_a.append(["Nguyễn Văn A", "12,5"])
                wb_a.save(file_a)

                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "Sheet1"
                ws_b.append(["Họ và tên", "Điểm"])
                ws_b.append(["Nguyễn Văn A", None])
                wb_b.save(file_b)

                engine = ConversionEngine(SettingsStore(tmp_dir / "settings.json"))
                session = ConversionSession(
                    source_path=str(file_a),
                    dest_path=str(file_b),
                    source_sheet="Sheet1",
                    dest_sheet="Sheet1",
                    source_header_row=1,
                    source_data_start=2,
                    dest_header_row=1,
                    dest_data_start=2,
                    source_name_key="1",
                    dest_name_key="1",
                    mappings=[ColumnMappingItem("2", "2", "Điểm", "Điểm")],
                )
                _, _, _, review = engine.review(session)
                self.assertEqual(len(review.errors), 1)
                self.assertEqual(review.errors[0].code, "SCORE_OUT_OF_RANGE")

    def test_verify_numeric_cells_rejects_text_cell(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad_text_score.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["B2"] = "8.5"  # Chuỗi Text thay vì số.
            wb.save(path)

            with self.assertRaises(AppError):
                verify_numeric_cells(path, "Sheet1", [(2, 2, 8.5)])


if __name__ == "__main__":
    unittest.main()
