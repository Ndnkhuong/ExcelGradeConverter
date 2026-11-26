"""Kiểm thử xử lý cột Mã HCM khi xuất file: chuẩn hóa General và giữ nguyên giá trị chuỗi."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.core.conversion_engine import ConversionEngine
from app.core.errors import AppError
from app.core.excel_writer import verify_exported_cells
from app.core.template_detector import find_hcm_code_column
from app.database.settings_store import SettingsStore
from app.models.column_mapping import COLUMN_TYPE_ASSESSMENT, COLUMN_TYPE_NUMBER, ColumnMappingItem
from app.models.excel_file import ColumnInfo
from app.models.session import ConversionSession


class TestHcmCodeColumn(unittest.TestCase):
    def test_find_hcm_code_column_variants(self) -> None:
        cols = [
            ColumnInfo(1, "A", "STT", "1"),
            ColumnInfo(2, "B", "  Mã   HCM ", "2"),
            ColumnInfo(3, "C", "Họ và tên", "3"),
        ]
        found = find_hcm_code_column(cols)
        self.assertIsNotNone(found)
        assert found is not None
        self.assertEqual(found.index, 2)

        cols_no_hcm = [
            ColumnInfo(1, "A", "STT", "1"),
            ColumnInfo(2, "B", "Họ và tên", "2"),
        ]
        self.assertIsNone(find_hcm_code_column(cols_no_hcm))

    def test_hcm_column_normalized_to_general_without_changing_values(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "src_hcm.xlsx"
                file_b = tmp_dir / "dst_hcm.xlsx"
                output = tmp_dir / "out_hcm.xlsx"

                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Sheet1"
                ws_a.append(["Họ và tên", "Điểm GK", "Mức đạt GK"])
                ws_a.append(["Nguyễn Văn An", "8,5", "T"])
                ws_a.append(["Trần Thị Bình", "9", "H"])
                ws_a.append(["Lê Minh Châu", "7.5", "C"])
                wb_a.save(file_a)

                # Mẫu đích có cột Mã HCM (đang format '@'), cột Mã lớp (ngoài phạm vi, format '@'),
                # một học sinh chỉ có ở đích (vẫn phải chuẩn hóa Mã HCM), và 1 dòng tổng kết bên dưới (không đụng tới).
                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "Sheet1"
                ws_b.append(["Mã HCM", "Mã lớp", "Họ và tên", "Điểm GK", "Mức đạt GK"])
                ws_b["A1"].number_format = "@"  # Tiêu đề cột Mã HCM -> không đổi.

                dest_rows = [
                    ("HSSSO21803640947", "10A1", "Nguyễn Văn An"),
                    ("HSSSO2132749202", "10A1", "Trần Thị Bình"),
                    ("000123456789", "10A1", "Lê Minh Châu"),
                    ("HSSSO000999888", "10A1", "Học Sinh Chỉ Có Ở Đích"),
                ]
                for idx, (hcm, cls_code, name) in enumerate(dest_rows, start=2):
                    c_hcm = ws_b.cell(idx, 1, hcm)
                    c_hcm.number_format = "@"
                    c_cls = ws_b.cell(idx, 2, cls_code)
                    c_cls.number_format = "@"
                    ws_b.cell(idx, 3, name)
                    ws_b.cell(idx, 4, None).number_format = "@"
                    ws_b.cell(idx, 5, None).number_format = "@"

                # Dòng ghi chú cuối sheet (thuộc vùng dữ liệu đến dòng cuối có dữ liệu).
                footer_cell = ws_b.cell(6, 1, "TONG_CONG")
                footer_cell.number_format = "@"

                # Sheet khác trong cùng workbook không được thay đổi cấu trúc hay format.
                other_ws = wb_b.create_sheet("SheetKhac")
                other_ws["A1"] = "GIU_NGUYEN"
                other_ws["A1"].number_format = "@"
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
                    dest_name_key="3",
                    auto_convert_assessment=True,
                    mappings=[
                        ColumnMappingItem("2", "4", "Điểm GK", "Điểm GK", COLUMN_TYPE_NUMBER),
                        ColumnMappingItem("3", "5", "Mức đạt GK", "Mức đạt GK", COLUMN_TYPE_ASSESSMENT),
                    ],
                )

                _, _, _, review = engine.review(session)
                self.assertFalse(any(w.code == "MISSING_HCM_CODE_COLUMN" for w in review.warnings))

                result = engine.export(session, output, acknowledge_warnings=False)
                self.assertTrue(result.hcm_column_found)
                self.assertEqual(result.hcm_cells_formatted, 4)

                out_wb = load_workbook(output, data_only=False)
                out_ws = out_wb["Sheet1"]

                # Tất cả ô dữ liệu cột Mã HCM (dòng 2..5) có General và giữ nguyên chính xác giá trị chuỗi!
                expected_hcm = [
                    (2, "HSSSO21803640947"),
                    (3, "HSSSO2132749202"),
                    (4, "000123456789"),
                    (5, "HSSSO000999888"),
                ]
                for row_idx, expected_code in expected_hcm:
                    cell = out_ws.cell(row_idx, 1)
                    self.assertEqual(cell.number_format, "General")
                    self.assertEqual(cell.data_type, "s")
                    self.assertIs(type(cell.value), str)
                    self.assertEqual(cell.value, expected_code)

                # Toàn bộ ô trên sheet đích (kể cả dòng tiêu đề 1 và dữ liệu 2..6) đều có General!
                for row in out_ws.iter_rows():
                    for cell in row:
                        self.assertEqual(cell.number_format, "General")

                # Nội dung tiêu đề giữ nguyên; SheetKhac không bị thay đổi; Mã lớp '10A1' và 'TONG_CONG' vẫn giữ nguyên chuỗi.
                self.assertEqual(out_ws["A1"].value, "Mã HCM")
                self.assertEqual(out_ws["B2"].value, "10A1")
                self.assertEqual(out_ws["B2"].data_type, "s")
                self.assertEqual(out_ws["A6"].value, "TONG_CONG")
                self.assertEqual(out_ws["A6"].data_type, "s")
                self.assertIn("SheetKhac", out_wb.sheetnames)
                self.assertEqual(out_wb["SheetKhac"]["A1"].value, "GIU_NGUYEN")
                self.assertEqual(out_wb["SheetKhac"]["A1"].number_format, "@")

                # Các cột đã map vẫn đúng kiểu & General.
                self.assertEqual(out_ws["D2"].value, 8.5)
                self.assertEqual(out_ws["D2"].data_type, "n")
                self.assertEqual(out_ws["D2"].number_format, "General")
                self.assertEqual(out_ws["E2"].value, "Hoàn thành tốt")
                self.assertEqual(out_ws["E2"].data_type, "s")
                self.assertEqual(out_ws["E2"].number_format, "General")
                out_wb.close()

    def test_dynamic_data_region_custom_header_and_start_row(self) -> None:
        """Kiểm tra vùng dữ liệu xác định động (tiêu đề dòng 4, dữ liệu dòng 5..7) không hardcode."""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "src_dyn.xlsx"
                file_b = tmp_dir / "dst_dyn.xlsx"
                output = tmp_dir / "out_dyn.xlsx"

                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Nguon"
                ws_a.append(["Họ và tên", "Điểm", "Nhận xét"])
                ws_a.append(["Nguyễn Văn A", "9,25", "Em đọc bài tốt, chữ viết đều, đẹp."])
                ws_a.append(["Trần Thị B", "8", "Tiến bộ rõ rệt."])
                wb_a.save(file_a)

                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "DichCustom"
                ws_b["A1"] = "TRƯỜNG TIỂU HỌC"
                ws_b["A1"].number_format = "@"
                ws_b.cell(4, 1, "Mã HCM").number_format = "@"
                ws_b.cell(4, 2, "Họ và tên").number_format = "@"
                ws_b.cell(4, 3, "Lớp").number_format = "@"
                ws_b.cell(4, 4, "Điểm").number_format = "@"
                ws_b.cell(4, 5, "Nhận xét").number_format = "@"
                ws_b.cell(4, 6, "Ghi chú thêm").number_format = "@"

                rows_b = [
                    ("HSSSO21803640947", "Nguyễn Văn A", "01A", None, None, "0099"),
                    ("HSSSO2132749202", "Trần Thị B", "01A", None, None, "0100"),
                ]
                for offset, vals in enumerate(rows_b):
                    r = 5 + offset
                    for c, val in enumerate(vals, start=1):
                        cell = ws_b.cell(r, c, val)
                        cell.number_format = "@"
                wb_b.save(file_b)

                engine = ConversionEngine(SettingsStore(tmp_dir / "settings.json"))
                session = ConversionSession(
                    source_path=str(file_a),
                    dest_path=str(file_b),
                    source_sheet="Nguon",
                    dest_sheet="DichCustom",
                    source_header_row=1,
                    source_data_start=2,
                    dest_header_row=4,
                    dest_data_start=5,
                    source_name_key="1",
                    dest_name_key="2",
                    mappings=[
                        ColumnMappingItem("2", "4", "Điểm", "Điểm", COLUMN_TYPE_NUMBER),
                        ColumnMappingItem("3", "5", "Nhận xét", "Nhận xét", COLUMN_TYPE_ASSESSMENT),
                    ],
                )
                result = engine.export(session, output, acknowledge_warnings=True)
                self.assertEqual(result.region_cells_formatted, 36)  # Toàn bộ sheet đích: 6 dòng (1..6) × 6 cột (A..F).

                out_wb = load_workbook(output, data_only=False)
                out_ws = out_wb["DichCustom"]
                # Toàn bộ sheet đích (tiêu đề bảng dòng 1, tiêu đề cột dòng 4 và dữ liệu dòng 5..6) đều có General.
                for row in out_ws.iter_rows():
                    for cell in row:
                        self.assertEqual(cell.number_format, "General")
                self.assertEqual(out_ws["A1"].value, "TRƯỜNG TIỂU HỌC")
                self.assertEqual(out_ws["A4"].value, "Mã HCM")
                self.assertEqual(out_ws["A5"].value, "HSSSO21803640947")
                self.assertEqual(out_ws["A5"].data_type, "s")
                self.assertEqual(out_ws["C5"].value, "01A")
                self.assertEqual(out_ws["C5"].data_type, "s")
                self.assertEqual(out_ws["D5"].value, 9.25)
                self.assertEqual(out_ws["D5"].data_type, "n")
                self.assertEqual(out_ws["E5"].value, "Em đọc bài tốt, chữ viết đều, đẹp.")
                self.assertEqual(out_ws["E5"].data_type, "s")
                self.assertEqual(out_ws["F5"].value, "0099")
                self.assertEqual(out_ws["F5"].data_type, "s")
                out_wb.close()

    def test_missing_hcm_column_shows_warning_and_blocks_until_acknowledged(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a = tmp_dir / "src_no_hcm.xlsx"
                file_b = tmp_dir / "dst_no_hcm.xlsx"
                output = tmp_dir / "out_no_hcm.xlsx"

                wb_a = Workbook()
                ws_a = wb_a.active
                ws_a.title = "Sheet1"
                ws_a.append(["Họ và tên", "Điểm GK"])
                ws_a.append(["Nguyễn Văn An", 9.0])
                wb_a.save(file_a)

                wb_b = Workbook()
                ws_b = wb_b.active
                ws_b.title = "Sheet1"
                ws_b.append(["Họ và tên", "Điểm GK"])
                ws_b.append(["Nguyễn Văn An", None])
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
                    mappings=[ColumnMappingItem("2", "2", "Điểm GK", "Điểm GK", COLUMN_TYPE_NUMBER)],
                )

                _, _, _, review = engine.review(session)
                self.assertTrue(any(w.code == "MISSING_HCM_CODE_COLUMN" for w in review.warnings))

                # Nếu chưa xác nhận cảnh báo -> nhắc người dùng kiểm tra trước khi xuất.
                with self.assertRaises(AppError):
                    engine.export(session, output, acknowledge_warnings=False)

                # Sau khi xác nhận cảnh báo -> vẫn xuất bình thường và báo hcm_column_found=False.
                result = engine.export(session, output, acknowledge_warnings=True)
                self.assertFalse(result.hcm_column_found)
                self.assertEqual(result.hcm_cells_formatted, 0)

    def test_verify_exported_cells_detects_changed_hcm_value_or_format(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad_hcm.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws["A2"] = 12345  # Bị đổi thành int thay vì chuỗi "00012345".
            ws["A2"].number_format = "General"
            wb.save(path)

            with self.assertRaises(AppError):
                verify_exported_cells(path, "Sheet1", [], [(2, 1, "00012345", "s")])


if __name__ == "__main__":
    unittest.main()
