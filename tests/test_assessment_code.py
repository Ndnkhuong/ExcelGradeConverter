"""Kiểm thử tính năng tự động chuyển mã đánh giá T, H, C và lưu/tải cấu hình mapping."""

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
from app.database.database import Database
from app.database.mapping_repository import MappingRepository
from app.database.settings_store import SettingsStore
from app.models.column_mapping import (
    COLUMN_TYPE_ASSESSMENT,
    COLUMN_TYPE_NUMBER,
    ColumnMappingItem,
    MappingProfile,
)
from app.models.session import ConversionSession
from app.utils.validators import convert_assessment_code


class TestConvertAssessmentCode(unittest.TestCase):
    def test_t_variants_when_enabled(self) -> None:
        self.assertEqual(convert_assessment_code("T", True), "Hoàn thành tốt")
        self.assertEqual(convert_assessment_code("t", True), "Hoàn thành tốt")
        self.assertEqual(convert_assessment_code(" T ", True), "Hoàn thành tốt")
        self.assertEqual(convert_assessment_code("\u00a0t\t", True), "Hoàn thành tốt")

    def test_h_variants_when_enabled(self) -> None:
        self.assertEqual(convert_assessment_code("H", True), "Hoàn thành")
        self.assertEqual(convert_assessment_code("h", True), "Hoàn thành")
        self.assertEqual(convert_assessment_code(" h ", True), "Hoàn thành")

    def test_c_variants_when_enabled(self) -> None:
        self.assertEqual(convert_assessment_code("C", True), "Chưa hoàn thành")
        self.assertEqual(convert_assessment_code("c", True), "Chưa hoàn thành")
        self.assertEqual(convert_assessment_code(" C ", True), "Chưa hoàn thành")

    def test_disabled_keeps_original_values(self) -> None:
        self.assertEqual(convert_assessment_code("T", False), "T")
        self.assertEqual(convert_assessment_code("H", False), "H")
        self.assertEqual(convert_assessment_code("C", False), "C")
        self.assertEqual(convert_assessment_code(" t ", False), " t ")

    def test_non_single_character_and_other_texts_unchanged(self) -> None:
        for text in ("TH", "TC", "Tốt", "Hoàn thành", "Chưa hoàn thành", "TT", "HC", "Đạt"):
            self.assertEqual(convert_assessment_code(text, True), text)

    def test_blank_and_none_unchanged(self) -> None:
        self.assertEqual(convert_assessment_code("", True), "")
        self.assertEqual(convert_assessment_code("   ", True), "   ")
        self.assertIsNone(convert_assessment_code(None, True))

    def test_numeric_values_not_converted(self) -> None:
        self.assertEqual(convert_assessment_code(8, True), 8)
        self.assertIs(type(convert_assessment_code(8, True)), int)
        self.assertEqual(convert_assessment_code(8.5, True), 8.5)
        self.assertEqual(convert_assessment_code(0, True), 0)
        self.assertEqual(convert_assessment_code("8.5", True), "8.5")
        self.assertEqual(convert_assessment_code("8,5", True), "8,5")


class TestAssessmentPipelineAndPersistence(unittest.TestCase):
    def setUp(self) -> None:
        patcher = patch("app.core.excel_writer.normalize_sheet_with_excel_com")
        patcher.start()
        self.addCleanup(patcher.stop)

    def _build_files(self, tmp_dir: Path) -> tuple[Path, Path]:
        file_a = tmp_dir / "source_assess.xlsx"
        file_b = tmp_dir / "dest_assess.xlsx"

        wb_a = Workbook()
        ws_a = wb_a.active
        ws_a.title = "Sheet1"
        ws_a.append(["Họ và tên", "Điểm GK", "Mức đạt GK", "Nhận xét GK"])
        ws_a.append(["Nguyễn Văn A", 9.0, " T ", "Tốt"])
        ws_a.append(["Trần Thị B", "8,5", "h", "TH"])
        ws_a.append(["Lê Văn C", 6.5, "c", "TC"])
        ws_a.append(["Phạm Thị D", 7.0, "", "Hoàn thành"])
        ws_a.append(["Hoàng Văn E", "T", "Hoàn thành", "Đọc tốt"])  # 'T' nằm ở cột Điểm số (Number) -> không chuyển!
        wb_a.save(file_a)

        wb_b = Workbook()
        ws_b = wb_b.active
        ws_b.title = "Sheet1"
        ws_b.append(["Họ và tên", "Điểm GK", "Mức đạt GK", "Nhận xét GK"])
        for name in ("Nguyễn Văn A", "Trần Thị B", "Lê Văn C", "Phạm Thị D", "Hoàng Văn E"):
            ws_b.append([name, None, None, None])
        wb_b.save(file_b)

        return file_a, file_b

    def test_preview_and_export_enabled_vs_disabled_and_number_column_isolated(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                file_a, file_b = self._build_files(tmp_dir)
                engine = ConversionEngine(SettingsStore(tmp_dir / "settings.json"))

                mappings = [
                    ColumnMappingItem("2", "2", "Điểm GK", "Điểm GK", COLUMN_TYPE_NUMBER),
                    ColumnMappingItem("3", "3", "Mức đạt GK", "Mức đạt GK", COLUMN_TYPE_ASSESSMENT),
                    ColumnMappingItem("4", "4", "Nhận xét GK", "Nhận xét GK", COLUMN_TYPE_ASSESSMENT),
                ]

                # 1. Bật (mặc định = True): Preview hiển thị giá trị đã chuyển đổi & export ghi đúng.
                session_on = ConversionSession(
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
                    mappings=mappings,
                )
                self.assertTrue(session_on.auto_convert_assessment)
                _, _, _, review_on = engine.review(session_on)
                preview_map = {(c.dest_row, c.dest_column): c.source_value for c in review_on.cells}
                self.assertEqual(preview_map[(2, 2)], 9.0)
                self.assertEqual(preview_map[(2, 3)], "Hoàn thành tốt")
                self.assertEqual(preview_map[(3, 2)], 8.5)
                self.assertEqual(preview_map[(3, 3)], "Hoàn thành")
                self.assertEqual(preview_map[(4, 3)], "Chưa hoàn thành")
                self.assertIsNone(preview_map[(5, 3)])  # Ô trống không chuyển thành nội dung đánh giá.
                self.assertEqual(preview_map[(2, 4)], "Tốt")
                self.assertEqual(preview_map[(3, 4)], "TH")
                self.assertEqual(preview_map[(4, 4)], "TC")
                self.assertEqual(preview_map[(5, 4)], "Hoàn thành")
                # Cột Điểm GK (COLUMN_TYPE_NUMBER) KHÔNG áp dụng chuyển mã đánh giá ngay cả khi có chữ 'T'.
                self.assertEqual(preview_map[(6, 2)], "T")

                out_on = tmp_dir / "out_on.xlsx"
                engine.export(session_on, out_on, acknowledge_warnings=True)
                wb_on = load_workbook(out_on)
                ws_on = wb_on["Sheet1"]
                self.assertEqual(ws_on["C2"].value, "Hoàn thành tốt")
                self.assertEqual(ws_on["C3"].value, "Hoàn thành")
                self.assertEqual(ws_on["C4"].value, "Chưa hoàn thành")
                self.assertIsNone(ws_on["C5"].value)
                self.assertEqual(ws_on["B6"].value, "T")
                self.assertEqual(ws_on["D2"].value, "Tốt")
                self.assertEqual(ws_on["D3"].value, "TH")
                self.assertEqual(ws_on["D4"].value, "TC")
                wb_on.close()

                # 2. Tắt (False): Giữ nguyên T, H, C trong cả Preview lẫn Export.
                session_off = ConversionSession(
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
                    mappings=mappings,
                    auto_convert_assessment=False,
                )
                _, _, _, review_off = engine.review(session_off)
                preview_off = {(c.dest_row, c.dest_column): c.source_value for c in review_off.cells}
                self.assertEqual(preview_off[(2, 3)], " T ")
                self.assertEqual(preview_off[(3, 3)], "h")
                self.assertEqual(preview_off[(4, 3)], "c")

                out_off = tmp_dir / "out_off.xlsx"
                engine.export(session_off, out_off, acknowledge_warnings=True)
                wb_off = load_workbook(out_off)
                ws_off = wb_off["Sheet1"]
                self.assertEqual(ws_off["C2"].value, " T ")
                self.assertEqual(ws_off["C3"].value, "h")
                self.assertEqual(ws_off["C4"].value, "c")
                wb_off.close()

    def test_mapping_profile_saves_and_restores_auto_convert_state(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            db = Database(tmp_dir / "app.db")
            repo = MappingRepository(db)

            profile_off = MappingProfile(
                name="Mau_Tat",
                source_name_header="Họ và tên",
                dest_name_header="Họ và tên",
                items=[
                    ColumnMappingItem("2", "2", "Điểm GK", "Điểm GK", COLUMN_TYPE_NUMBER),
                    ColumnMappingItem("3", "3", "Mức đạt GK", "Mức đạt GK", COLUMN_TYPE_ASSESSMENT),
                ],
                auto_convert_assessment=False,
            )
            profile_on = MappingProfile(
                name="Mau_Bat",
                source_name_header="Họ và tên",
                dest_name_header="Họ và tên",
                items=[
                    ColumnMappingItem("3", "3", "Mức đạt GK", "Mức đạt GK", COLUMN_TYPE_ASSESSMENT),
                ],
                auto_convert_assessment=True,
            )
            repo.save(profile_off)
            repo.save(profile_on)

            loaded_off = repo.get("Mau_Tat")
            loaded_on = repo.get("Mau_Bat")
            assert loaded_off is not None and loaded_on is not None
            self.assertFalse(loaded_off.auto_convert_assessment)
            self.assertEqual(loaded_off.items[0].column_type, COLUMN_TYPE_NUMBER)
            self.assertEqual(loaded_off.items[1].column_type, COLUMN_TYPE_ASSESSMENT)
            self.assertTrue(loaded_on.auto_convert_assessment)


if __name__ == "__main__":
    unittest.main()
