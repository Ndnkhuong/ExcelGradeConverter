"""Kiểm thử mapping điểm mặc định: mẫu A (cột E, từ dòng 14) -> mẫu B (cột E, từ dòng 2)."""

from __future__ import annotations

import hashlib
import os
import sys
import tempfile
import unicodedata
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.core.column_mapper import suggest_mapping
from app.core.conversion_engine import ConversionEngine
from app.database.settings_store import SettingsStore
from app.models.column_mapping import ColumnMappingItem
from app.models.session import ConversionSession
from config import (
    DEFAULT_DEST_DATA_START,
    DEFAULT_DEST_HEADER_ROW,
    DEFAULT_DEST_SCORE_COLUMN,
    DEFAULT_SOURCE_DATA_START,
    DEFAULT_SOURCE_HEADER_ROW,
    DEFAULT_SOURCE_SCORE_COLUMN,
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _create_template_a(path: Path) -> None:
    """Mẫu A: thông tin chung dòng 1-12, tiêu đề dòng 12-13 (có merge), học sinh từ dòng 14, điểm ở cột E."""
    wb = Workbook()
    ws = wb.active
    ws.title = "BangDiemA"

    # Dữ liệu tiêu đề phía trên dòng 14 không được tính là học sinh.
    ws["A1"] = "TRƯỜNG THPT MẪU"
    ws["A2"] = "BẢNG ĐIỂM HỌC KỲ"
    ws["B10"] = "Lớp: 10A1"
    ws["E10"] = 99  # Giá trị nằm trên dòng 14, phải bị bỏ qua.

    # Tiêu đề gộp dọc 12:13 để kiểm tra đọc cột khi có merged header.
    ws["A12"] = "STT"
    ws["B12"] = "Họ và tên"
    ws["C12"] = "Lớp"
    ws["D12"] = "Ghi chú"
    ws["E12"] = "Điểm nguồn"
    for col in ("A", "B", "C", "D", "E"):
        ws.merge_cells(f"{col}12:{col}13")

    # Dữ liệu học sinh bắt đầu từ dòng 14 (khác thứ tự với mẫu B).
    nfd_name = unicodedata.normalize("NFD", "Trần   Thị   Bình")
    rows = [
        (1, "  nguyễn văn an ", "10A1", "Ghi chú A1", 8.5),
        (2, nfd_name, "10A1", "Ghi chú A2", 9.25),
        (3, "Lê Minh Châu", "10A1", "Ghi chú A3", 7.0),
        (4, "Phạm Trùng Tên", "10A1", "A4", 6.5),
        (5, "Phạm Trùng Tên", "10A1", "A5", 7.5),
        (6, "Học Sinh Không Có Ở B", "10A1", "A6", 10.0),
    ]
    for offset, row in enumerate(rows):
        row_idx = 14 + offset
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row_idx, col_idx, value)

    wb.save(path)


def _create_template_b(path: Path) -> None:
    """Mẫu B: tiêu đề dòng 1, học sinh từ dòng 2 (thứ tự khác mẫu A), cột E là cột điểm đích."""
    wb = Workbook()
    ws = wb.active
    ws.title = "BangDiemB"

    ws.append(["STT", "Họ và tên", "Lớp", "Ghi chú", "Điểm đích", "Công thức"])
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["E"].width = 14

    students = [
        (1, "Lê Minh Châu", "10A1-B", "Giữ nguyên D2", None, "=E2*2"),
        (2, "Trần Thị Bình", "10A1-B", "Giữ nguyên D3", None, "=E3*2"),
        (3, "Học Sinh Chỉ Có Ở B", "10A1-B", "Giữ nguyên D4", 5.0, "=E4*2"),
        (4, "Nguyễn Văn An", "10A1-B", "Giữ nguyên D5", None, "=E5*2"),
        (5, "Phạm Trùng Tên", "10A1-B", "Giữ nguyên D6", None, "=E6*2"),
    ]

    score_font = Font(name="Times New Roman", size=12, bold=True, color="1F4E79")
    score_fill = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
    thin = Side(style="thin", color="000000")
    score_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    score_align = Alignment(horizontal="center", vertical="center")

    for offset, row in enumerate(students):
        row_idx = 2 + offset
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row_idx, col_idx, value)
            if col_idx == 5:
                cell.font = score_font
                cell.fill = score_fill
                cell.border = score_border
                cell.alignment = score_align
                cell.number_format = "0.00"

    wb.save(path)


class TestDefaultScoreMapping(unittest.TestCase):
    def test_defaults_and_end_to_end_conversion(self) -> None:
        self.assertEqual(DEFAULT_SOURCE_HEADER_ROW, 13)
        self.assertEqual(DEFAULT_SOURCE_DATA_START, 14)
        self.assertEqual(DEFAULT_SOURCE_SCORE_COLUMN, "E")
        self.assertEqual(DEFAULT_DEST_HEADER_ROW, 1)
        self.assertEqual(DEFAULT_DEST_DATA_START, 2)
        self.assertEqual(DEFAULT_DEST_SCORE_COLUMN, "E")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            with patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}):
                self._run_end_to_end(tmp_dir)

    def _run_end_to_end(self, tmp_dir: Path) -> None:
        file_a = tmp_dir / "mau_A.xlsx"
        file_b = tmp_dir / "mau_B.xlsx"
        output = tmp_dir / "ket_qua.xlsx"
        _create_template_a(file_a)
        _create_template_b(file_b)

        hash_a_before = _sha256(file_a)
        hash_b_before = _sha256(file_b)

        settings = SettingsStore(tmp_dir / "settings.json")
        engine = ConversionEngine(settings)

        session = ConversionSession(
            source_path=str(file_a),
            dest_path=str(file_b),
            source_sheet="BangDiemA",
            dest_sheet="BangDiemB",
        )
        self.assertEqual(session.source_header_row, 13)
        self.assertEqual(session.source_data_start, 14)
        self.assertEqual(session.dest_header_row, 1)
        self.assertEqual(session.dest_data_start, 2)

        source, dest = engine.load(session)
        # Dữ liệu mẫu A phải bắt đầu từ dòng 14 (bỏ qua dòng 10).
        self.assertEqual(source.rows[0].row_index, 14)
        self.assertEqual(dest.rows[0].row_index, 2)

        # Mapping mặc định: chỉ ghép cột E -> cột E, không ghép STT/Lớp/Ghi chú.
        suggestion = suggest_mapping(source, dest)
        self.assertEqual(suggestion.source_name_key, "2")
        self.assertEqual(suggestion.dest_name_key, "2")
        self.assertEqual(suggestion.pairs, [("5", "5")])

        session.source_name_key = suggestion.source_name_key or ""
        session.dest_name_key = suggestion.dest_name_key or ""
        session.mappings = [
            ColumnMappingItem(
                source_key="5",
                dest_key="5",
                source_header=source.column_by_key("5").header,
                dest_header=dest.column_by_key("5").header,
            )
        ]

        _, _, match, _ = engine.review(session)
        # 3 học sinh khớp duy nhất; "Phạm Trùng Tên" nằm ở ambiguous, không tự ghép.
        self.assertEqual(match.matched_count, 3)
        self.assertEqual(len(match.ambiguous), 1)
        self.assertEqual(len(match.unmatched_source), 1)
        self.assertEqual(len(match.dest_only), 1)

        result = engine.export(session, output, acknowledge_warnings=True)
        self.assertEqual(result.updated_cells, 3)

        # Kiểm tra file gốc không hề thay đổi.
        self.assertEqual(_sha256(file_a), hash_a_before)
        self.assertEqual(_sha256(file_b), hash_b_before)

        # Kiểm tra file kết quả: đúng dòng, đúng điểm, không ghi đè cột khác, giữ định dạng & công thức.
        out_wb = load_workbook(output)
        out_ws = out_wb["BangDiemB"]

        # Thứ tự học sinh giữ nguyên theo mẫu B và ghi đúng điểm từ cột E mẫu A.
        self.assertEqual(out_ws["B2"].value, "Lê Minh Châu")
        self.assertEqual(out_ws["E2"].value, 7.0)

        self.assertEqual(out_ws["B3"].value, "Trần Thị Bình")
        self.assertEqual(out_ws["E3"].value, 9.25)

        self.assertEqual(out_ws["B4"].value, "Học Sinh Chỉ Có Ở B")
        self.assertEqual(out_ws["E4"].value, 5.0)  # Giữ nguyên điểm cũ vì không có trong A.

        self.assertEqual(out_ws["B5"].value, "Nguyễn Văn An")
        self.assertEqual(out_ws["E5"].value, 8.5)

        self.assertEqual(out_ws["B6"].value, "Phạm Trùng Tên")
        self.assertIsNone(out_ws["E6"].value)  # Trùng tên nên không tự ghép.

        # Các cột khác (STT, Lớp, Ghi chú, Công thức) không bị ghi đè.
        self.assertEqual(out_ws["A2"].value, 1)
        self.assertEqual(out_ws["C2"].value, "10A1-B")
        self.assertEqual(out_ws["D2"].value, "Giữ nguyên D2")
        self.assertEqual(out_ws["F2"].value, "=E2*2")
        self.assertEqual(out_ws["D5"].value, "Giữ nguyên D5")
        self.assertEqual(out_ws["F5"].value, "=E5*2")

        # Định dạng cột E và độ rộng cột được giữ nguyên.
        e2 = out_ws["E2"]
        self.assertEqual(e2.font.name, "Times New Roman")
        self.assertEqual(e2.font.size, 12)
        self.assertTrue(e2.font.bold)
        self.assertEqual(e2.fill.start_color.rgb[-6:], "FFF2CC")
        self.assertEqual(e2.border.left.style, "thin")
        self.assertEqual(e2.alignment.horizontal, "center")
        self.assertEqual(e2.number_format, "0.00")
        self.assertEqual(out_ws.column_dimensions["B"].width, 28)
        self.assertEqual(out_ws.column_dimensions["E"].width, 14)

        out_wb.close()


if __name__ == "__main__":
    unittest.main()
