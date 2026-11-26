"""Điều phối đọc, ghép, kiểm tra và xuất. Không thay đổi file gốc."""

from __future__ import annotations

from pathlib import Path

from app.core.column_mapper import suggest_mapping
from app.core.excel_reader import list_sheet_names, read_preview, read_sheet, suggest_header_row
from app.core.excel_writer import CellUpdate, DataRegionSpec, HcmColumnSpec, write_updates_with_report
from app.core.preview_service import build_review
from app.core.student_matcher import match_students
from app.core.template_detector import find_hcm_code_column
from app.database.settings_store import SettingsStore
from app.models.conversion_result import ACTION_UPDATE, ConversionResult, ReviewResult
from app.models.excel_file import CellData, SheetData
from app.models.session import ConversionSession
from app.models.student_record import MatchResult, StudentRecord
from app.utils.logger import get_logger
from app.utils.name_normalizer import display_name, normalize_person_name
from app.utils.validators import coerce_cell_value, is_blank


class ConversionEngine:
    def __init__(self, settings: SettingsStore):
        self.settings = settings

    def list_sheets(self, path: str) -> list[str]:
        return list_sheet_names(path)

    def suggest_header(self, path: str, sheet_name: str) -> int:
        return suggest_header_row(path, sheet_name)

    def preview(self, path: str, sheet_name: str, max_rows: int | None = None) -> list[list[str]]:
        if max_rows is None:
            return read_preview(path, sheet_name)
        return read_preview(path, sheet_name, max_rows=max_rows)

    def load(self, session: ConversionSession) -> tuple[SheetData, SheetData]:
        source = read_sheet(
            session.source_path,
            session.source_sheet,
            session.source_header_row,
            session.source_data_start,
        )
        dest = read_sheet(
            session.dest_path,
            session.dest_sheet,
            session.dest_header_row,
            session.dest_data_start,
        )
        get_logger().info(
            "Đã đọc file nguồn (%s dòng) và file đích (%s dòng)",
            len(source.rows),
            len(dest.rows),
        )
        return source, dest

    def suggest(self, source: SheetData, dest: SheetData):
        return suggest_mapping(source, dest)

    def review(self, session: ConversionSession) -> tuple[SheetData, SheetData, MatchResult, ReviewResult]:
        source, dest = self.load(session)
        source_students = build_students(source, session.source_name_key, _source_keys(session))
        dest_students = build_students(dest, session.dest_name_key, _dest_keys(session))
        match = match_students(
            source_students,
            dest_students,
            session.manual_pairs,
            session.blocked_pairs,
        )
        review = build_review(
            source,
            dest,
            session.source_name_key,
            session.dest_name_key,
            session.mappings,
            match,
            self.settings.score_min,
            self.settings.score_max,
            auto_convert_assessment=session.auto_convert_assessment,
        )
        session.match = match
        session.review = review
        get_logger().info(
            "Đối chiếu xong: %s học sinh khớp, %s trường hợp chưa chuyển điểm",
            match.matched_count,
            match.unmatched_count,
        )
        return source, dest, match, review

    def export(self, session: ConversionSession, output_path: Path, acknowledge_warnings: bool) -> ConversionResult:
        _source, dest, match, review = self.review(session)
        allowed, message = review.can_export(acknowledge_warnings)
        if not allowed:
            from app.core.errors import AppError

            raise AppError(message)
        updates = [
            CellUpdate(cell.dest_row, cell.dest_column, coerce_cell_value(cell.source_value))
            for cell in review.cells
            if cell.action == ACTION_UPDATE and not is_blank(cell.source_value)
        ]
        hcm_column = find_hcm_code_column(dest.columns)
        hcm_spec: HcmColumnSpec | None = None
        if hcm_column is not None:
            dest_students = build_students(dest, session.dest_name_key, [])
            if dest_students:
                end_row = max(student.row_index for student in dest_students)
            elif dest.rows:
                end_row = max(row.row_index for row in dest.rows)
            else:
                end_row = 0
            if end_row >= dest.data_start_row:
                hcm_spec = HcmColumnSpec(
                    column=hcm_column.index,
                    start_row=dest.data_start_row,
                    end_row=end_row,
                )

        # Xác định động toàn bộ vùng dữ liệu của sheet đích theo cấu trúc file mẫu đích.
        data_region: DataRegionSpec | None = None
        row_candidates = [row.row_index for row in dest.rows] + [u.row for u in updates]
        col_candidates = [col.index for col in dest.columns] + [u.column for u in updates]
        if row_candidates and col_candidates:
            data_end_row = max(row_candidates)
            if data_end_row >= dest.data_start_row:
                data_region = DataRegionSpec(
                    start_row=dest.data_start_row,
                    end_row=data_end_row,
                    start_col=min(col_candidates),
                    end_col=max(col_candidates),
                )

        written, verified_numeric, verified_hcm, verified_region = write_updates_with_report(
            session.dest_path,
            output_path,
            session.dest_sheet,
            updates,
            session.source_path,
            hcm_column=hcm_spec,
            data_region=data_region,
        )
        result = ConversionResult(
            output_path=str(output_path),
            matched_count=match.matched_count,
            unmatched_count=match.unmatched_count,
            updated_cells=written,
            skipped_cells=sum(1 for cell in review.cells if cell.action != ACTION_UPDATE),
            mapping_name=session.profile_name,
            numeric_cells_verified=verified_numeric,
            hcm_column_found=hcm_column is not None,
            hcm_cells_formatted=verified_hcm,
            region_cells_formatted=verified_region,
        )
        session.last_result = result
        get_logger().info(
            "Xuất xong: %s học sinh khớp, %s ô được ghi",
            result.matched_count,
            result.updated_cells,
        )
        return result


def build_students(sheet: SheetData, name_key: str, score_keys: list[str]) -> list[StudentRecord]:
    students: list[StudentRecord] = []
    if not name_key:
        return students
    for row in sheet.rows:
        name_cell = row.get(name_key)
        raw = "" if name_cell is None else display_name(name_cell.value)
        if not raw:
            continue
        scores: dict[str, CellData] = {}
        for key in score_keys:
            cell = row.get(key)
            if cell is not None:
                scores[key] = cell
        students.append(
            StudentRecord(
                row_index=row.row_index,
                name_raw=raw,
                name_key=normalize_person_name(raw),
                scores=scores,
            )
        )
    return students


def _source_keys(session: ConversionSession) -> list[str]:
    return [item.source_key for item in session.mappings]


def _dest_keys(session: ConversionSession) -> list[str]:
    return [item.dest_key for item in session.mappings]
