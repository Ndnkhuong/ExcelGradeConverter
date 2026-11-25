"""Điều phối đọc, ghép, kiểm tra và xuất. Không thay đổi file gốc."""

from __future__ import annotations

from pathlib import Path

from app.core.column_mapper import suggest_mapping
from app.core.excel_reader import list_sheet_names, read_preview, read_sheet, suggest_header_row
from app.core.excel_writer import CellUpdate, write_updates
from app.core.preview_service import build_review
from app.core.student_matcher import match_students
from app.database.settings_store import SettingsStore
from app.models.conversion_result import ACTION_UPDATE, ConversionResult, ReviewResult
from app.models.excel_file import CellData, SheetData
from app.models.session import ConversionSession
from app.models.student_record import MatchResult, StudentRecord
from app.utils.logger import get_logger
from app.utils.name_normalizer import display_name, normalize_person_name
from app.utils.validators import is_number


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
        _source, _dest, match, review = self.review(session)
        allowed, message = review.can_export(acknowledge_warnings)
        if not allowed:
            from app.core.errors import AppError

            raise AppError(message)
        updates = [
            CellUpdate(cell.dest_row, cell.dest_column, cell.source_value)
            for cell in review.cells
            if cell.action == ACTION_UPDATE and is_number(cell.source_value)
        ]
        written = write_updates(
            session.dest_path,
            output_path,
            session.dest_sheet,
            updates,
            session.source_path,
        )
        result = ConversionResult(
            output_path=str(output_path),
            matched_count=match.matched_count,
            unmatched_count=match.unmatched_count,
            updated_cells=written,
            skipped_cells=sum(1 for cell in review.cells if cell.action != ACTION_UPDATE),
            mapping_name=session.profile_name,
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
