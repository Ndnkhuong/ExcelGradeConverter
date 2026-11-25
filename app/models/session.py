"""Trạng thái một phiên chuyển đổi trên giao diện."""

from __future__ import annotations

from dataclasses import dataclass, field

from app.models.column_mapping import ColumnMappingItem
from app.models.conversion_result import ConversionResult, ReviewResult
from app.models.student_record import MatchResult
from config import (
    DEFAULT_DEST_DATA_START,
    DEFAULT_DEST_HEADER_ROW,
    DEFAULT_SOURCE_DATA_START,
    DEFAULT_SOURCE_HEADER_ROW,
)


@dataclass
class ConversionSession:
    source_path: str = ""
    dest_path: str = ""
    source_sheet: str = ""
    dest_sheet: str = ""
    source_header_row: int = DEFAULT_SOURCE_HEADER_ROW
    dest_header_row: int = DEFAULT_DEST_HEADER_ROW
    source_data_start: int = DEFAULT_SOURCE_DATA_START
    dest_data_start: int = DEFAULT_DEST_DATA_START
    source_name_key: str = ""
    dest_name_key: str = ""
    mappings: list[ColumnMappingItem] = field(default_factory=list)
    manual_pairs: list[tuple[int, int]] = field(default_factory=list)
    blocked_pairs: list[tuple[int, int]] = field(default_factory=list)
    profile_name: str = ""
    pending_profile_name: str = ""
    acknowledge_warnings: bool = False
    output_dir: str = ""
    output_name: str = ""
    match: MatchResult | None = None
    review: ReviewResult | None = None
    last_result: ConversionResult | None = None

    def reset(self) -> None:
        self.__dict__.update(ConversionSession().__dict__)
