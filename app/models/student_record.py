"""Học sinh đọc từ một dòng Excel và kết quả đối chiếu."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from app.models.excel_file import CellData


class MatchStatus(str, Enum):
    AUTO = "auto"
    MANUAL = "manual"
    UNMATCHED_SOURCE = "unmatched_source"
    DEST_ONLY = "dest_only"
    AMBIGUOUS = "ambiguous"


@dataclass
class StudentRecord:
    row_index: int
    name_raw: str
    name_key: str
    scores: dict[str, CellData] = field(default_factory=dict)


@dataclass
class MatchPair:
    source: StudentRecord
    dest: StudentRecord
    status: MatchStatus
    note: str


@dataclass
class AmbiguousGroup:
    name_key: str
    display_name: str
    sources: list[StudentRecord]
    dests: list[StudentRecord]


@dataclass
class MatchResult:
    matched: list[MatchPair]
    unmatched_source: list[StudentRecord]
    dest_only: list[StudentRecord]
    ambiguous: list[AmbiguousGroup]

    @property
    def matched_count(self) -> int:
        return len(self.matched)

    @property
    def unmatched_count(self) -> int:
        ambiguous_sources = sum(len(group.sources) for group in self.ambiguous)
        return len(self.unmatched_source) + ambiguous_sources
