"""Đối chiếu học sinh chỉ bằng họ tên đã chuẩn hóa.

Không dùng mã học sinh, số thứ tự hay độ giống gần đúng.
Tên trùng không được ghép tự động.
"""

from __future__ import annotations

from collections import defaultdict

from app.core.errors import AppError
from app.models.student_record import (
    AmbiguousGroup,
    MatchPair,
    MatchResult,
    MatchStatus,
    StudentRecord,
)


def match_students(
    source: list[StudentRecord],
    dest: list[StudentRecord],
    manual_pairs: list[tuple[int, int]] | None = None,
    blocked_pairs: list[tuple[int, int]] | None = None,
) -> MatchResult:
    manual = list(manual_pairs or [])
    blocked = set(map(tuple, blocked_pairs or []))
    source_by_row = {student.row_index: student for student in source}
    dest_by_row = {student.row_index: student for student in dest}
    if len(source_by_row) != len(source) or len(dest_by_row) != len(dest):
        raise AppError("Có hai học sinh cùng một dòng trong một file. Không thể đối chiếu an toàn.")

    used_source: set[int] = set()
    used_dest: set[int] = set()
    matched: list[MatchPair] = []
    seen_manual: set[tuple[int, int]] = set()

    for source_row, dest_row in manual:
        if source_row not in source_by_row or dest_row not in dest_by_row:
            raise AppError("Ghép thủ công đang trỏ tới dòng không còn trong danh sách học sinh.")
        if source_row in used_source or dest_row in used_dest or (source_row, dest_row) in seen_manual:
            raise AppError("Một học sinh đang được ghép với nhiều học sinh khác.")
        seen_manual.add((source_row, dest_row))
        used_source.add(source_row)
        used_dest.add(dest_row)
        matched.append(
            MatchPair(
                source_by_row[source_row],
                dest_by_row[dest_row],
                MatchStatus.MANUAL,
                "Người dùng ghép thủ công",
            )
        )

    source_groups: dict[str, list[StudentRecord]] = defaultdict(list)
    dest_groups: dict[str, list[StudentRecord]] = defaultdict(list)
    for student in source:
        if student.row_index in used_source or not student.name_key:
            continue
        source_groups[student.name_key].append(student)
    for student in dest:
        if student.row_index in used_dest or not student.name_key:
            continue
        dest_groups[student.name_key].append(student)

    unmatched_source: list[StudentRecord] = []
    dest_only: list[StudentRecord] = []
    ambiguous: list[AmbiguousGroup] = []

    for key in sorted(set(source_groups) | set(dest_groups)):
        sources = source_groups.get(key, [])
        dests = dest_groups.get(key, [])
        if len(sources) == 1 and len(dests) == 1:
            pair = (sources[0].row_index, dests[0].row_index)
            if pair in blocked:
                unmatched_source.append(sources[0])
                dest_only.append(dests[0])
            else:
                matched.append(
                    MatchPair(
                        sources[0],
                        dests[0],
                        MatchStatus.AUTO,
                        "Khớp duy nhất theo họ tên",
                    )
                )
            continue
        if sources and dests:
            display = sources[0].name_raw if sources else dests[0].name_raw
            ambiguous.append(AmbiguousGroup(key, display, list(sources), list(dests)))
            continue
        unmatched_source.extend(sources)
        dest_only.extend(dests)

    matched.sort(key=lambda pair: pair.dest.row_index)
    return MatchResult(matched, unmatched_source, dest_only, ambiguous)
