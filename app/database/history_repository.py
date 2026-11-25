"""Lịch sử chuyển đổi. Chỉ lưu tên file và số lượng, không lưu điểm."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from app.database.database import Database


@dataclass
class HistoryEntry:
    id: int
    converted_at: str
    source_name: str
    dest_name: str
    mapping_name: str
    matched_count: int
    unmatched_count: int
    result_path: str


class HistoryRepository:
    def __init__(self, database: Database):
        self.database = database

    def add(
        self,
        source_name: str,
        dest_name: str,
        mapping_name: str,
        matched_count: int,
        unmatched_count: int,
        result_path: str,
    ) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        with self.database.connect() as connection:
            connection.execute(
                """
                INSERT INTO conversion_history (
                    converted_at, source_name, dest_name, mapping_name,
                    matched_count, unmatched_count, result_path
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    now,
                    source_name,
                    dest_name,
                    mapping_name,
                    matched_count,
                    unmatched_count,
                    result_path,
                ),
            )

    def list_recent(self, limit: int = 30) -> list[HistoryEntry]:
        with self.database.connect() as connection:
            rows = connection.execute(
                """
                SELECT id, converted_at, source_name, dest_name, mapping_name,
                       matched_count, unmatched_count, result_path
                FROM conversion_history
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [
            HistoryEntry(
                id=int(row["id"]),
                converted_at=str(row["converted_at"]),
                source_name=str(row["source_name"]),
                dest_name=str(row["dest_name"]),
                mapping_name=str(row["mapping_name"] or ""),
                matched_count=int(row["matched_count"]),
                unmatched_count=int(row["unmatched_count"]),
                result_path=str(row["result_path"] or ""),
            )
            for row in rows
        ]
