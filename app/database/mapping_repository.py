"""Lưu và tải cấu hình mapping. Không lưu nội dung bảng điểm."""

from __future__ import annotations

import json
from datetime import datetime

from app.database.database import Database
from app.models.column_mapping import MappingProfile


class MappingRepository:
    def __init__(self, database: Database):
        self.database = database

    def list_names(self) -> list[str]:
        with self.database.connect() as connection:
            rows = connection.execute(
                "SELECT name FROM mapping_profiles ORDER BY updated_at DESC"
            ).fetchall()
        return [str(row["name"]) for row in rows]

    def get(self, name: str) -> MappingProfile | None:
        with self.database.connect() as connection:
            row = connection.execute(
                "SELECT name, config_json FROM mapping_profiles WHERE name = ?",
                (name,),
            ).fetchone()
        if row is None:
            return None
        data = json.loads(row["config_json"])
        return MappingProfile.from_dict(str(row["name"]), data)

    def save(self, profile: MappingProfile) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        payload = json.dumps(profile.to_dict(), ensure_ascii=False)
        with self.database.connect() as connection:
            connection.execute(
                """
                INSERT INTO mapping_profiles (name, config_json, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE SET
                    config_json = excluded.config_json,
                    updated_at = excluded.updated_at
                """,
                (profile.name, payload, now, now),
            )

    def delete(self, name: str) -> None:
        with self.database.connect() as connection:
            connection.execute("DELETE FROM mapping_profiles WHERE name = ?", (name,))
