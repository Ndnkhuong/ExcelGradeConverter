"""Đọc và ghi cài đặt thang điểm. Không chứa dữ liệu học sinh."""

from __future__ import annotations

import json
from pathlib import Path

from app.database.database import Database
from app.utils.logger import get_logger


class SettingsStore:
    def __init__(self, path: Path, database: Database | None = None):
        self.path = Path(path)
        self.database = database
        self.score_min = 0.0
        self.score_max = 10.0
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            return
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            self.score_min = float(data.get("score_min", 0))
            self.score_max = float(data.get("score_max", 10))
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            get_logger().warning("Không đọc được settings.json, dùng thang điểm 0–10.")

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "score_min": self.score_min,
            "score_max": self.score_max,
        }
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        if self.database is not None:
            self.database.set_setting("score_min", str(self.score_min))
            self.database.set_setting("score_max", str(self.score_max))
