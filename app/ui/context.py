"""Phụ thuộc dùng chung cho các màn hình."""

from __future__ import annotations

from app.core.conversion_engine import ConversionEngine
from app.database.history_repository import HistoryRepository
from app.database.mapping_repository import MappingRepository
from app.database.settings_store import SettingsStore
from app.models.session import ConversionSession


class AppContext:
    def __init__(
        self,
        session: ConversionSession,
        settings: SettingsStore,
        mapping_repo: MappingRepository,
        history_repo: HistoryRepository,
        engine: ConversionEngine,
    ):
        self.session = session
        self.settings = settings
        self.mapping_repo = mapping_repo
        self.history_repo = history_repo
        self.engine = engine
        self.window = None
