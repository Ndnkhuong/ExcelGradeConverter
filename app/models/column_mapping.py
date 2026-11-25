"""Cấu hình mapping cột điểm."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ColumnMappingItem:
    source_key: str
    dest_key: str
    source_header: str
    dest_header: str


@dataclass
class MappingProfile:
    name: str
    source_name_header: str
    dest_name_header: str
    items: list[ColumnMappingItem] = field(default_factory=list)
    source_header_row: int | None = None
    dest_header_row: int | None = None
    source_data_start: int | None = None
    dest_data_start: int | None = None

    def to_dict(self) -> dict:
        return {
            "source_name_header": self.source_name_header,
            "dest_name_header": self.dest_name_header,
            "source_header_row": self.source_header_row,
            "dest_header_row": self.dest_header_row,
            "source_data_start": self.source_data_start,
            "dest_data_start": self.dest_data_start,
            "items": [
                {
                    "source_key": item.source_key,
                    "dest_key": item.dest_key,
                    "source_header": item.source_header,
                    "dest_header": item.dest_header,
                }
                for item in self.items
            ],
        }

    @classmethod
    def from_dict(cls, name: str, data: dict) -> "MappingProfile":
        items = [
            ColumnMappingItem(
                source_key=str(item.get("source_key", "")),
                dest_key=str(item.get("dest_key", "")),
                source_header=str(item.get("source_header", "")),
                dest_header=str(item.get("dest_header", "")),
            )
            for item in data.get("items", [])
        ]
        return cls(
            name=name,
            source_name_header=str(data.get("source_name_header", "")),
            dest_name_header=str(data.get("dest_name_header", "")),
            items=items,
            source_header_row=data.get("source_header_row"),
            dest_header_row=data.get("dest_header_row"),
            source_data_start=data.get("source_data_start"),
            dest_data_start=data.get("dest_data_start"),
        )
