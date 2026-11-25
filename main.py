"""Điểm khởi chạy Excel Grade Converter."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.utils.file_utils import cleanup_temp
from app.utils.logger import setup_logging


def main() -> None:
    setup_logging()
    cleanup_temp()
    from app.ui.main_window import run

    try:
        run()
    finally:
        cleanup_temp()


if __name__ == "__main__":
    main()
