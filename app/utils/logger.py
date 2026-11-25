"""Ghi log local. Không ghi điểm hay họ tên học sinh."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler

from config import log_path

_CONFIGURED = False


def setup_logging() -> logging.Logger:
    global _CONFIGURED
    logger = logging.getLogger("excel_grade_converter")
    if _CONFIGURED:
        return logger
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = RotatingFileHandler(
        log_path(), maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.propagate = False
    _CONFIGURED = True
    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger("excel_grade_converter")
