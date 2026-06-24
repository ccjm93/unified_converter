"""로깅 설정: 회전 파일 핸들러 + UI 텍스트 위젯 핸들러."""
from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler


def log_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "UnifiedConverter", "logs")
    os.makedirs(d, exist_ok=True)
    return d


def setup_file_logging() -> logging.Logger:
    logger = logging.getLogger("unified_converter")
    logger.setLevel(logging.INFO)
    if any(isinstance(h, RotatingFileHandler) for h in logger.handlers):
        return logger
    path = os.path.join(log_dir(), "converter.log")
    handler = RotatingFileHandler(path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger
