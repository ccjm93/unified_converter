"""엔진 팩토리. import 시 COM을 끌어오지 않도록 지연 import한다."""
from __future__ import annotations

from typing import Callable


def make_hancom():
    from .hancom_engine import HancomEngine
    return HancomEngine()


def make_excel():
    from .excel_engine import ExcelEngine
    return ExcelEngine()


DEFAULT_FACTORIES: dict[str, Callable[[], object]] = {
    "hancom": make_hancom,
    "excel": make_excel,
}
