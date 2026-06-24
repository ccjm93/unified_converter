"""통합 변환기 코어.

core.converters를 import하여 모든 변환 경로를 디스패치 레지스트리에 등록한다.
(이 import는 COM을 끌어오지 않는다 — 헤드리스 안전.)
"""
from . import converters  # noqa: F401  (등록 부수효과)
from .converter import (  # noqa: F401
    ConversionResult,
    engine_for,
    registered_pairs,
    resolve,
    supported_inputs,
)
from .orchestrator import BatchSummary, Orchestrator  # noqa: F401
