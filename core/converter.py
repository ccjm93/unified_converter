"""변환 디스패치 및 결과 자료형.

이 모듈은 COM/tkinter 의존이 전혀 없어 헤드리스로 import·테스트 가능하다(원칙 4).
실제 변환 콜백은 core.converters 패키지가 import 시점에 register()로 등록한다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

# ── 입력 확장자 → 엔진 종류 (입력 확장자가 엔진을 100% 결정) ──────────────
ENGINE_FOR_EXT: dict[str, str] = {
    ".hwp": "hancom",
    ".hwpx": "hancom",
    ".xls": "excel",
    ".xlsx": "excel",
}

# ── 출력 형식별 지원 입력 확장자 (UI 라우팅 / '건너뜀' 판정용) ────────────
SUPPORTED_INPUTS: dict[str, set[str]] = {
    ".hwpx": {".hwp"},
    ".pdf": {".hwp", ".hwpx", ".xls", ".xlsx"},
    ".xlsx": {".xls"},
}


@dataclass
class ConversionResult:
    """단일 (입력파일, 출력형식) 변환의 결과."""

    ok: bool
    src: str
    out: Optional[str] = None      # 성공 시 출력 경로
    error: Optional[str] = None    # 실패/건너뜀 시 메시지


# ── 디스패치 레지스트리: (입력확장자, 출력확장자) → 변환 콜백 ────────────
# 콜백 시그니처: fn(src_path: str, out_path: str, engine) -> ConversionResult
ConverterFn = Callable[[str, str, object], ConversionResult]
_DISPATCH: dict[tuple[str, str], ConverterFn] = {}


def register(in_ext: str, out_ext: str, fn: ConverterFn) -> None:
    """변환 콜백을 등록한다. 신규 경로 추가는 이 호출 1줄이면 충분(원칙 2)."""
    _DISPATCH[(in_ext.lower(), out_ext.lower())] = fn


def resolve(in_ext: str, out_ext: str) -> Optional[ConverterFn]:
    """입력/출력 확장자에 대응하는 변환 콜백을 반환. 없으면 None."""
    return _DISPATCH.get((in_ext.lower(), out_ext.lower()))


def engine_for(in_ext: str) -> Optional[str]:
    """입력 확장자가 필요로 하는 엔진 종류('hancom'|'excel')를 반환."""
    return ENGINE_FOR_EXT.get(in_ext.lower())


def supported_inputs(out_ext: str) -> set[str]:
    """해당 출력 형식으로 변환 가능한 입력 확장자 집합."""
    return SUPPORTED_INPUTS.get(out_ext.lower(), set())


def registered_pairs() -> set[tuple[str, str]]:
    """현재 등록된 (입력, 출력) 변환 경로 집합 (테스트/진단용)."""
    return set(_DISPATCH.keys())
