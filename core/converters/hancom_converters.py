"""한컴(HWP/HWPX) 변환 콜백.

이 모듈은 모듈 레벨에서 COM을 import하지 않는다(헤드리스 import 가능).
실제 COM 호출은 주입된 engine(HancomEngine)이 담당한다.
hwp→hwpx 와 hwp→pdf 가 같은 배치에 함께 있으면, 오케스트레이터가 같은 엔진을 공유하고
engine.ensure_open()이 동일 문서를 한 번만 Open하므로 '단일 Open 후 SaveAs 2회'가 자연히 성립한다.
"""
from __future__ import annotations

from ..converter import ConversionResult


def hwp_to_hwpx(src: str, out_path: str, engine) -> ConversionResult:
    engine.ensure_open(src)
    engine.save_as(out_path, "HWPX")
    return ConversionResult(True, src, out_path)


def to_pdf(src: str, out_path: str, engine) -> ConversionResult:
    """HWP/HWPX → PDF (입력 확장자 무관, 한컴 엔진이 동일 처리)."""
    engine.ensure_open(src)
    engine.save_as(out_path, "PDF")
    return ConversionResult(True, src, out_path)
