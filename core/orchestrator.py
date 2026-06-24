"""배치 변환 오케스트레이터.

책임:
- 입력 파일 × 선택 출력형식 → 변환 가능 작업 라우팅, 불가 항목은 '건너뜀'.
- 엔진 수명주기 관리(1차 C2: 파일마다 엔진 생성/종료). 같은 파일의 여러 출력형식은
  하나의 엔진(=하나의 한컴/Excel 프로세스)을 공유 → hwp→hwpx+pdf 동시출력이 단일 Open으로 처리됨.
- 파일별 워치독 타임아웃: 초과 시 엔진 프로세스 강제 종료(손상 문서 좀비화 방어).
- 파일 단위 try/except 실패 격리(원칙 1): 예외가 배치 루프를 깨지 않는다.

COM 초기화 불변식: pythoncom.CoInitialize/CoUninitialize는 이 오케스트레이터의 워커 진입/종료에서
'1회'만 호출한다. 엔진 어댑터는 절대 CoInitialize를 호출하지 않는다.
(주의: pyhwpx는 내부적으로 조건부 CoInitialize를 호출하므로 본 프로젝트는 한컴 엔진에서 pyhwpx를 쓰지 않는다.)
"""
from __future__ import annotations

import os
import threading
from itertools import groupby
from typing import Callable, Optional

from .converter import ConversionResult, resolve, engine_for, supported_inputs
from ..utils.filescan import split_name_ext, unique_output_path

DEFAULT_TIMEOUT = 180  # 파일당 워치독 타임아웃(초)


def _noop(*_a, **_k) -> None:
    pass


class BatchSummary:
    def __init__(self) -> None:
        self.success: int = 0
        self.failed: list[tuple[str, str]] = []    # (src, error)
        self.skipped: list[tuple[str, str]] = []   # (src, reason)
        self.deleted: list[str] = []               # 변환 성공 후 삭제된 원본 경로

    @property
    def total(self) -> int:
        return self.success + len(self.failed)


class Orchestrator:
    """엔진 팩토리를 주입받아 배치를 실행한다.

    engine_factories: {"hancom": Callable[[], engine], "excel": Callable[[], engine]}
        engine 객체는 다음 메서드를 제공해야 한다(덕 타이핑):
          - quit(): 정상 종료
          - force_kill(): 프로세스 강제 종료(워치독용)
    """

    def __init__(self, engine_factories: dict[str, Callable[[], object]],
                 timeout: float = DEFAULT_TIMEOUT) -> None:
        self.engine_factories = engine_factories
        self.timeout = timeout

    # ── 작업 계획 ────────────────────────────────────────────────
    def _plan(self, files: list[str], targets: list[str]):
        """(work, skipped) 반환. work=[(index, src, out_ext)...], skipped=[(index, src)...]"""
        targets = [t.lower() for t in targets]
        work: list[tuple[int, str, str]] = []
        skipped: list[tuple[int, str]] = []
        for i, src in enumerate(files):
            _stem, in_ext = split_name_ext(src)
            applicable = [t for t in targets
                          if in_ext in supported_inputs(t) and resolve(in_ext, t)]
            if applicable:
                for t in applicable:
                    work.append((i, src, t))
            else:
                skipped.append((i, src))
        return work, skipped

    # ── 단일 변환 + 워치독 ───────────────────────────────────────
    def _convert_one(self, src: str, out_ext: str, out_path: str, engine,
                     on_log: Callable) -> ConversionResult:
        _stem, in_ext = split_name_ext(src)
        fn = resolve(in_ext, out_ext)
        if fn is None:
            return ConversionResult(False, src, None, "지원하지 않는 변환")

        timed_out = {"v": False}

        def _on_timeout():
            timed_out["v"] = True
            on_log(f"  ⏱ 타임아웃({self.timeout:.0f}s) → 엔진 강제 종료", "error")
            try:
                engine.force_kill()
            except Exception:
                pass

        timer = threading.Timer(self.timeout, _on_timeout)
        timer.start()
        try:
            result = fn(src, out_path, engine)
        except Exception as e:  # noqa: BLE001 - 실패 격리: 모든 예외를 결과로 변환
            if timed_out["v"]:
                result = ConversionResult(False, src, None, f"타임아웃 {self.timeout:.0f}s 초과")
            else:
                result = ConversionResult(False, src, None, str(e))
        finally:
            timer.cancel()
        return result

    # ── 원본 삭제(변환 성공 후 opt-in) ───────────────────────────
    @staticmethod
    def _delete_source(src: str, summary: "BatchSummary", on_log: Callable) -> None:
        """변환에 성공한 원본 파일을 삭제한다. 실패는 격리하여 로그로만 보고."""
        try:
            os.remove(src)
            summary.deleted.append(src)
            on_log(f"  🗑 원본 삭제됨: {os.path.basename(src)}", "warn")
        except Exception as e:  # noqa: BLE001 - 삭제 실패도 배치를 깨지 않는다
            on_log(f"  ⚠ 원본 삭제 실패: {os.path.basename(src)} ({e})", "error")

    # ── 배치 실행 ────────────────────────────────────────────────
    def run(self, files: list[str], targets: list[str], out_dir: Optional[str] = None,
            on_log: Callable = _noop, on_progress: Callable = _noop,
            on_item_result: Callable = _noop, on_done: Callable = _noop,
            init_com: bool = True, delete_originals: bool = False) -> BatchSummary:
        """배치를 실행한다. 보통 UI의 daemon 스레드에서 호출된다.

        init_com=True: 워커 진입 시 pythoncom.CoInitialize 1회, 종료 시 CoUninitialize 1회.
                       (헤드리스 테스트는 init_com=False + 가짜 엔진으로 호출)
        delete_originals=True: 한 파일의 '선택된 모든 출력형식'이 성공한 경우에만 원본을 삭제한다.
                       (데이터 손실 방지: 하나라도 실패하면 보존. 엔진 quit 후 잠금 해제 상태에서 삭제.)
        """
        summary = BatchSummary()
        work, skipped = self._plan(files, targets)

        # 건너뜀 먼저 보고
        for i, src in skipped:
            reason = "선택한 출력형식으로 변환 불가"
            summary.skipped.append((src, reason))
            on_item_result(i, ConversionResult(False, src, None, "건너뜀"))
            on_log(f"⊘ 건너뜀: {os.path.basename(src)} ({reason})", "warn")

        total = len(work)
        on_progress(0, total)
        if total == 0:
            on_done(summary)
            return summary

        pythoncom = None
        if init_com:
            import pythoncom as _pc  # 지연 import: 헤드리스 환경 보호
            pythoncom = _pc
            pythoncom.CoInitialize()

        try:
            done = 0
            # work는 파일 단위로 연속 구성됨 → groupby로 파일당 엔진 1개 생성(C2)
            for src, group in groupby(work, key=lambda w: w[1]):
                items = list(group)
                _stem, in_ext = split_name_ext(src)
                ekind = engine_for(in_ext)
                factory = self.engine_factories.get(ekind)
                if factory is None:
                    for (i, _s, out_ext) in items:
                        res = ConversionResult(False, src, None, f"엔진 없음: {ekind}")
                        summary.failed.append((src, res.error))
                        on_item_result(i, res)
                        on_log(f"  ✗ {os.path.basename(src)} → {out_ext}: {res.error}", "error")
                        done += 1
                        on_progress(done, total)
                    continue

                engine = None
                try:
                    engine = factory()
                except Exception as e:  # noqa: BLE001 - 엔진 생성 실패도 격리
                    for (i, _s, out_ext) in items:
                        res = ConversionResult(False, src, None, f"엔진 생성 실패: {e}")
                        summary.failed.append((src, res.error))
                        on_item_result(i, res)
                        on_log(f"  ✗ {os.path.basename(src)} → {out_ext}: {res.error}", "error")
                        done += 1
                        on_progress(done, total)
                    continue

                src_all_ok = True
                try:
                    out_base_dir = out_dir or os.path.dirname(src)
                    stem, _ext = split_name_ext(src)
                    for (i, _s, out_ext) in items:
                        out_path = unique_output_path(out_base_dir, stem, out_ext)
                        on_log(f"▶ 변환: {os.path.basename(src)} → {out_ext}", "info")
                        res = self._convert_one(src, out_ext, out_path, engine, on_log)
                        if res.ok:
                            summary.success += 1
                            on_log(f"  ✓ 저장됨 → {os.path.basename(res.out or '')}", "success")
                        else:
                            src_all_ok = False
                            summary.failed.append((src, res.error or "알 수 없는 오류"))
                            on_log(f"  ✗ 실패: {res.error}", "error")
                        on_item_result(i, res)
                        done += 1
                        on_progress(done, total)
                finally:
                    try:
                        engine.quit()
                    except Exception:
                        pass

                # 원본 삭제: 이 파일의 모든 출력형식이 성공했고 opt-in일 때만(엔진 종료 후 잠금 해제 상태).
                if delete_originals and src_all_ok:
                    self._delete_source(src, summary, on_log)
        finally:
            if init_com and pythoncom is not None:
                pythoncom.CoUninitialize()

        on_done(summary)
        return summary
