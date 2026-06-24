"""1단계 헤드리스 코어 검증 (COM/tkinter 불필요).

수용 기준:
  - 디스패치 resolve가 6개 변환 경로를 정확히 반환
  - 출력 충돌 시 자동 리네임('name (1).ext')
  - orchestrator가 on_log/on_progress/on_done를 순서대로 호출
  - 변환 불가 입력은 '건너뜀'으로 보고
  - core import 시 tkinter 미import (UI/로직 분리)
"""
from __future__ import annotations

import os
import sys

import pytest

# 프로젝트 루트(unified_converter의 상위)를 path에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from unified_converter.core import converter as C  # noqa: E402
from unified_converter.core.converter import ConversionResult, resolve, supported_inputs  # noqa: E402
from unified_converter.core.orchestrator import Orchestrator  # noqa: E402
from unified_converter.utils.filescan import unique_output_path, split_name_ext  # noqa: E402


# ── 디스패치 ─────────────────────────────────────────────────────
def test_six_conversion_paths_registered():
    pairs = C.registered_pairs()
    expected = {
        (".hwp", ".hwpx"), (".hwp", ".pdf"), (".hwpx", ".pdf"),
        (".xls", ".xlsx"), (".xls", ".pdf"), (".xlsx", ".pdf"),
    }
    assert expected <= pairs


def test_resolve_and_routing():
    assert resolve(".HWP", ".HWPX") is not None      # 대소문자 무관
    assert resolve(".hwpx", ".hwpx") is None          # 미지원 경로
    assert supported_inputs(".pdf") == {".hwp", ".hwpx", ".xls", ".xlsx"}
    assert supported_inputs(".xlsx") == {".xls"}


def test_engine_routing():
    assert C.engine_for(".hwp") == "hancom"
    assert C.engine_for(".xlsx") == "excel"
    assert C.engine_for(".zip") is None


# ── 자동 리네임 ──────────────────────────────────────────────────
def test_unique_output_path_autorename(tmp_path):
    d = str(tmp_path)
    p1 = unique_output_path(d, "doc", ".pdf")
    assert os.path.basename(p1) == "doc.pdf"
    open(p1, "w").close()
    p2 = unique_output_path(d, "doc", ".pdf")
    assert os.path.basename(p2) == "doc (1).pdf"
    open(p2, "w").close()
    p3 = unique_output_path(d, "doc", ".pdf")
    assert os.path.basename(p3) == "doc (2).pdf"


def test_split_name_ext():
    assert split_name_ext(r"C:\a\b\Report.HWP") == ("Report", ".hwp")


# ── 가짜 엔진 + 더미 변환기로 오케스트레이터 검증 ──────────────────
class FakeEngine:
    def __init__(self):
        self.opened = []
        self.killed = False
    def write_out(self, src, out_path):
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(f"converted from {src}")
    def quit(self):
        pass
    def force_kill(self):
        self.killed = True


def _dummy_fn(src, out_path, engine):
    engine.write_out(src, out_path)
    return ConversionResult(True, src, out_path)


@pytest.fixture
def fake_setup(tmp_path):
    # 더미 변환 경로 등록: .txt -> .out
    C.register(".txt", ".out", _dummy_fn)
    src1 = tmp_path / "a.txt"
    src1.write_text("hello")
    src2 = tmp_path / "b.zip"   # 변환 불가(엔진/경로 없음)
    src2.write_text("x")
    return tmp_path, [str(src1), str(src2)]


def test_orchestrator_callbacks_and_skip(fake_setup, monkeypatch):
    tmp_path, files = fake_setup
    # .txt 입력의 엔진 라우팅 추가
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".out", {".txt"})

    logs, progress, items, done = [], [], [], []
    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})

    summary = orch.run(
        files, targets=[".out"], out_dir=str(tmp_path),
        on_log=lambda m, l: logs.append((m, l)),
        on_progress=lambda d, t: progress.append((d, t)),
        on_item_result=lambda i, r: items.append((i, r)),
        on_done=lambda s: done.append(s),
        init_com=False,   # 헤드리스
    )

    # .txt 1건 성공, .zip 1건 건너뜀
    assert summary.success == 1
    assert len(summary.skipped) == 1
    # 진행률은 0/1로 시작해 1/1로 끝
    assert progress[0] == (0, 1)
    assert progress[-1] == (1, 1)
    # on_done 1회
    assert len(done) == 1 and done[0] is summary
    # 출력 파일 생성됨
    assert os.path.exists(os.path.join(str(tmp_path), "a.out"))
    # 건너뜀 로그 존재
    assert any(l == "warn" and "건너뜀" in m for m, l in logs)


def test_orchestrator_autorename_on_collision(fake_setup, monkeypatch):
    tmp_path, files = fake_setup
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".out", {".txt"})
    # 기존 출력 선점
    (tmp_path / "a.out").write_text("preexisting")

    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})
    summary = orch.run([files[0]], targets=[".out"], out_dir=str(tmp_path), init_com=False)

    assert summary.success == 1
    # 기존 a.out 보존 + a (1).out 생성
    assert (tmp_path / "a.out").read_text() == "preexisting"
    assert os.path.exists(os.path.join(str(tmp_path), "a (1).out"))


# ── 변환 후 원본 삭제(opt-in) ─────────────────────────────────────
def _fail_fn(src, out_path, engine):
    return ConversionResult(False, src, None, "강제 실패")


def test_delete_original_on_success(fake_setup, monkeypatch):
    tmp_path, files = fake_setup
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".out", {".txt"})

    src_txt = files[0]  # a.txt
    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})
    summary = orch.run([src_txt], targets=[".out"], out_dir=str(tmp_path),
                       init_com=False, delete_originals=True)

    assert summary.success == 1
    assert not os.path.exists(src_txt)          # 원본 삭제됨
    assert src_txt in summary.deleted
    assert os.path.exists(os.path.join(str(tmp_path), "a.out"))  # 출력은 보존


def test_keep_original_on_failure(fake_setup, monkeypatch):
    tmp_path, files = fake_setup
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".bad", {".txt"})
    C.register(".txt", ".bad", _fail_fn)

    src_txt = files[0]  # a.txt
    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})
    summary = orch.run([src_txt], targets=[".bad"], out_dir=str(tmp_path),
                       init_com=False, delete_originals=True)

    assert len(summary.failed) == 1
    assert os.path.exists(src_txt)              # 실패 → 원본 보존
    assert summary.deleted == []


def test_keep_original_when_one_output_fails(fake_setup, monkeypatch):
    """여러 출력형식 중 하나라도 실패하면 원본을 삭제하지 않는다."""
    tmp_path, files = fake_setup
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".out", {".txt"})
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".bad", {".txt"})
    C.register(".txt", ".bad", _fail_fn)

    src_txt = files[0]  # a.txt
    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})
    summary = orch.run([src_txt], targets=[".out", ".bad"], out_dir=str(tmp_path),
                       init_com=False, delete_originals=True)

    assert summary.success == 1
    assert len(summary.failed) == 1
    assert os.path.exists(src_txt)              # 일부 실패 → 원본 보존
    assert summary.deleted == []


def test_delete_disabled_by_default(fake_setup, monkeypatch):
    tmp_path, files = fake_setup
    monkeypatch.setitem(C.ENGINE_FOR_EXT, ".txt", "fake")
    monkeypatch.setitem(C.SUPPORTED_INPUTS, ".out", {".txt"})

    src_txt = files[0]
    fake = FakeEngine()
    orch = Orchestrator({"fake": lambda: fake})
    summary = orch.run([src_txt], targets=[".out"], out_dir=str(tmp_path),
                       init_com=False)  # delete_originals 미지정 → 기본 False

    assert summary.success == 1
    assert os.path.exists(src_txt)              # 기본값: 원본 보존
    assert summary.deleted == []


# ── UI/로직 분리: core import 시 tkinter 미적재 ────────────────────
def test_core_does_not_import_tkinter():
    # 깨끗한 하위 프로세스에서 검증
    import subprocess
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    code = (
        "import sys; import unified_converter.core; "
        "assert 'tkinter' not in sys.modules, 'tkinter leaked into core'; "
        "print('OK')"
    )
    r = subprocess.run([sys.executable, "-c", code], cwd=root,
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert "OK" in r.stdout
