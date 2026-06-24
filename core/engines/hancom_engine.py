"""한컴(HWPFrame.HwpObject) COM 엔진.

설계 불변식:
  1. Dispatch(late-binding)만 사용한다. gencache.EnsureDispatch 금지
     → frozen exe(PyInstaller) 배포 시 gen_py 캐시 부재로 인한 실패를 방지.
  2. 이 클래스는 pythoncom.CoInitialize를 호출하지 않는다.
     워커 스레드(Orchestrator.run)가 진입 시 1회 CoInitialize를 책임진다.
     (참고: pyhwpx는 내부적으로 조건부 CoInitialize + gencache를 호출하므로 본 엔진은
      pyhwpx를 쓰지 않고 win32com.client.Dispatch로 한컴 COM에 직접 접근한다.)

수명주기: 1차 C2(파일마다 생성/종료). 같은 파일의 여러 출력은 ensure_open() 캐시로 단일 Open 공유.
"""
from __future__ import annotations

import os

from ._proc import ProcessTracker

_PROGID = "HWPFrame.HwpObject"
_HWP_IMAGE = "Hwp.exe"


class HancomEngine:
    def __init__(self) -> None:
        import win32com.client  # 지연 import

        tracker = ProcessTracker(_HWP_IMAGE)
        tracker.before()
        # late-binding Dispatch (gencache 미사용 — 배포 안정성)
        self.hwp = win32com.client.Dispatch(_PROGID)
        tracker.after()
        self._tracker = tracker

        # 보안 승인 모듈 등록 → 파일 접근 시 모달 팝업 방지
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        self._open_path: str | None = None

    # ── 문서 열기(캐시) ──────────────────────────────────────────
    def ensure_open(self, src_path: str):
        """동일 문서를 한 번만 Open한다(hwp→hwpx+pdf 동시출력 시 단일 Open 공유)."""
        src_path = os.path.abspath(src_path)
        if self._open_path != src_path:
            self.hwp.Open(src_path, "", "forceopen:true")
            self._open_path = src_path
        return self.hwp

    # ── 저장 ─────────────────────────────────────────────────────
    def save_as(self, out_path: str, fmt: str) -> None:
        """fmt: 'HWPX' | 'PDF' 등 한컴 FileSaveAs 포맷 문자열."""
        out_path = os.path.abspath(out_path)
        self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        params = self.hwp.HParameterSet.HFileOpenSave
        params.filename = out_path
        params.Format = fmt
        self.hwp.HAction.Execute("FileSaveAs_S", params.HSet)

    # ── 종료 / 강제 종료 ─────────────────────────────────────────
    def quit(self) -> None:
        try:
            self.hwp.Quit()
        except Exception:
            pass
        self._open_path = None

    def force_kill(self) -> None:
        """워치독: 손상 문서로 행이 걸린 한컴 프로세스를 강제 종료."""
        self._tracker.kill()
        self._open_path = None
