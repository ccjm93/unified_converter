"""Excel(Excel.Application) COM 엔진.

설계 불변식:
  1. Dispatch(late-binding)만 사용 (기존 링크해제 유틸에서 검증된 방식).
  2. 이 클래스는 pythoncom.CoInitialize를 호출하지 않는다(워커가 1회 책임).

수명주기: 1차 C2(파일마다 생성/종료). 워크북은 open()마다 열고 변환 후 Close.
"""
from __future__ import annotations

import os

from ._proc import ProcessTracker

_PROGID = "Excel.Application"
_EXCEL_IMAGE = "EXCEL.EXE"


class ExcelEngine:
    def __init__(self) -> None:
        import win32com.client  # 지연 import

        tracker = ProcessTracker(_EXCEL_IMAGE)
        tracker.before()
        self.app = win32com.client.Dispatch(_PROGID)
        tracker.after()
        self._tracker = tracker

        self.app.Visible = False
        self.app.DisplayAlerts = False
        try:
            self.app.AskToUpdateLinks = False
        except Exception:
            pass
        try:
            self.app.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass

    # ── 워크북 열기 ──────────────────────────────────────────────
    def open(self, src_path: str):
        """읽기전용 + 링크 미갱신 + 손상 복구 모드로 워크북을 연다."""
        src_path = os.path.abspath(src_path)
        return self.app.Workbooks.Open(
            src_path, UpdateLinks=0, ReadOnly=True, CorruptLoad=1
        )

    # ── 종료 / 강제 종료 ─────────────────────────────────────────
    def quit(self) -> None:
        try:
            self.app.Quit()
        except Exception:
            pass

    def force_kill(self) -> None:
        self._tracker.kill()
