"""COM 서버 프로세스 PID 추적/종료 헬퍼 (워치독 force_kill용).

엔진 생성 직전 스냅샷을 찍고 생성 직후 차집합으로 새 프로세스 PID를 식별한다.
사용자가 별도로 열어 둔 한컴/Excel 창을 죽이지 않기 위해 '새로 생긴' PID만 종료한다.
"""
from __future__ import annotations

import subprocess


def image_pids(image_name: str) -> set[int]:
    """주어진 이미지명(예: 'Hwp.exe')의 실행 중 PID 집합."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
        ).stdout
    except Exception:
        return set()
    pids: set[int] = set()
    for line in out.splitlines():
        # CSV: "Image","PID","Session","Session#","Mem"
        parts = [p.strip().strip('"') for p in line.split('","')]
        if len(parts) >= 2:
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def kill_pid(pid: int) -> None:
    try:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                       capture_output=True, timeout=10)
    except Exception:
        pass


class ProcessTracker:
    """with 블록 또는 before()/after()로 새 PID를 포착한다."""

    def __init__(self, image_name: str) -> None:
        self.image_name = image_name
        self._before: set[int] = set()
        self.pid: int | None = None

    def before(self) -> None:
        self._before = image_pids(self.image_name)

    def after(self) -> None:
        new = image_pids(self.image_name) - self._before
        self.pid = next(iter(new), None)

    def kill(self) -> None:
        if self.pid is not None:
            kill_pid(self.pid)
            self.pid = None
