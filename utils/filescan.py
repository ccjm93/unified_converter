"""파일 스캔 및 출력 경로 산출 유틸. COM/tkinter 의존 없음."""
from __future__ import annotations

import os

# 통합 변환기가 입력으로 받아들이는 확장자
INPUT_EXTS = (".hwp", ".hwpx", ".xls", ".xlsx")


def split_name_ext(path: str) -> tuple[str, str]:
    """('basename_without_ext', '.ext') 반환. ext는 소문자."""
    base = os.path.basename(path)
    stem, ext = os.path.splitext(base)
    return stem, ext.lower()


def unique_output_path(out_dir: str, base_name: str, ext: str) -> str:
    """출력 경로를 산출하되, 이미 존재하면 'name (1).ext' 식으로 자동 리네임한다.

    데이터 손실 방지(확정 정책): 기존 파일을 절대 덮어쓰지 않는다.
    """
    if not ext.startswith("."):
        ext = "." + ext
    candidate = os.path.join(out_dir, base_name + ext)
    if not os.path.exists(candidate):
        return candidate
    i = 1
    while True:
        candidate = os.path.join(out_dir, f"{base_name} ({i}){ext}")
        if not os.path.exists(candidate):
            return candidate
        i += 1


def scan_folder(folder: str, exts: tuple[str, ...] = INPUT_EXTS) -> list[str]:
    """폴더를 재귀 스캔하여 지원 확장자 파일 경로 목록을 반환(정렬됨)."""
    found: list[str] = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(exts):
                found.append(os.path.join(root, f))
    found.sort()
    return found
