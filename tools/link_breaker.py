"""Excel 외부 링크 일괄 해제 (부가 도구).

정본: 기존 excel_link_breaker_gui.py 로직(개별 링크 try/except, CorruptLoad=1)을 흡수.
차이점(데이터 손실 방지): in-place Save 전에 원본을 '.bak'으로 백업한다.

COM 초기화: 이 도구는 변환 오케스트레이터와 별개의 워커에서 실행되므로
자체적으로 pythoncom.CoInitialize/CoUninitialize를 1회 쌍으로 호출한다.
"""
from __future__ import annotations

import glob
import os
import shutil
from typing import Callable

_EXCEL_PATTERNS = ("*.xlsx", "*.xls", "*.xlsm", "*.xlsb")


def _noop(*_a, **_k) -> None:
    pass


def scan_excel(folder: str) -> list[str]:
    files: list[str] = []
    for pat in _EXCEL_PATTERNS:
        files.extend(glob.glob(os.path.join(folder, pat)))
    return sorted(files)


def break_links_in_folder(folder: str, on_log: Callable = _noop,
                          backup: bool = True) -> tuple[int, int]:
    """폴더 내 엑셀 파일의 외부 링크를 끊는다. (성공, 실패) 카운트 반환."""
    import pythoncom
    import win32com.client

    files = scan_excel(folder)
    if not files:
        on_log("선택한 폴더에 엑셀 파일이 없습니다.", "warn")
        return 0, 0

    on_log(f"📂 대상: {folder} — {len(files)}개 파일", "info")
    pythoncom.CoInitialize()
    success = fails = 0
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = 1
        except Exception:
            pass

        for idx, filepath in enumerate(files, 1):
            filepath = os.path.abspath(filepath)
            name = os.path.basename(filepath)
            on_log(f"[{idx}/{len(files)}] {name} 처리 중...", "info")
            wb = None
            try:
                try:
                    wb = excel.Workbooks.Open(filepath, UpdateLinks=0,
                                              ReadOnly=False, CorruptLoad=1)
                except Exception as e:  # noqa: BLE001
                    on_log(f"    ✗ 열기 실패: {e}", "error")
                    fails += 1
                    continue

                try:
                    links = wb.LinkSources(1)  # xlExcelLinks
                except Exception as e:  # noqa: BLE001
                    on_log(f"    ✗ 링크 목록 읽기 실패: {e}", "error")
                    fails += 1
                    wb.Close(SaveChanges=False)
                    continue

                if not links:
                    on_log("    ℹ 외부 링크 없음 (통과)", "info")
                    wb.Close(SaveChanges=False)
                    success += 1
                    continue

                broken = 0
                err = False
                for link in links:
                    try:
                        wb.BreakLink(Name=link, Type=2)  # xlLinkTypeExcelLinks
                        broken += 1
                    except Exception as e:  # noqa: BLE001
                        on_log(f"    ⚠ 링크 끊기 실패 ('{link}'): {e}", "warn")
                        err = True

                if broken > 0 or not err:
                    if backup:
                        try:
                            shutil.copy2(filepath, filepath + ".bak")
                        except Exception as e:  # noqa: BLE001
                            on_log(f"    ⚠ 백업 실패(저장 중단): {e}", "warn")
                            fails += 1
                            wb.Close(SaveChanges=False)
                            continue
                    try:
                        wb.Save()
                        on_log(f"    ✓ {broken}개 링크 제거 완료", "success")
                        success += 1 if not err else 0
                        fails += 1 if err else 0
                    except Exception as e:  # noqa: BLE001
                        on_log(f"    ✗ 저장 실패: {e}", "error")
                        fails += 1
                else:
                    fails += 1
            except Exception as e:  # noqa: BLE001
                on_log(f"    ✗ 예기치 않은 오류: {e}", "error")
                fails += 1
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
    except Exception as e:  # noqa: BLE001
        on_log(f"🚨 Excel 초기화 오류: {e} (Excel 설치 확인)", "error")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    on_log(f"🎉 완료 — 성공 {success} / 실패 {fails}", "success")
    return success, fails
