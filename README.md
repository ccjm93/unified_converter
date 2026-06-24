# 통합 파일형식 변환기 (Unified Converter)

기존 분리되어 있던 변환 프로그램을 하나로 합치고 Excel 변환을 추가한 Windows 데스크톱 GUI 도구.

## 지원 변환

| 입력 → 출력 | HWPX | PDF | XLSX |
|---|---|---|---|
| HWP | ✓ | ✓ | - |
| HWPX | - | ✓ | - |
| XLS | - | ✓ | ✓ |
| XLSX | - | ✓ | - |

- **동시 출력**: HWP는 HWPX·PDF를 한 번에(파일을 한 번만 열어 두 형식 저장).
- **혼합 배치**: HWP·XLS·HWPX를 한 목록에 섞어 추가 → 선택한 출력형식으로 변환 가능한 항목만 처리, 나머지는 "건너뜀".
- **부가 도구**: [도구] 메뉴 → Excel 외부 링크 일괄 해제(원본 `.bak` 백업 후 처리).

## 요구 사항

- Windows + Python 3.10+
- **한컴오피스(한글)** — HWP/HWPX 변환에 필요
- **Microsoft Excel** — XLS/XLSX 변환에 필요
- 둘 중 하나만 설치돼 있어도 해당 변환만 활성화됨(나머지는 자동 안내).

```bash
pip install -r requirements.txt
```

## 실행

```bash
python main.py
```

## 빌드 (단일 exe)

```bash
# unified_converter 폴더에서
pyinstaller build/unified_converter.spec
# 결과: dist/통합변환기.exe
```

## 테스트

```bash
python -m pytest tests/ -q       # COM 불필요한 헤드리스 코어 검증
```

## 설계 메모

- 한컴/Excel COM은 `win32com.client.Dispatch`(late-binding)만 사용한다. `gencache.EnsureDispatch`는 PyInstaller 배포 시 `gen_py` 캐시 부재로 실패하므로 쓰지 않는다.
- `pyhwpx`는 내부적으로 gencache와 조건부 CoInitialize를 호출하므로 의도적으로 사용하지 않고, 한컴 COM에 직접 접근한다.
- **COM 초기화 불변식**: `pythoncom.CoInitialize`는 워커 스레드 진입 시 1회만 호출하며, 엔진 어댑터(`core/engines/*`)는 절대 호출하지 않는다.
- 변환 코어(`core/`)는 tkinter import 없이 헤드리스로 동작·테스트된다.
- 손상 문서로 한컴이 멈추면 파일별 워치독 타임아웃이 해당 프로세스를 강제 종료한다.

## 구조

```
core/        변환 디스패치·오케스트레이션·엔진·변환기 (COM 격리)
ui/          단일 화면 GUI (Catppuccin Mocha 테마)
tools/       부가 도구(Excel 링크 해제)
utils/       파일 스캔·자동 리네임·로깅
build/       PyInstaller 스펙
tests/       헤드리스 코어 단위 테스트
```
