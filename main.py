"""통합 파일형식 변환기 — 엔트리포인트."""
from __future__ import annotations


def main() -> None:
    # 패키지/스크립트 양쪽 실행 지원
    try:
        from unified_converter.ui import App
    except ImportError:
        import os
        import sys
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from unified_converter.ui import App

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
