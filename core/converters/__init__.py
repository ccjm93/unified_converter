"""변환 콜백을 디스패치 레지스트리에 등록한다.

이 패키지를 import하면 모든 변환 경로가 등록된다. (COM은 import되지 않음)
"""
from ..converter import register
from . import hancom_converters as H
from . import excel_converters as E

# (입력확장자, 출력확장자) → 콜백
register(".hwp", ".hwpx", H.hwp_to_hwpx)
register(".hwp", ".pdf", H.to_pdf)
register(".hwpx", ".pdf", H.to_pdf)
register(".xls", ".xlsx", E.xls_to_xlsx)
register(".xls", ".pdf", E.to_pdf)
register(".xlsx", ".pdf", E.to_pdf)
