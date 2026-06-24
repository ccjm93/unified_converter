"""Excel(XLS/XLSX) 변환 콜백.

모듈 레벨 COM import 없음. 실제 COM 호출은 주입된 engine(ExcelEngine)이 담당.
"""
from __future__ import annotations

import os

from ..converter import ConversionResult

# Excel SaveAs FileFormat 상수
XL_OPENXML_WORKBOOK = 51   # xlOpenXMLWorkbook (.xlsx)
XL_TYPE_PDF = 0            # xlTypePDF


def xls_to_xlsx(src: str, out_path: str, engine) -> ConversionResult:
    wb = engine.open(src)
    try:
        wb.SaveAs(os.path.abspath(out_path), FileFormat=XL_OPENXML_WORKBOOK)
    finally:
        wb.Close(SaveChanges=False)
    return ConversionResult(True, src, out_path)


def to_pdf(src: str, out_path: str, engine) -> ConversionResult:
    """XLS/XLSX → PDF."""
    wb = engine.open(src)
    try:
        wb.ExportAsFixedFormat(XL_TYPE_PDF, os.path.abspath(out_path))
    finally:
        wb.Close(SaveChanges=False)
    return ConversionResult(True, src, out_path)
