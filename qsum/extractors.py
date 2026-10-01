# -*- coding: utf-8 -*-
"""多格式文件文本抽取模块。

支持格式: .docx / .xlsx / .txt / .md / .csv / .pdf
(旧版 .doc / .xls 建议先转换为 .xlsx/.docx，会给出提示)
"""
from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from pathlib import Path

SUPPORTED_EXTS = {".docx", ".xlsx", ".xls", ".txt", ".md", ".csv", ".pdf"}
LEGACY_EXTS = {".doc", ".wps"}

MAX_XLSX_ROWS_PER_SHEET = 2000
MAX_PDF_PAGES = 200


class ExtractError(Exception):
    pass


def extract_xlsx_sheets(path: Path) -> list[dict]:
    """结构化读取 xlsx：返回 [{"title", "headers", "rows"}]，值已格式化。"""
    try:
        import openpyxl
    except ImportError:
        raise ExtractError("缺少 openpyxl 库，请执行: pip install openpyxl")
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    sheets: list[dict] = []
    for ws in wb.worksheets:
        rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            vals = [_fmt_cell(v) for v in row]
            if any(vals):
                rows.append(vals)
            if len(rows) > MAX_XLSX_ROWS_PER_SHEET:
                break
        headers = rows[0] if rows else []
        sheets.append({"title": ws.title, "headers": headers,
                       "rows": rows[1:] if len(rows) > 1 else []})
    wb.close()
    return sheets


def _fmt_xls_cell(ctype: int, value, datemode: int) -> str:
    """xlrd 单元格格式化：日期、整数浮点、换行归一。"""
    if ctype == 0 or value in (None, ""):
        return ""
    if ctype == 3:  # XL_CELL_DATE
        try:
            from xlrd import xldate

            dt = xldate.xldate_as_datetime(value, datemode)
            return _fmt_cell(dt)
        except Exception:
            return str(value)
    if ctype == 2 and isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip().replace("\r\n", "；").replace("\r", "；").replace("\n", "；")
    return re.sub(r"[；;]{2,}", "；", s).strip("；; ")


def _open_xls(path: Path):
    try:
        import xlrd
    except ImportError:
        raise ExtractError("缺少 xlrd 库（读取旧版 .xls），请执行: pip install xlrd")
    return xlrd.open_workbook(str(path))


def _extract_xls(path: Path) -> str:
    wb = _open_xls(path)
    parts: list[str] = []
    for ws in wb.sheets():
        parts.append(f"## 工作表: {ws.name}")
        count = 0
        for r in range(ws.nrows):
            vals = [_fmt_xls_cell(ws.cell_type(r, c), ws.cell_value(r, c), wb.datemode)
                    for c in range(ws.ncols)]
            if any(vals):
                parts.append(" | ".join(vals))
                count += 1
            if count >= MAX_XLSX_ROWS_PER_SHEET:
                parts.append(f"（工作表 {ws.name} 内容过多，已截断）")
                break
    return "\n".join(parts)


def extract_xls_sheets(path: Path) -> list[dict]:
    """结构化读取 .xls：返回 [{"title", "headers", "rows"}]，值已格式化。"""
    wb = _open_xls(path)
    sheets: list[dict] = []
    for ws in wb.sheets():
        rows: list[list[str]] = []
        for r in range(min(ws.nrows, MAX_XLSX_ROWS_PER_SHEET)):
            vals = [_fmt_xls_cell(ws.cell_type(r, c), ws.cell_value(r, c), wb.datemode)
                    for c in range(ws.ncols)]
            if any(vals):
                rows.append(vals)
        headers = rows[0] if rows else []
        sheets.append({"title": ws.name, "headers": headers,
                       "rows": rows[1:] if len(rows) > 1 else []})
    return sheets


def extract_file(path: Path) -> str:
    """抽取单个文件的文本内容，返回纯文本。"""
    path = Path(path)
    if not path.exists():
        raise ExtractError(f"文件不存在: {path}")
    ext = path.suffix.lower()
    if ext == ".docx":
        return _extract_docx(path)
    if ext == ".xlsx":
        return _extract_xlsx(path)
    if ext == ".xls":
        return _extract_xls(path)
    if ext == ".pdf":
        return _extract_pdf(path)
    if ext in (".txt", ".md"):
        return _read_text(path)
    if ext == ".csv":
        return _extract_csv(path)
    if ext in LEGACY_EXTS:
        raise ExtractError(f"暂不支持旧版格式 {ext}，请另存为 .docx/.xlsx: {path}")
    raise ExtractError(f"不支持的文件格式 {ext}: {path}")


def _extract_docx(path: Path) -> str:
    try:
        from docx import Document
    except ImportError:
        raise ExtractError("缺少 python-docx 库，请执行: pip install python-docx")
    doc = Document(str(path))
    parts: list[str] = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _fmt_cell(v) -> str:
    """Excel 单元格值格式化：日期去时分秒、整数浮点去 .0、换行归一。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        if v.hour == v.minute == v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip().replace("\r\n", "；").replace("\r", "；").replace("\n", "；")
    return re.sub(r"[；;]{2,}", "；", s).strip("；; ")


def _extract_xlsx(path: Path) -> str:
    try:
        import openpyxl
    except ImportError:
        raise ExtractError("缺少 openpyxl 库，请执行: pip install openpyxl")
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    parts: list[str] = []
    for ws in wb.worksheets:
        parts.append(f"## 工作表: {ws.title}")
        count = 0
        for row in ws.iter_rows(values_only=True):
            vals = [_fmt_cell(v) for v in row]
            if any(vals):
                parts.append(" | ".join(vals))
                count += 1
            if count >= MAX_XLSX_ROWS_PER_SHEET:
                parts.append(f"（工作表 {ws.title} 内容过多，已截断）")
                break
    wb.close()
    return "\n".join(parts)


def _extract_pdf(path: Path) -> str:
    parts: list[str] = []
    try:
        import pdfplumber

        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages):
                if i >= MAX_PDF_PAGES:
                    parts.append("（PDF 页数过多，已截断）")
                    break
                parts.append(page.extract_text() or "")
    except ImportError:
        try:
            from pypdf import PdfReader
        except ImportError:
            raise ExtractError("缺少 pdfplumber/pypdf 库，请执行: pip install pdfplumber")
        reader = PdfReader(str(path))
        for i, page in enumerate(reader.pages):
            if i >= MAX_PDF_PAGES:
                break
            parts.append(page.extract_text() or "")
    except Exception as exc:  # pdf 解析库抛出的各种异常
        raise ExtractError(f"PDF 解析失败 {path}: {exc}")
    return "\n".join(parts)


def _extract_csv(path: Path) -> str:
    text = _read_text(path)
    return text


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def scan_inputs(inputs: list[str | Path]) -> list[Path]:
    """展开输入：目录递归扫描 / 单文件，返回支持格式的文件列表。"""
    result: list[Path] = []
    seen: set[Path] = set()
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS and f not in seen:
                    seen.add(f)
                    result.append(f)
        elif p.is_file():
            if p.suffix.lower() in SUPPORTED_EXTS and p not in seen:
                seen.add(p)
                result.append(p)
        else:
            print(f"[警告] 输入路径不存在，已跳过: {p}")
    return result
