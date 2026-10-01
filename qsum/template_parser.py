# -*- coding: utf-8 -*-
"""模板解析模块：从 docx/xlsx/txt/md/csv/pdf 模板中提取结构。

解析产物 TemplateSpec：
- title: 报告标题
- sections: 章节列表，每章含 标题/类型/提示文字/条目占位标题
  类型 kind: overview（总体情况）| completed（已完成）| wip（进行中）| custom（自定义章节）
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class TemplateSection:
    heading: str
    kind: str  # overview | completed | wip | custom
    hints: list[str] = field(default_factory=list)
    item_titles: list[str] = field(default_factory=list)


@dataclass
class TableFormSpec:
    """表格表单模板（如绩效 PBC）：标题 + 抬头字段 + 多列表格 + 表尾。"""
    title: str
    fields: list[str] = field(default_factory=list)      # 抬头信息，如 "员工姓名/工号：xx"
    columns: list[str] = field(default_factory=list)     # 扁平化列名
    col_roles: list[str] = field(default_factory=list)   # index|item|weight|std|other
    default_std: dict[int, str] = field(default_factory=dict)  # 评价标准列的模板默认值
    footer: list[str] = field(default_factory=list)      # 总计/签字等尾部行


@dataclass
class TemplateSpec:
    title: str
    sections: list[TemplateSection]
    source: str  # 模板来源格式
    form: TableFormSpec | None = None


# ---------- 标题/章节/条目识别规则 ----------
_KIND_KEYWORDS = [
    ("overview", ("总体情况", "总体", "概况", "总览", "概述", "整体情况")),
    ("completed", ("已完成", "完成事项", "主要成果", "工作成果", "成果",
                 "今日工作", "本日工作", "今日完成", "本周工作", "本周完成", "上周工作",
                 "本月工作", "本月完成", "工作内容", "完成情况", "工作进展", "进展",
                 "总结", "小结")),
    ("wip", ("进行中", "在研", "推进中", "在办", "在建")),
    ("plan", ("下一步", "下步", "后续计划", "下季度计划", "工作计划", "下年", "明年",
             "规划", "计划", "明日", "明天", "下周工作", "下周安排", "下月工作",
             "后续安排", "工作安排", "下步安排")),
]
_SECTION_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百0-9]+[章节部分篇]|[一二三四五六七八九十]+、)"
)
_ITEM_RE = re.compile(r"^\s*[（(][一二三四五六七八九十0-9]+[）)]|^\s*\d{1,2}[、.]\s+")
_TITLE_KEYWORDS = ("总结", "报告", "汇报", "模板", "小结")
MAX_TITLE_LEN = 40


def classify_heading(text: str) -> str:
    for kind, kws in _KIND_KEYWORDS:
        if any(k in text for k in kws):
            return kind
    return "custom"


def _looks_like_title(line: str) -> bool:
    return len(line) <= MAX_TITLE_LEN and any(k in line for k in _TITLE_KEYWORDS)


# ---------- 各格式抽取为“行文本” ----------
def _lines_from_docx(path: Path) -> list[tuple[str, str]]:
    """返回 [(style, text)]，style: title/h1/h2/body。"""
    from docx import Document

    doc = Document(str(path))
    out: list[tuple[str, str]] = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        try:
            name = p.style.name if p.style is not None else ""
        except Exception:
            name = ""
        if name == "Title":
            out.append(("title", t))
        elif name.startswith("Heading 1"):
            out.append(("h1", t))
        elif name.startswith("Heading 2"):
            out.append(("h2", t))
        else:
            out.append(("body", t))
    return out


def _lines_from_md(path: Path) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            out.append(("title", line[2:].strip()))
        elif line.startswith("## "):
            out.append(("h1", line[3:].strip()))
        elif line.startswith("### "):
            out.append(("h2", line[4:].strip()))
        else:
            out.append(("body", line.lstrip(">-*• ").strip()))
    return out


def _lines_from_plain(text: str) -> list[tuple[str, str]]:
    """txt/csv 等纯文本：按行识别标题/章节/条目/正文。"""
    out: list[tuple[str, str]] = []
    for raw in text.splitlines():
        line = raw.strip().strip("|").strip()  # csv 行可能带 |
        if not line:
            continue
        if _SECTION_RE.match(line):
            out.append(("h1", line))
        elif _ITEM_RE.match(line):
            out.append(("h2", line))
        elif not out and _looks_like_title(line):
            out.append(("title", line))
        else:
            out.append(("body", line))
    return out


def _lines_from_xlsx(path: Path) -> list[tuple[str, str]]:
    """Excel 模板：按行主序展开非空单元格，再走纯文本识别。"""
    from .extractors import extract_xlsx_sheets

    texts: list[str] = []
    for sheet in extract_xlsx_sheets(path):
        for row in sheet["rows"]:
            for cell in row:
                if cell:
                    texts.append(cell)
    return _lines_from_plain("\n".join(texts))


# ---------- 组装 TemplateSpec ----------
def _build_spec(styled_lines: list[tuple[str, str]], source: str) -> TemplateSpec:
    title = ""
    sections: list[TemplateSection] = []
    current: TemplateSection | None = None

    for style, text in styled_lines:
        if style == "title":
            title = title or text
            continue
        if style == "h1" or (style == "body" and _SECTION_RE.match(text)):
            heading = text if style == "h1" else text.strip()
            current = TemplateSection(heading=heading, kind=classify_heading(heading))
            sections.append(current)
            continue
        if style == "h2" or (style == "body" and _ITEM_RE.match(text)):
            if current is not None:
                current.item_titles.append(text)
            continue
        if not title and not sections and _looks_like_title(text):
            title = text
            continue
        if current is not None:
            current.hints.append(text)
        # 标题之前、章节之外的散行忽略

    if not title:
        title = "工作总结"
    return TemplateSpec(title=title, sections=sections, source=source)


# ---------- 表格表单识别（绩效 PBC 等） ----------
_FORM_HEADER_KW = ("指标", "kpi", "KPI", "权重", "标准", "目标", "分值", "自评",
                  "得分", "考核", "评价", "承诺", "职责", "任务", "内容", "序")
_FORM_FOOTER_KW = ("总计", "合计", "签字", "签名", "日期：", "确认")


def _col_role(name: str) -> str:
    if name.startswith("序") or name in ("序号", "No", "no"):
        return "index"
    if any(k in name for k in ("权重", "占比", "分值")):
        return "weight"
    if any(k in name for k in ("标准", "衡量", "评分", "说明")):
        return "std"
    if any(k in name for k in ("指标", "KPI", "kpi", "内容", "任务", "目标",
                             "事项", "职责", "承诺", "工作")):
        return "item"
    return "other"


def _flatten_headers(header_rows: list[list[str]], ncols: int) -> list[str]:
    """多级表头扁平化：取每列最深的非空层级名。"""
    cols: list[str] = []
    for c in range(ncols):
        name = ""
        for row in header_rows:  # 自上而下，深的覆盖浅的
            if c < len(row) and row[c]:
                name = row[c]
        # 合并列的文字归给首列，后续空列用上一级名字补全
        cols.append(re.sub(r"[\s；;]+", "", name) if name else "")
    # 空列名：继承左侧非空（合并单元格场景）
    last = ""
    for i, name in enumerate(cols):
        if name:
            last = name
        elif last:
            cols[i] = ""
    return cols


def detect_form(sheets: list[dict]) -> TableFormSpec | None:
    """从表格工作表中检测表单结构；不符合返回 None。"""
    for sheet in sheets:
        rows = [sheet["headers"]] + sheet["rows"] if sheet["headers"] else sheet["rows"]
        # 找表头行：≥3 非空单元格且 ≥2 个含表单关键词
        header_idx = None
        for i, row in enumerate(rows):
            non_empty = [c for c in row if c]
            hits = sum(1 for c in non_empty if any(k in c for k in _FORM_HEADER_KW))
            if len(non_empty) >= 3 and hits >= 2:
                header_idx = i
                break
        if header_idx is None:
            continue
        ncols = max((len(r) for r in rows), default=0)
        # 多级表头：表头行 + 其后的子表头行（如“底线/达标/卓越标准”）
        header_rows = [rows[header_idx]]
        hdr_end = header_idx + 1
        if hdr_end < len(rows):
            nxt_cells = [c for c in rows[hdr_end] if c]
            if len(nxt_cells) >= 2 and all(any(k in c for k in _FORM_HEADER_KW)
                                          for c in nxt_cells):
                header_rows.append(rows[hdr_end])
                hdr_end += 1
        columns = _flatten_headers(header_rows, ncols)
        roles = [_col_role(c) for c in columns]
        if "item" not in roles:
            continue

        # 标题与抬头字段（表头之前）
        title = ""
        fields: list[str] = []
        for row in rows[:header_idx]:
            cells = [c for c in row if c]
            if not cells:
                continue
            if not title and len(cells) <= 2 and _looks_like_title(cells[0]):
                title = cells[0]
                continue
            if not title and not fields:
                title = cells[0]
                continue
            for j in range(0, len(cells) - 1, 2):
                fields.append(f"{cells[j].rstrip('：:')}：{cells[j + 1]}")
            if len(cells) % 2 == 1:
                fields.append(cells[-1])

        # 数据行 + 表尾 + 默认评价标准
        footer: list[str] = []
        default_std: dict[int, str] = {}
        for row in rows[hdr_end:]:
            text = "".join(row)
            if any(k in text for k in _FORM_FOOTER_KW):
                footer.append("  ".join(c for c in row if c))
                continue
            if not default_std:
                for c, role in enumerate(roles):
                    if role == "std" and c < len(row) and row[c]:
                        default_std[c] = row[c]
        return TableFormSpec(
            title=title or "工作总结",
            fields=fields, columns=columns, col_roles=roles,
            default_std=default_std, footer=footer,
        )
    return None


def parse_template(path: Path | str) -> TemplateSpec:
    """解析模板文件，返回结构化 TemplateSpec。"""
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".docx":
        lines = _lines_from_docx(path)
        source = "docx"
    elif ext == ".md":
        lines = _lines_from_md(path)
        source = "md"
    elif ext in (".xlsx", ".xls"):
        if ext == ".xlsx":
            from .extractors import extract_xlsx_sheets

            sheets = extract_xlsx_sheets(path)
        else:
            from .extractors import extract_xls_sheets

            sheets = extract_xls_sheets(path)
        form = detect_form(sheets)
        if form is not None:
            return TemplateSpec(title=form.title, sections=[], source=ext.lstrip("."),
                                form=form)
        texts: list[str] = []
        for sheet in sheets:
            for row in sheet["rows"]:
                texts.extend(c for c in row if c)
        lines = _lines_from_plain("\n".join(texts))
        source = ext.lstrip(".")
    elif ext in (".txt", ".csv"):
        text = path.read_text(encoding="utf-8", errors="replace")
        lines = _lines_from_plain(text)
        source = ext.lstrip(".")
    elif ext == ".pdf":
        from .extractors import extract_file

        lines = _lines_from_plain(extract_file(path))
        source = "pdf"
    else:
        raise ValueError(f"不支持的模板格式: {ext}（支持 docx/xlsx/txt/md/csv/pdf）")
    spec = _build_spec(lines, source)
    if not spec.sections:
        raise ValueError(f"未能从模板中识别出章节结构（如「一、总体情况」这样的标题行）或表格表单结构: {path.name}")
    return spec
