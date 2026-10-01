# -*- coding: utf-8 -*-
"""工作内容挖掘模块：日期识别、季度过滤、工作事项分类与项目分组。"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

# ---------- 日期识别 ----------
_DATE_PATTERNS = [
    re.compile(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?"),
    re.compile(r"(\d{4})[-/年.](\d{1,2})月?"),
]
_DATE_ANY = re.compile(
    r"\d{4}[-/年.]\d{1,2}([-/月.]\d{1,2})?日?|\d{1,2}月\d{1,2}日|\d{1,2}日"
)


def extract_dates(text: str) -> list[tuple[int, int]]:
    """返回文本中出现的 (year, month) 列表。"""
    found: list[tuple[int, int]] = []
    for pat in _DATE_PATTERNS:
        for m in pat.finditer(text):
            y, mo = int(m.group(1)), int(m.group(2))
            if 2000 <= y <= 2100 and 1 <= mo <= 12:
                found.append((y, mo))
    return found


# ---------- 报告周期（日/周/月/季） ----------
_KIND_SHORT = {"day": "本日", "week": "本周", "month": "本月", "quarter": "本季度"}
_KIND_SUFFIX = {"day": "日报", "week": "周报", "month": "月报", "quarter": "工作总结"}


@dataclass
class Period:
    kind: str          # day | week | month | quarter
    start: date
    end: date
    label: str         # 2026年10月1日 / 2026年第40周 / 2026年10月 / 2026年第三季度
    short: str         # 本日/本周/本月/本季度
    suffix: str        # 日报/周报/月报/工作总结
    raw: str           # 原始输入或 label


def parse_period(kind: str | None = None, value: str | None = None) -> Period:
    """解析报告周期。value 支持: 2026-10-01 / 2026W40 / 2026-10 / 2026Q3；
    kind 可指定日/周/月/季（与 value 格式不一致时按 value 的日期换算）。"""
    ref = date.today()
    explicit_kind = None
    if value:
        v = value.strip()
        m = re.fullmatch(r"(\d{4})[Qq]([1-4])", v)
        if m:
            explicit_kind = "quarter"
            ref = date(int(m.group(1)), (int(m.group(2)) - 1) * 3 + 1, 1)
        else:
            m = re.fullmatch(r"(\d{4})[Ww](\d{1,2})", v)
            if m:
                explicit_kind = "week"
                ref = date.fromisocalendar(int(m.group(1)), int(m.group(2)), 1)
            else:
                m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", v)
                if m:
                    explicit_kind = "day"
                    ref = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                else:
                    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})", v)
                    if m:
                        explicit_kind = "month"
                        ref = date(int(m.group(1)), int(m.group(2)), 1)
                    else:
                        raise ValueError(
                            f"无法识别的时间: {v}（示例: 2026-10-01 / 2026W40 / 2026-10 / 2026Q3）")
    k = kind or explicit_kind or "quarter"
    if k not in _KIND_SHORT:
        raise ValueError(f"未知报告类型: {k}（支持 day/week/month/quarter）")
    if k == "day":
        start = end = ref
        label = f"{ref.year}年{ref.month}月{ref.day}日"
    elif k == "week":
        start = ref - timedelta(days=ref.weekday())
        end = start + timedelta(days=6)
        iso = start.isocalendar()
        label = f"{iso[0]}年第{iso[1]}周"
    elif k == "month":
        start = ref.replace(day=1)
        end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        label = f"{start.year}年{start.month}月"
    else:
        q = (ref.month - 1) // 3 + 1
        start = date(ref.year, (q - 1) * 3 + 1, 1)
        em = (q - 1) * 3 + 3
        end = date(ref.year + (1 if em == 12 else 0), 1 if em == 12 else em + 1, 1) \
            - timedelta(days=1)
        label = f"{ref.year}年第{"一二三四"[q - 1]}季度"
    return Period(kind=k, start=start, end=end, label=label,
                  short=_KIND_SHORT[k], suffix=_KIND_SUFFIX[k],
                  raw=value.strip() if value else label)


_FULL_DATE_RE = re.compile(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?")


def text_in_period(text: str, period: Period) -> bool | None:
    """文本中的日期是否落在周期内。True/False；无日期可判返回 None。"""
    full: list[date] = []
    for m in _FULL_DATE_RE.finditer(text):
        try:
            full.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            continue
    if full:
        return any(period.start <= d <= period.end for d in full)
    partial = extract_dates(text)
    if partial:
        months: set[tuple[int, int]] = set()
        d = period.start
        while d <= period.end:
            months.add((d.year, d.month))
            d = (d.replace(day=28) + timedelta(days=4)).replace(day=1)
        return any(p in months for p in partial)
    return None


def file_in_period(path: Path, text: str, period: Period) -> bool:
    """文件是否属于目标周期（内容日期优先，其次文件修改时间）。"""
    hit = text_in_period(text, period)
    if hit is not None:
        return hit
    mtime = datetime.fromtimestamp(path.stat().st_mtime).date()
    return period.start <= mtime <= period.end


def filter_items(items: list["WorkItem"], period: Period) -> list["WorkItem"]:
    """条目级时间过滤：无日期的条目保留（无法判断时从宽）。"""
    return [it for it in items if text_in_period(it.text, period) is not False]


# ---------- 项目名识别 ----------
PROJECT_KEYWORDS = [
    "系统", "平台", "应用", "项目", "模块", "工具", "服务",
    "门户", "网站", "中台", "中心", "APP", "App", "app",
]
PRIMARY_KEYWORDS = {
    "系统", "平台", "应用", "项目", "工具", "中台", "中心", "网站", "门户",
    "APP", "App", "app",
}
KW_RE = re.compile("|".join(map(re.escape, PROJECT_KEYWORDS)))
NAME_CHAR_RE = re.compile(r"[A-Za-z0-9\u4e00-\u9fa5（）()·\-]")
NOISE_PREFIXES = sorted(
    [
        "正在推进", "正在进行", "正在", "本季度", "本年度", "本次", "该", "这个",
        "完成", "进入", "启动", "开展", "推进", "签订", "通过", "交付", "验收",
        "负责", "参与", "进行", "组织", "配合", "落实", "已", "待", "的", "了",
        "新", "将", "与", "和", "及", "对", "为", "在", "并", "后", "前",
    ],
    key=len, reverse=True,
)
FALLBACK_STRIP = re.compile(
    r"(工作日志|日志|工作记录|记录|纪要|周报|月报|季度总结|总结|台账|清单)$"
)
_LIST_MARKER = re.compile(r"^\s*(?:\d{1,2}[\.、\)）:：]\s*|[•·>\-]+\s*)")


def _clean_name(raw: str) -> str:
    """去掉项目名前导的日期、序号、动词等噪声。"""
    name = raw.strip()
    changed = True
    while changed and len(name) > 2:
        changed = False
        new = _DATE_ANY.sub("", name).lstrip("年月日")
        if new and new != name:
            name, changed = new.lstrip("年月日"), True
            continue
        new = _LIST_MARKER.sub("", name)
        if new and new != name:
            name, changed = new, True
            continue
        for pref in NOISE_PREFIXES:
            if name.startswith(pref) and len(name) > len(pref) + 1:
                name = name[len(pref):]
                changed = True
                break
    return name.strip("，,、：:；;·-— ")


def candidate_names(text: str) -> set[str]:
    """以「系统/平台/应用…」关键词为锚点，向左最大扩展提取候选项目名。"""
    out: set[str] = set()
    for m in KW_RE.finditer(text):
        start = m.start()
        while start > 0 and NAME_CHAR_RE.match(text[start - 1]):
            start -= 1
        name = _clean_name(text[start:m.end()])
        if 3 <= len(name) <= 20:
            out.add(name)
    return out


def _is_primary(name: str) -> bool:
    return any(name.endswith(k) for k in PRIMARY_KEYWORDS)


def resolve_projects(items: list["WorkItem"]) -> None:
    """全局投票确定每个条目的项目名：出现频次高 > 主体关键词结尾 > 名字短。"""
    freq: Counter = Counter()
    for it in items:
        for c in it.candidates:
            freq[c] += 1
    for it in items:
        if it.candidates:
            it.project = sorted(
                it.candidates,
                key=lambda c: (-freq[c], -int(_is_primary(c)), len(c)),
            )[0]
        else:
            it.project = it.fallback


# ---------- 工作事项分类 ----------
COMPLETED_VERBS = [
    "上线试运行", "试运行", "上线", "正式运行", "投入使用", "投产",
    "验收", "结项", "交付", "签订", "合同", "论证", "评审", "发布",
    "部署", "完成", "通过复测", "整改完成",
]
STRONG_WIP_VERBS = [
    "进行中", "开发中", "测试中", "联调中", "推进中", "筹备中", "规划中",
    "优化中", "整改中", "迭代中", "建设中", "研发中", "改造中", "完善中",
    "待上线", "待验收", "准备中", "即将", "尚未", "未完成",
]
WEAK_WIP_VERBS = ["推进", "优化", "联调", "筹备", "规划", "准备", "预计", "计划"]

_SEGMENT_SPLIT = re.compile(r"[。！？；\n]+")
_ROW_SEP = re.compile(r"\s*\|\s*")
_STATUS_HINT = {
    "已完成": "completed", "完成": "completed", "已上线": "completed",
    "done": "completed", "ok": "completed", "closed": "completed",
    "进行中": "wip", "开发中": "wip", "测试中": "wip", "未开始": "wip",
    "待开始": "wip", "已延期": "wip", "wip": "wip", "doing": "wip",
    "todo": "wip", "pending": "wip",
}
_STATUS_CELLS = set(_STATUS_HINT) | {"已取消", "状态", "进度", "阶段"}
MIN_SEG_LEN = 6
MAX_SEG_LEN = 400


@dataclass
class WorkItem:
    text: str
    status: str  # 'completed' | 'wip'
    project: str = ""
    candidates: set[str] = field(default_factory=set)
    fallback: str = "其他工作"
    source: str = ""
    has_metric: bool = False


@dataclass
class WorkDigest:
    period: str  # 周期原始串，如 2026Q3 / 2026W40 / 2026-10 / 2026-10-01
    files: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    items_completed: list[WorkItem] = field(default_factory=list)
    items_wip: list[WorkItem] = field(default_factory=list)
    metric_lines: list[str] = field(default_factory=list)

    @property
    def all_items(self) -> list[WorkItem]:
        return self.items_completed + self.items_wip


_METRIC_RE = re.compile(
    r"\d[\d,.]*\s*(?:个|项|次|人|名|家|台|套|件|篇|条|张|份|笔|户|批|轮|省|市|%"
    r"|％|万|亿|小时|天|周|分|秒|倍|万余元|万元)"
)


def _normalize(text: str) -> str:
    return re.sub(r"[\s，,。；;：:、（）()【】\[\]“”\"']", "", text)


def _classify(segment: str) -> str | None:
    """按「首个状态标记」判定：强标记按出现位置先后定状态。"""
    strong = [(segment.find(v), "completed") for v in COMPLETED_VERBS if v in segment]
    strong += [(segment.find(v), "wip") for v in STRONG_WIP_VERBS if v in segment]
    if strong:
        strong.sort(key=lambda x: x[0])
        return strong[0][1]
    if any(v in segment for v in WEAK_WIP_VERBS):
        return "wip"
    return None


def _clean_row(seg: str) -> tuple[str, str | None]:
    """表格行清洗：去掉状态单元格、重复单元格；返回 (文本, 状态提示)。
    状态列比动词猜测更可靠，优先作为分类依据。"""
    cells = [c.strip() for c in _ROW_SEP.split(seg) if c.strip()]
    hint = None
    kept: list[str] = []
    for i, c in enumerate(cells):
        low = c.lower()
        if low in _STATUS_CELLS:
            hint = hint or _STATUS_HINT.get(low)
            continue
        if any(i != j and c in other for j, other in enumerate(cells)):
            continue
        kept.append(c)
    return "，".join(kept), hint


def _split_segments(text: str) -> list[tuple[str, str | None]]:
    """切分为 (段文本, 行状态提示) 列表，提示仅来自表格状态列。"""
    segs: list[tuple[str, str | None]] = []
    for part in _SEGMENT_SPLIT.split(text):
        part = _LIST_MARKER.sub("", part.strip()).strip()
        if not part:
            continue
        hint = None
        if _ROW_SEP.search(part) and len(part) <= MAX_SEG_LEN:
            part, hint = _clean_row(part)
        if MIN_SEG_LEN <= len(part) <= MAX_SEG_LEN:
            segs.append((part, hint))
        elif len(part) > MAX_SEG_LEN:
            for sub in re.split(r"[，,]", part):
                sub = sub.strip()
                if MIN_SEG_LEN <= len(sub) <= MAX_SEG_LEN:
                    segs.append((sub, hint))
    return segs


def _is_metric(seg: str) -> bool:
    """是否包含有效量化数据（排除纯日期）。"""
    s = _DATE_ANY.sub("", seg)
    return bool(_METRIC_RE.search(s))


def _fallback_name(path: Path) -> str:
    cands = candidate_names(path.stem)
    if cands:
        return max(cands, key=len)
    stem = FALLBACK_STRIP.sub("", path.stem).strip("_- ")
    return stem if len(stem) >= 3 else "其他工作"


# ---------- 工单表结构化挖掘（表头语义识别） ----------
_FIELD_KEYS = [
    ("status", ["当前状态", "处理状态", "状态", "进度", "是否处理"]),
    ("date", ["处理时间", "结束时间", "开始时间", "完成时间", "创建时间", "日期", "时间"]),
    ("name", ["名称", "标题", "事项", "任务", "主题", "工单", "流程"]),
    ("solution", ["解决方法", "解决办法", "解决方案", "处理方法", "处理结果",
                 "解决措施", "处理措施", "措施", "结论"]),
    ("desc", ["描述", "内容", "说明", "详情", "需求", "背景", "现象", "问题", "备注"]),
]
_TICKET_STATUS_MAP = {
    "已完成": "completed", "完成": "completed", "已解决": "completed",
    "已处理": "completed", "已上线": "completed", "已关闭": "completed",
    "done": "completed", "closed": "completed", "resolved": "completed",
    "处理中": "wip", "进行中": "wip", "开发中": "wip", "测试中": "wip",
    "待处理": "wip", "未处理": "wip", "未完成": "wip", "待开始": "wip",
    "已延期": "wip", "挂起": "wip", "暂缓": "wip",
    "doing": "wip", "wip": "wip", "pending": "wip", "todo": "wip",
}
_URL_RE = re.compile(r"https?://\S+|www\.\S+")
_PHONE_RE = re.compile(r"(?:1[3-9]\d{9}|0\d{2,3}-?\d{7,8}|\b\d{7,12}\b|\d+\*+\d*)")
_ID_RE = re.compile(r"(?:工号|流水号|账号|编号|电话)[:：]?\s*[A-Za-z0-9\-]+")
_MD_RE = re.compile(r"[#*`]+")
_DIALOG_RE = re.compile(
    r"^(好的|收到|了解|稍等|请问|麻烦|谢谢|老师您好|您好|您|请|嗯|在的|可以的)[，,。！!？?\s]?.*$")


def _map_columns(headers: list[str]) -> dict[str, list[int]]:
    """把表头映射到语义字段，每列只归属一个字段。"""
    mapping: dict[str, list[int]] = {}
    used: set[int] = set()
    for field, keys in _FIELD_KEYS:
        for key in sorted(keys, key=len, reverse=True):
            for i, h in enumerate(headers):
                if i in used:
                    continue
                hh = str(h).strip()
                if hh == key or (len(key) >= 2 and key in hh):
                    mapping.setdefault(field, []).append(i)
                    used.add(i)
    return mapping


def is_ticket_table(headers: list[str]) -> bool:
    m = _map_columns(headers)
    return "name" in m and ("status" in m or "solution" in m) and bool(
        set(m) & {"desc", "solution", "status"})


def _clean_title(raw: str) -> str:
    """事项/工单标题清洗：只去序号与日期，不动词（标题里的动词是名字的一部分）。"""
    s = _LIST_MARKER.sub("", raw.strip())
    s = _DATE_ANY.sub("", s)
    s = re.sub(r"\s+", " ", s).strip("，,。；;：:、-— ")
    if len(s) > 24:
        s = s[:24]
    return s or "其他事项"


def _brief(text: str, limit: int = 110) -> str:
    """去噪并压缩到一句话/两句话的摘要。"""
    s = _URL_RE.sub("", text)
    s = _ID_RE.sub("", s)
    s = _PHONE_RE.sub("", s)
    s = _MD_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).strip("；;，,。 *")
    # 去掉对话式短句、纯疑问句
    parts = [p.strip() for p in re.split(r"[；;。！!？?]", s) if p.strip()]
    parts = [p for p in parts
             if not _DIALOG_RE.match(p) and len(p) >= 6
             and not (len(p) <= 12 and p.endswith(("么", "吗", "呢", "吧", "呀", "哈")))]
    s = "；".join(parts)
    if len(s) <= limit:
        return s
    cut = s[:limit]
    for sep in ("；", "，"):
        pos = cut.rfind(sep)
        if pos > limit // 2:
            return cut[:pos].strip("；；")
    return cut.rstrip() + "…"


def mine_ticket_rows(sheets: list[dict]) -> tuple[list[WorkItem], list[str]]:
    """工单/事项表结构化挖掘：名称=标题、解决方法=处理结果、状态列=分类。"""
    items: list[WorkItem] = []
    metrics: list[str] = []
    for sheet in sheets:
        m = _map_columns(sheet["headers"])
        if "name" not in m or not ("status" in m or "solution" in m):
            continue

        def cell(row: list[str], field: str) -> str:
            vals = [row[i] for i in m.get(field, []) if i < len(row) and row[i]]
            return "；".join(vals)

        for row in sheet["rows"]:
            name_raw = cell(row, "name")
            if not name_raw:
                continue
            name = _clean_title(name_raw.split("；")[0])
            solution = _brief(cell(row, "solution"), 120)
            desc = _brief(cell(row, "desc"), 60 if solution else 110)
            date_txt = cell(row, "date").split("；")[0]

            # 状态：优先取状态列，其次是否处理，最后动词猜测
            status = None
            for idx in m.get("status", []):
                if idx < len(row) and row[idx]:
                    status = _TICKET_STATUS_MAP.get(row[idx].lower())
                    if status:
                        break
            parts = [p for p in (desc, solution and f"处理方式：{solution}",) if p]
            text = "；".join(parts)
            if date_txt:
                text = f"{date_txt} {text}"
            if not text.strip() or len(text) < MIN_SEG_LEN:
                continue
            status = status or _classify(text) or "wip"
            it = WorkItem(
                text=text, status=status,
                candidates={name} if 3 <= len(name) <= 20 else set(),
                fallback=name, source=sheet["title"], has_metric=_is_metric(text),
            )
            items.append(it)
            if it.has_metric:
                metrics.append(text)
    return items, metrics


def mine_file(path: Path, text: str) -> tuple[list[WorkItem], list[str]]:
    """从单个文件文本中挖掘工作事项，返回 (items, metric_lines)。
    xlsx 若为工单/事项表，走结构化通道。"""
    if path.suffix.lower() == ".xlsx":
        try:
            from .extractors import extract_xlsx_sheets

            sheets = extract_xlsx_sheets(path)
            if any(is_ticket_table(s["headers"]) for s in sheets):
                t_items, t_metrics = mine_ticket_rows(sheets)
                if t_items:
                    return t_items, t_metrics
        except Exception:
            pass  # 退回通用文本挖掘

    items: list[WorkItem] = []
    metrics: list[str] = []
    fallback = _fallback_name(path)
    seen: set[str] = set()

    for seg, hint in _split_segments(text):
        status = hint or _classify(seg)
        if status is None:
            if _is_metric(seg) and 8 <= len(seg) <= 120 and " | " not in seg:
                metrics.append(seg)
            continue
        norm = _normalize(seg)
        if norm in seen:
            continue
        seen.add(norm)
        items.append(
            WorkItem(
                text=seg,
                status=status,
                candidates=candidate_names(seg),
                fallback=fallback,
                source=path.name,
                has_metric=_is_metric(seg),
            )
        )
        if _is_metric(seg) and len(seg) <= 120 and " | " not in seg:
            metrics.append(seg)
    return items, metrics


def group_by_project(items: list[WorkItem]) -> dict[str, list[WorkItem]]:
    """按项目分组，组内保持原始顺序。"""
    groups: dict[str, list[WorkItem]] = {}
    for it in items:
        groups.setdefault(it.project, []).append(it)
    return groups
