# -*- coding: utf-8 -*-
"""季度总结内容生成模块：由工作内容摘要生成模板所需的结构化内容。"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .miner import (WorkDigest, WorkItem, group_by_project, parse_period,
                    resolve_projects)

IMAGE_NOTE = "（此处配3－5幅具有代表性、有一定差异性的系统截图）"


def cn_number(n: int) -> str:
    """1 -> 一 ... 12 -> 十二"""
    digits = "零一二三四五六七八九"
    if n < 10:
        return digits[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + digits[n - 10]
    tens, ones = divmod(n, 10)
    return digits[tens] + "十" + (digits[ones] if ones else "")


@dataclass
class SectionItem:
    name: str
    paragraphs: list[str] = field(default_factory=list)


@dataclass
class SummaryContent:
    title: str
    overview: list[str]
    completed: list[SectionItem]
    wip: list[SectionItem]


def _quarter_label(quarter: str) -> str:
    """兼容旧调用：返回周期标签（如 2026年第三季度）。"""
    return parse_period(value=quarter).label


def _dedupe_keep_order(texts: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for t in texts:
        key = re.sub(r"\s+", "", t)
        if key in seen:
            continue
        seen.add(key)
        out.append(t)
    return out


def _compose_item(name: str, items: list[WorkItem], max_sentences: int = 8) -> SectionItem:
    """把同一项目下的工作条目组织成自然文字（保持时间顺序）。"""
    sentences = _dedupe_keep_order([it.text.strip() for it in items])[:max_sentences]
    paras: list[str] = []
    if sentences:
        body = "；".join(s.rstrip("。；;") for s in sentences) + "。"
        paras.append(body)
    else:
        paras.append(f"「{name}」相关工作详见工作记录。")
    return SectionItem(name=name, paragraphs=paras)


def _plan_clause(items: list[WorkItem]) -> str | None:
    """提取『预计/计划…』子句作为下一步计划。"""
    for it in items:
        m = re.search(r"(预计|计划)([^，。；,;]+)", it.text)
        if m:
            return (m.group(1) + m.group(2)).strip()
    return None


def generate_content(digest: WorkDigest, title: str | None = None,
                     max_completed: int | None = None,
                     max_wip: int | None = None) -> SummaryContent:
    """规则模式生成总结内容。max_* 为 None 表示不限制条目数量。"""
    period = parse_period(value=digest.period)
    label = period.label
    title = title or f"{period.label}{period.suffix}"

    # ---------- 项目名归一化（全局投票） ----------
    resolve_projects(digest.all_items)

    # ---------- 一、总体情况 ----------
    groups_done = group_by_project(digest.items_completed)
    groups_wip = group_by_project(digest.items_wip)
    overview: list[str] = []

    overview.append(
        f"{period.short}共梳理工作记录文件{len(digest.files)}份，沉淀工作事项{len(digest.all_items)}项，"
        f"其中已完成{len(digest.items_completed)}项、进行中{len(digest.items_wip)}项。"
    )
    if groups_done:
        launched = [p for p, its in groups_done.items()
                    if any(("上线" in it.text or "试运行" in it.text or "投入使用" in it.text) for it in its)]
        accepted = [p for p, its in groups_done.items()
                    if any(("验收" in it.text or "签订" in it.text or "合同" in it.text or "交付" in it.text) for it in its)]
        if launched:
            overview.append(f"{period.short}上线（含试运行）的系统/模块{len(launched)}个，分别为：{'、'.join(launched)}。")
        if accepted:
            overview.append(f"完成论证/合同签订/验收/交付{len(accepted)}项，分别为：{'、'.join(accepted)}。")
        other = [p for p in groups_done if p not in launched and p not in accepted]
        if other:
            overview.append(f"其他形式工作成果涉及{len(other)}个方向，分别为：{'、'.join(other)}。")
    if groups_wip:
        overview.append(f"在研/推进中的工作方向{len(groups_wip)}个，分别为：{'、'.join(groups_wip)}，均按计划推进。")

    # 数据亮点：挑最多2条带数字的简短记录
    highlights = [m for m in _dedupe_keep_order(digest.metric_lines) if 8 <= len(m) <= 80][:2]
    for h in highlights:
        overview.append(f"关键数据：{h.rstrip('。；;')}。")

    # ---------- 二、已完成事项 ----------
    completed: list[SectionItem] = []
    ranked = sorted(
        groups_done.items(),
        key=lambda kv: (-sum(1 for it in kv[1] if it.has_metric), -len(kv[1])),
    )
    for name, its in ranked[:max_completed if max_completed else len(ranked)]:
        completed.append(_compose_item(name, its))
    if not completed and not groups_wip:
        completed.append(SectionItem(f"{period.short}主要工作", [
            "（未从工作记录中识别出已完成事项，请根据实际情况补充。）"
        ]))

    # ---------- 三、进行中事项 ----------
    wip: list[SectionItem] = []
    for name, its in list(groups_wip.items())[:max_wip if max_wip else len(groups_wip)]:
        sec = _compose_item(name, its)
        plan = _plan_clause(its)
        if plan:
            sec.paragraphs.append(f"下一步计划：{plan.rstrip('。；;')}。")
        else:
            sec.paragraphs.append("当前按计划推进中，下一步继续跟进落实相关工作。")
        wip.append(sec)

    return SummaryContent(title=title, overview=overview, completed=completed, wip=wip)


def render_markdown(content: SummaryContent, digest: WorkDigest) -> str:
    """生成便于人工核对的 Markdown 摘要。"""
    lines = [f"# {content.title}", "", "## 一、总体情况", ""]
    lines += [f"- {p}" for p in content.overview]
    lines += ["", "## 二、已完成事项", ""]
    for i, sec in enumerate(content.completed, 1):
        lines.append(f"### （{cn_number(i)}）{sec.name}")
        lines.append("")
        for p in sec.paragraphs:
            lines.append(f"> {p}")
        lines.append("")
    if content.wip:
        lines += ["## 三、进行中事项", ""]
        for i, sec in enumerate(content.wip, len(content.completed) + 1):
            lines.append(f"### （{cn_number(i)}）{sec.name}")
            lines.append("")
            for p in sec.paragraphs:
                lines.append(f"> {p}")
            lines.append("")
    lines += ["---", "", "## 附：数据来源", ""]
    lines += [f"- 已解析文件：{'、'.join(digest.files) or '无'}"]
    if digest.skipped:
        lines += [f"- 跳过文件：{'、'.join(digest.skipped)}"]
    return "\n".join(lines)
