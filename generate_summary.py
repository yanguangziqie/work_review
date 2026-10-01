#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""季度工作总结生成系统 - 命令行入口

用法示例:
    python3 generate_summary.py -i ./input -q 2026Q3
    python3 generate_summary.py -i 工作日志.md 周报.xlsx -q 2026Q3 -o output/总结.docx
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from qsum.extractors import extract_file, scan_inputs  # noqa: E402
from qsum.generator import generate_content, render_markdown  # noqa: E402
from qsum.llm import LLMError, ai_generate_content  # noqa: E402
from qsum.miner import (  # noqa: E402
    WorkDigest,
    file_in_period,
    filter_items,
    mine_file,
    parse_period,
    text_in_period,
)
from qsum.template_filler import fill_template  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(
        description="根据日常工作内容文件 + 模板，生成日报/周报/月报/季度总结",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    ap.add_argument("-i", "--input", nargs="+", required=True,
                    help="工作内容文件或目录（支持 docx/xlsx/txt/md/csv/pdf，目录递归扫描）")
    ap.add_argument("--type", dest="report_type",
                    choices=["day", "week", "month", "quarter"], default=None,
                    help="报告类型：日报/周报/月报/季度总结（默认按时间参数格式自动识别）")
    ap.add_argument("-p", "--period", default=None,
                    help="时间范围：2026-10-01（日）/2026W40（周）/2026-10（月）/2026Q3（季），默认当前")
    ap.add_argument("-q", "--quarter", default=None,
                    help="兼容旧用法，等价于 --type quarter --period <值>")
    ap.add_argument("-t", "--template", default=None,
                    help="模板路径（docx/xlsx/xls/txt/md/csv/pdf，输出格式跟随模板；"
                         "不传则按内置标准结构输出 docx）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出路径（默认 output/<周期标签><报告类型>，扩展名随模板格式）")
    ap.add_argument("--no-filter", action="store_true",
                    help="不做时间过滤，纳入全部输入文件与事项")
    ap.add_argument("--max-completed", type=int, default=None,
                    help="已完成事项最多条目数（默认不限制）")
    ap.add_argument("--max-wip", type=int, default=None,
                    help="进行中事项最多条目数（默认不限制）")
    ap.add_argument("--ai", action="store_true",
                    help="使用 AI 大模型生成总结内容（默认规则模式）")
    ap.add_argument("--provider", choices=["openai", "ollama"], default="openai",
                    help="AI 接入方式：openai=OpenAI 兼容接口（API Key），ollama=本地 Ollama")
    ap.add_argument("--api-base", default=None,
                    help="接口地址。openai 默认 https://api.openai.com/v1；"
                         "ollama 默认 http://localhost:11434")
    ap.add_argument("--api-key", default=None,
                    help="API Key（也可用环境变量 QSUM_API_KEY）")
    ap.add_argument("--model", default="",
                    help="模型名，如 gpt-4o-mini / deepseek-chat / qwen2.5:14b")
    ap.add_argument("--extra-prompt", default=None,
                    help="AI 生成的额外要求（自然语言，如“侧重写数据中台，语气正式”；"
                         "不填走默认提示词）")
    ap.add_argument("--title", default=None,
                    help="覆盖文档标题（默认：<季度>工作总结）")
    ap.add_argument("--no-digest", action="store_true",
                    help="不输出 Markdown 摘要文件")
    args = ap.parse_args()

    raw = args.period or args.quarter or None
    kind = args.report_type or ("quarter" if (args.quarter and not args.period) else None)
    try:
        period = parse_period(kind, raw)
    except ValueError as exc:
        print(f"[错误] {exc}")
        return 2

    # 1. 扫描输入
    files = scan_inputs(args.input)
    if not files:
        print("[错误] 未找到可解析的文件，请检查输入路径与文件格式")
        return 2
    print(f"[1/4] 扫描到 {len(files)} 个候选文件")

    # 2. 抽取 + 时间过滤 + 挖掘
    digest = WorkDigest(period=period.raw)
    for f in files:
        try:
            text = extract_file(f)
        except Exception as exc:
            digest.skipped.append(f"{f.name}（{exc}）")
            print(f"  [跳过] {f.name}: {exc}")
            continue
        if not args.no_filter and not file_in_period(f, text, period):
            digest.skipped.append(f"{f.name}（不属于 {period.label}）")
            print(f"  [过滤] {f.name}: 不属于 {period.label}")
            continue
        items, metrics = mine_file(f, text)
        if not args.no_filter:
            items = filter_items(items, period)
            metrics = [m for m in metrics if text_in_period(m, period) is not False]
        if not items and not metrics:
            digest.skipped.append(f"{f.name}（未识别到工作事项）")
            print(f"  [提示] {f.name}: 未识别到工作事项")
            continue
        digest.files.append(f.name)
        digest.items_completed.extend(i for i in items if i.status == "completed")
        digest.items_wip.extend(i for i in items if i.status == "wip")
        digest.metric_lines.extend(metrics)
        print(f"  [解析] {f.name}: 已完成 {sum(1 for i in items if i.status == 'completed')} 条 / "
              f"进行中 {sum(1 for i in items if i.status == 'wip')} 条")
    print(f"[2/4] 纳入 {len(digest.files)} 个文件，"
          f"已完成 {len(digest.items_completed)} 条 / 进行中 {len(digest.items_wip)} 条")

    if not digest.all_items:
        print("[错误] 未从输入中识别到任何工作事项，请检查内容或使用 --no-filter")
        return 2

    # 3. 生成内容（规则模式 或 AI 模式）
    if args.ai:
        print(f"[3/4] AI 生成中（provider={args.provider}, model={args.model}）…")
        try:
            content = ai_generate_content(
                digest, provider=args.provider, api_base=args.api_base,
                api_key=args.api_key, model=args.model, title=args.title,
                extra_prompt=args.extra_prompt,
            )
        except LLMError as exc:
            print(f"[错误] AI 生成失败: {exc}")
            print("      可检查 API Key/接口地址/模型名，或去掉 --ai 使用规则模式")
            return 2
    else:
        content = generate_content(
            digest, title=args.title,
            max_completed=args.max_completed, max_wip=args.max_wip,
        )
    print(f"[3/4] 内容生成完成：已完成条目 {len(content.completed)} 个 / "
          f"进行中条目 {len(content.wip)} 个")

    # 4. 填充模板并输出
    out = Path(args.output) if args.output else BASE_DIR / "output" / f"{period.label}{period.suffix}"
    if not out.is_absolute():
        out = Path.cwd() / out
    tpl = None
    if args.template:
        tpl = Path(args.template)
        if not tpl.exists():
            print(f"[错误] 模板不存在: {tpl}")
            return 2
    out = fill_template(tpl, content, out)
    print(f"[4/4] 已生成: {out}")

    if not args.no_digest:
        digest_path = out.with_name(out.stem + "_工作内容摘要.md")
        digest_path.write_text(render_markdown(content, digest), encoding="utf-8")
        print(f"      摘要(供人工核对): {digest_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
