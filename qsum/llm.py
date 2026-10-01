# -*- coding: utf-8 -*-
"""AI 内容生成模块：调用大模型把工作记录润色为总结正文。

支持两种接入方式：
1. OpenAI 兼容接口（API Key）：DeepSeek / 通义 / Kimi / GPT 等均兼容
2. 本地 Ollama：无需 Key，默认 http://localhost:11434

配置来源（优先级：显式参数 > 环境变量 QSUM_API_BASE / QSUM_API_KEY / QSUM_MODEL）
"""
from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request

from .generator import SectionItem, SummaryContent
from .miner import WorkDigest, group_by_project, parse_period, resolve_projects


class LLMError(Exception):
    pass


DEFAULT_TIMEOUT = 1800  # 单次模型调用超时（秒）；本地大模型生成慢，预留 30 分钟
MAX_DIGEST_CHARS = 24000

SYSTEM_PROMPT = (
    "你是资深的企业总结报告撰写专家。你将收到某周期的工作记录（来自工作日志/周报/任务表），"
    "目标产物可能是日报/周报/月报/季度总结，请按目标周期的口径撰写结构化总结报告。硬性要求：\n"
    "1. 只能使用工作记录中的真实信息（系统/项目名称、时间、数字），严禁编造数据；\n"
    "2. 语言精炼、书面化，避免流水账：每个条目按『背景/工作内容/成效』组织，"
    "进行中事项需写明当前进展与下一步计划；\n"
    "3. 合并同义、重复的条目，突出亮点与量化成果；\n"
    "4. 严格输出 JSON（不要输出其它文字），字符串值中的英文双引号必须转义为 \"，不要使用尾逗号；结构：\n"
    '{"title":"<周期><报告类型>","overview":["总体情况段落1","段落2"],'
    '"completed":[{"name":"系统/项目名","paragraphs":["正文段落1","段落2"]}],'
    '"wip":[{"name":"系统/项目名","paragraphs":["正文段落","下一步计划段落"]}]}'
)


def _post_json(url: str, headers: dict, payload: dict, timeout: int) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json", **headers},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")[:400]
        raise LLMError(f"模型接口返回错误 HTTP {exc.code}: {detail}")
    except urllib.error.URLError as exc:
        raise LLMError(f"无法连接模型接口 {url}: {exc.reason}")
    except Exception as exc:  # 超时、JSON 解析等
        raise LLMError(f"调用模型接口失败: {exc}")


def chat(messages: list[dict], provider: str, api_base: str | None,
         api_key: str | None, model: str, timeout: int = DEFAULT_TIMEOUT) -> str:
    """调用模型对话接口，返回助手文本。"""
    if not model:
        raise LLMError("未指定模型名称（model）")
    provider = (provider or "openai").lower()

    if provider == "ollama":
        base = (api_base or os.environ.get("QSUM_API_BASE")
                or "http://localhost:11434").rstrip("/")
        data = _post_json(f"{base}/api/chat", {}, {
            "model": model, "messages": messages, "stream": False,
            "format": "json",  # 服务端保证输出合法 JSON
        }, timeout)
        try:
            return data["message"]["content"]
        except (KeyError, TypeError):
            raise LLMError(f"Ollama 响应格式异常: {str(data)[:300]}")

    # OpenAI 兼容接口
    base = (api_base or os.environ.get("QSUM_API_BASE")
            or "https://api.openai.com/v1").rstrip("/")
    key = api_key or os.environ.get("QSUM_API_KEY") or ""
    if not key:
        raise LLMError("未提供 API Key（也可通过环境变量 QSUM_API_KEY 配置）")
    data = _post_json(f"{base}/chat/completions",
                      {"Authorization": f"Bearer {key}"}, {
                          "model": model, "messages": messages,
                          "temperature": 0.4,
                      }, timeout)
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError):
        raise LLMError(f"接口响应格式异常: {str(data)[:300]}")


def build_digest_text(digest: WorkDigest) -> str:
    """把工作内容摘要整理为发给模型的文本。"""
    resolve_projects(digest.all_items)
    lines: list[str] = []
    done_groups = group_by_project(digest.items_completed)
    wip_groups = group_by_project(digest.items_wip)
    if done_groups:
        lines.append("## 已完成事项（按项目）")
        for name, items in done_groups.items():
            lines.append(f"### {name}")
            for it in items:
                lines.append(f"- [{it.source}] {it.text}")
    if wip_groups:
        lines.append("## 进行中事项（按项目）")
        for name, items in wip_groups.items():
            lines.append(f"### {name}")
            for it in items:
                lines.append(f"- [{it.source}] {it.text}")
    text = "\n".join(lines)
    if len(text) > MAX_DIGEST_CHARS:
        text = text[:MAX_DIGEST_CHARS] + "\n…（记录过多已截断）"
    return text


def _escape_inner_quotes(text: str) -> str:
    """修复字符串值内未转义的英文引号（小模型常见错误）。
    规则：引号后（跳过空白）若不是 , : } ] 或结尾，则视为字符串内引号并转义。"""
    out = []
    in_str = False
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and in_str and i + 1 < n:
            out.append(ch)
            out.append(text[i + 1])
            i += 2
            continue
        if ch == '"':
            if not in_str:
                in_str = True
                out.append(ch)
                i += 1
                continue
            j = i + 1
            while j < n and text[j] in " \t\r\n":
                j += 1
            if j >= n or text[j] in ",:}]":
                in_str = False
                out.append(ch)
            else:
                out.append('\\"')
            i += 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _repair_json(text: str) -> str:
    """常见 JSON 瑕疵修复：代码块、注释、尾逗号、未转义引号。"""
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"\s*```$", "", t)
    t = re.sub(r"^\s*//.*$", "", t, flags=re.M)
    t = re.sub(r",\s*([}\]])", r"\1", t)
    return _escape_inner_quotes(t)


def _extract_json(text: str) -> dict:
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"\s*```$", "", t)
    i, j = t.find("{"), t.rfind("}")
    if i < 0 or j <= i:
        raise LLMError("模型未返回 JSON 内容，请重试或更换模型")
    raw = t[i:j + 1]
    for candidate in (raw, _repair_json(raw)):
        try:
            obj = json.loads(candidate)
            if isinstance(obj, dict):
                return obj
        except json.JSONDecodeError:
            continue
    raise LLMError("模型返回的 JSON 解析失败（已尝试自动修复），请重试或更换指令遵循更好的模型")



def _as_paras(value) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        return [str(v) for v in value if str(v).strip()]
    return []


def _as_sections(value) -> list[SectionItem]:
    out: list[SectionItem] = []
    if not isinstance(value, list):
        return out
    for entry in value:
        if isinstance(entry, dict):
            name = str(entry.get("name") or "").strip()
            paras = _as_paras(entry.get("paragraphs"))
            if name and paras:
                out.append(SectionItem(name=name, paragraphs=paras))
        elif isinstance(entry, str) and entry.strip():
            out.append(SectionItem(name=entry.strip(), paragraphs=[""]))
    return out


def ai_generate_content(digest: WorkDigest, provider: str = "openai",
                        api_base: str | None = None, api_key: str | None = None,
                        model: str = "", title: str | None = None,
                        extra_prompt: str | None = None,
                        timeout: int = DEFAULT_TIMEOUT) -> SummaryContent:
    """调用大模型生成总结内容（失败抛 LLMError）。"""
    period = parse_period(value=digest.period)
    quarter_label = period.label
    user_prompt = (
        f"目标周期：{quarter_label}（{period.raw}）\n"
        f"报告类型：{period.suffix}\n"
        f"来源文件：{'、'.join(digest.files)}\n\n"
        f"工作记录：\n{build_digest_text(digest)}\n\n"
        "请生成{suffix}，严格按系统提示中的 JSON 结构输出。"
        "title 默认为「{label}{suffix}」，overview 2-4 段，"
        "completed/wip 条目数量不限、按重要性排序。".format(
            label=quarter_label, suffix=period.suffix)
    )
    if extra_prompt and extra_prompt.strip():
        user_prompt += ("\n\n【用户的额外生成要求】" + extra_prompt.strip()
                        + "\n（如与上述默认结构/风格冲突，以额外要求为准）")
    obj = None
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt}]
    llm_kwargs = {"provider": provider, "api_base": api_base, "model": model,
                  "timeout": timeout}
    llm_kwargs["api_key"] = api_key
    for attempt in range(2):
        reply = chat(messages, **llm_kwargs)
        try:
            obj = _extract_json(reply)
            if not _as_paras(obj.get("overview")) or not _as_sections(obj.get("completed")):
                raise LLMError("模型输出缺少 overview/completed 内容")
            break
        except LLMError:
            if attempt == 1:
                raise
            messages = messages + [
                {"role": "assistant", "content": reply[:4000]},
                {"role": "user", "content":
                 "你上次输出的 JSON 有语法错误。请重新输出，只输出合法 JSON："
                 "字符串内的英文双引号必须转义，不要尾逗号，不要输出任何 JSON 以外的文字。"},
            ]

    content = SummaryContent(
        title=(title or str(obj.get("title") or "").strip()
               or f"{period.label}{period.suffix}"),
        overview=_as_paras(obj.get("overview")),
        completed=_as_sections(obj.get("completed")),
        wip=_as_sections(obj.get("wip")),
    )
    if not content.overview or not content.completed:
        raise LLMError("模型输出缺少 overview/completed 内容，请重试或更换模型")
    return content
