# -*- coding: utf-8 -*-
"""travel_agent 自带的模型层：get_llm / text_of / _structured。

抽这一层的三个本机坑（同 demo1/common.py）：
  1) opus-5 会强制 thinking 块 → .content 变列表；默认用 claude-opus-4-8 且显式关 thinking，
     并用 text_of() 兜底任何要做字符串操作的地方。
  2) 低 max_tokens 会被思考吃空 → 默认给足，可按需 max_tokens=... 覆盖。
  3) 结构化输出偶发抽风 → _structured() 一律 try/except 返回 fallback，保证图不中断。
自带一份、读本目录 .env，开箱即用、不依赖外部目录。"""
from __future__ import annotations
import os
import sys

try:  # Windows 控制台打印中文
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

from dotenv import load_dotenv

HERE = os.path.dirname(os.path.abspath(__file__))


def _load_env():
    """加载本目录下的 .env（不存在则从当前工作目录兜底找一次）。override 压过会话变量。"""
    p = os.path.join(HERE, ".env")
    if os.path.exists(p):
        load_dotenv(p, override=True)
        return os.path.abspath(p)
    load_dotenv(override=True)
    return None


ENV_PATH = _load_env()
DEFAULT_MODEL = os.environ.get("VERIFY_MODEL", "claude-opus-4-8")


def get_llm(**kwargs):
    """返回 ChatAnthropic（走本机代理，thinking 关闭 → content 是纯字符串）。
    需要更长回答传 max_tokens=...；换模型传 model=...。"""
    from langchain_anthropic import ChatAnthropic
    token = os.environ.get("ANTHROPIC_AUTH_TOKEN") or os.environ.get("ANTHROPIC_API_KEY")
    base = os.environ.get("ANTHROPIC_BASE_URL")
    if not (token and base):
        raise RuntimeError(
            f"缺少代理配置（ANTHROPIC_BASE_URL / ANTHROPIC_AUTH_TOKEN）。检查 .env：{ENV_PATH}")
    params = dict(
        model=kwargs.pop("model", DEFAULT_MODEL),
        api_key=token, base_url=base,
        max_tokens=1024, timeout=120,
        thinking={"type": "disabled"},
    )
    params.update(kwargs)
    return ChatAnthropic(**params)


def text_of(msg_or_content):
    """健壮读取文本：.content 可能是 str，也可能是 [thinking, text] 块列表。"""
    c = getattr(msg_or_content, "content", msg_or_content)
    if isinstance(c, str):
        return c
    if isinstance(c, list):
        out = []
        for b in c:
            if isinstance(b, dict) and b.get("type") == "text":
                out.append(b.get("text", ""))
            elif isinstance(b, str):
                out.append(b)
        return "".join(out)
    return str(c)


def _structured(model_cls, system, user, fallback, **llm_kwargs):
    """结构化输出 + 兜底。with_structured_output 走 tool-calling；抽风返回 fallback，保证图不中断。
    用消息元组直接 invoke（user 里可能含 {} 字面量，不走会当模板解析的 ChatPromptTemplate）。
    大输出可传 max_tokens=... 经 llm_kwargs 透传给 get_llm。"""
    try:
        llm = get_llm(**llm_kwargs).with_structured_output(model_cls)
        return llm.invoke([("system", system), ("user", user)])
    except Exception as e:
        print(f"   ⚠ 结构化输出兜底（{type(e).__name__}）")
        return fallback


if __name__ == "__main__":
    print("ENV_PATH =", ENV_PATH)
    print("MODEL    =", DEFAULT_MODEL)
    print("TEXT_OF  =", text_of([{"type": "text", "text": "hello"}]))
    llm = get_llm(max_tokens=64)
    print("SMOKE    =", text_of(llm.invoke([("user", "只回复两个字：收到")])))
