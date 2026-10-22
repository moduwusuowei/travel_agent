# -*- coding: utf-8 -*-
"""TinyFish 联网检索封装 + 离线桩。对外只暴露 web_search(query, use_web, category)。

要点（复用 demo1/v2_agent.py、sh_news.py 的做法）：
  · Windows 上 tinyfish 是 .cmd，免 shell 的 list 调用跑不了，需经 cmd /c；
  · tinyfish 有时在 JSON 后追加一行升级提示 → 用 JSONDecoder().raw_decode 只取第一个 JSON；
  · 超时/无网/无 tinyfish 一律降级为一句说明字符串（不抛异常、不炸图）；
  · use_web=False → 返回内置桩数据（离线自测/断网降级）。"""
from __future__ import annotations
import json
import os
import shutil
import subprocess


def _tinyfish_cmd(sub_args):
    """拼出可执行的 tinyfish 命令列表；找不到可执行文件返回 None。"""
    exe = shutil.which("tinyfish")
    if not exe:
        return None
    if os.name == "nt" and exe.lower().endswith((".cmd", ".bat")):
        return ["cmd", "/c", exe, *sub_args]
    return [exe, *sub_args]


def _parse_search_json(raw):
    """从 tinyfish search 原始 stdout 提取 results，拼成「- 标题｜摘要（url）」多行。
    无 JSON → 返回空串；JSON 解析失败 → 返回截断原文。"""
    i = raw.find("{")
    if i < 0:
        return ""
    try:
        obj, _ = json.JSONDecoder().raw_decode(raw[i:])
    except Exception:
        return raw[:1500].strip()
    items = (obj.get("results") or [])[:5]
    lines = [f"- {it.get('title', '')}｜{it.get('snippet', '')}（{it.get('url', '')}）"
             for it in items if it.get("title") or it.get("snippet")]
    return "\n".join(lines)


# 离线桩：按类别给一段确定的示例文本，供 --no-web / 降级使用
_STUB = {
    "景点": "- 示例地标A｜必去打卡（stub）\n- 示例景点B｜小众文艺（stub）",
    "天气": "- 近期晴到多云为主，气温适宜，注意早晚温差（stub）",
    "交通": "- 市内地铁覆盖广，景点间打车约 15-30 分钟（stub）",
    "美食": "- 本地招牌小吃与老字号餐馆若干（stub）",
}


def web_search(query, use_web=True, category=""):
    """联网检索 query；use_web=False 用内置桩。返回适合喂 LLM 的多行文本或一句降级说明。"""
    if not use_web:
        return _STUB.get(category, f"（离线桩）关于「{query}」的示例信息")
    cmd = _tinyfish_cmd(["search", "query", query])
    if not cmd:
        return "(联网检索不可用：未找到 tinyfish)"
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=90)
        out = (r.stdout or "").strip()
        if not out:
            return "(联网无结果)"
        parsed = _parse_search_json(out)
        return parsed or "(联网无结果)"
    except Exception as e:  # 超时/网络等一律降级，不让联网拖垮整张图
        return f"(联网检索失败：{type(e).__name__})"
