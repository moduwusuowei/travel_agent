# -*- coding: utf-8 -*-
"""把结构化行程渲染成 Markdown（总览表 + 逐日明细 + 预算合计 + mermaid 时间轴 + 实用贴士）。
纯函数、无副作用、不联网、不 import 运行期依赖，故可完全离线单测。
接收 schemas 里的 TripSpec / DayPlan 对象（只读属性），不在此 import schemas。"""
from __future__ import annotations


def _cell(x):
    """单元格文本：把竖线替换成全角，避免破坏 Markdown 表格。"""
    return str(x).replace("|", "／")


# mermaid 节点标签里若混入这些字符会破坏 D1S0["..."] 语法（下游 md2docx 渲染时更明显）：
# 双引号收成单引号、圆括号转全角、方括号/井号/反引号直接删掉。
_MERMAID_BAD = str.maketrans({'"': "'", "(": "（", ")": "）",
                              "[": "", "]": "", "#": "", "`": "", "\n": " ", "\r": " "})


def _safe_label(text):
    """清洗 mermaid 节点标签：去掉破坏语法的字符、压掉多余空白、截断到 12 字。"""
    return " ".join(str(text).translate(_MERMAID_BAD).split())[:12]


def _overview_table(days):
    lines = ["| 天 | 主题 | 时段数 | 预计花费(元) |", "|---|---|---|---|"]
    for d in days:
        lines.append(f"| 第{d.day}天 | {_cell(d.theme or '—')} | {len(d.slots)} | {d.est_cost} |")
    return "\n".join(lines)


def _day_section(d):
    out = [f"### 第{d.day}天 · {d.theme}".rstrip(" ·"), ""]
    out += ["| 时段 | 活动 | 地点 | 交通 | 餐饮 | 备注 |", "|---|---|---|---|---|---|"]
    for s in d.slots:
        row = [s.time, s.activity, s.spot or "—", s.transit or "—", s.meal or "—", s.note or "—"]
        out.append("| " + " | ".join(_cell(x) for x in row) + " |")
    out += ["", f"**当日预计花费：约 {d.est_cost} 元**", ""]
    return "\n".join(out)


def _mermaid_timeline(days):
    """逐日把时段串成一条 flowchart：D<天>S<序号> --> 下一个。节点标签取地点/活动。"""
    lines = ["```mermaid", "flowchart LR"]
    for d in days:
        prev = None
        for idx, s in enumerate(d.slots):
            nid = f"D{d.day}S{idx}"
            label = _safe_label(s.spot or s.activity or s.time)
            lines.append(f'    {nid}["{label}"]')
            if prev:
                lines.append(f"    {prev} --> {nid}")
            prev = nid
    lines.append("```")
    return "\n".join(lines)


def _is_real_research(v):
    """判断某类检索结果是否是「真材料」：非空、且不是离线桩 / 降级说明。"""
    if not v:
        return False
    v = str(v).strip()
    return "stub" not in v and not v.startswith("(") and not v.startswith("（离线桩）")


def _research_appendix(research):
    """把真实的检索资料汇成一节「资料参考」；全是桩/降级/空 → 返回空串（不进报告）。"""
    real = [(k, str(v).strip()) for k, v in (research or {}).items() if _is_real_research(v)]
    if not real:
        return ""
    out = ["## 资料参考\n", "> 以下为实时检索所得，供参考核对（时间敏感，出行前请再确认）。\n"]
    for k, v in real:
        out.append(f"**{k}**\n")
        out.append(v)
        out.append("")
    return "\n".join(out).rstrip()


def _spot_intros_section(spot_intros):
    """把景点介绍汇成一节「景点介绍」（每个景点一个三级小节）；全空 → 返回空串（不进报告）。"""
    items = [(k, str(v).strip()) for k, v in (spot_intros or {}).items() if str(v).strip()]
    if not items:
        return ""
    out = ["## 景点介绍\n", "> 以下为各景点的简要介绍，供出行前了解。\n"]
    for name, intro in items:
        out.append(f"### {name}")
        out.append(intro)
        out.append("")
    return "\n".join(out).rstrip()


def render_markdown(spec, research, days, tips="", spot_intros=None):
    """spec: TripSpec；research: dict（此处仅备用，不强依赖）；days: list[DayPlan]；
    tips: 贴士文本；spot_intros: {景点名: 介绍文本}，空则不渲染「景点介绍」节。"""
    total = sum(d.est_cost for d in days)
    parts = [f"# {spec.destination}·{spec.days}天行程"]

    meta = [f"出发地：{spec.origin or '—'}", f"人群：{spec.party}",
            f"节奏：{spec.pace}", f"兴趣：{'、'.join(spec.interests) or '—'}"]
    if spec.budget_cny:
        meta.append(f"预算：{spec.budget_cny} 元")
    if spec.start_date:
        meta.append(f"出发日期：{spec.start_date}")
    parts.append("> " + " ｜ ".join(meta))

    parts.append("\n## 总览\n")
    parts.append(_overview_table(days))
    parts.append(f"\n**全程预计花费：约 {total} 元**"
                 + (f"（预算 {spec.budget_cny} 元）" if spec.budget_cny else ""))

    parts.append("\n## 逐日明细\n")
    for d in days:
        parts.append(_day_section(d))

    parts.append("## 路线时间轴\n")
    parts.append(_mermaid_timeline(days))

    intros = _spot_intros_section(spot_intros)
    if intros:
        parts.append("\n" + intros)

    parts.append("\n## 实用贴士\n")
    parts.append(tips.strip() or "· 提前查好各景点开放时间与门票预约，错峰出行更从容。")

    appendix = _research_appendix(research)
    if appendix:
        parts.append("\n" + appendix)

    return "\n".join(parts).rstrip() + "\n"
