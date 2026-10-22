# -*- coding: utf-8 -*-
"""逐日规划子图（DESIGN §6）：sequence_stops（排动线+时段）→ add_logistics（补交通/餐饮/预算）。
一张编译好的小图，被主图 plan_days 逐天 .invoke()；子状态 DayState 与主状态隔离，
这是「子图」特性的落点。所有 LLM 调用经 llm._structured（模块属性调用，便于测试 monkeypatch）。"""
from __future__ import annotations
from typing import TypedDict

from langgraph.graph import StateGraph, START, END

import llm
from schemas import DayPlan


class DayState(TypedDict, total=False):
    day: int
    outline: dict        # 当天 DayOutline.model_dump()
    spec: dict           # TripSpec.model_dump()
    research: dict
    prev_days: list      # 已排天 [{day, spots:[...]}]，供避免重复
    fix_hint: str        # 反思回炉时的调整提示
    draft: DayPlan       # sequence_stops 产出
    day_plan: DayPlan    # add_logistics 产出（最终）


def _fmt_prev(prev):
    if not prev:
        return "（无）"
    return "；".join(f"第{p.get('day')}天已含 {'、'.join(p.get('spots', []))}" for p in (prev or []))


def sequence_stops(state: DayState) -> dict:
    """把候选景点按就近排成一条动线，配上时段；先只填 time/activity/spot。"""
    o = state.get("outline", {})
    spec = state.get("spec", {})
    day = state.get("day", o.get("day", 1))
    fallback = DayPlan(day=day, theme=o.get("theme", ""), slots=[], est_cost=0)
    draft = llm._structured(
        DayPlan,
        "你是行程排布助手。把当天候选景点按地理就近排成一条顺畅动线，配上时段"
        "（上午/午餐/下午/傍晚/晚上），只填 time/activity/spot，交通与餐饮先留空。",
        f"目的地：{spec.get('destination', '')}｜节奏：{spec.get('pace', '适中')}\n"
        f"当天主题：{o.get('theme', '')}｜区域：{o.get('area', '')}\n"
        f"候选景点：{'、'.join(o.get('candidate_spots', [])) or '（自定）'}\n"
        f"已排过的天（避免重复景点）：{_fmt_prev(state.get('prev_days'))}\n"
        f"调整提示：{state.get('fix_hint', '') or '（无）'}\n"
        f"输出第 {day} 天的 DayPlan（day={day}）。",
        fallback=fallback,
        max_tokens=2500,
    )
    draft.day = day
    return {"draft": draft}


def add_logistics(state: DayState) -> dict:
    """在动线基础上补交通接驳/餐饮建议/备注，并估算当日预算 est_cost。"""
    draft = state.get("draft")
    spec = state.get("spec", {})
    o = state.get("outline", {})
    day = state.get("day", getattr(draft, "day", 1))
    fallback = draft or DayPlan(day=day, theme=o.get("theme", ""), slots=[], est_cost=0)
    draft_json = draft.model_dump() if draft else {}
    plan = llm._structured(
        DayPlan,
        "你是行程完善助手。在给定动线基础上补全每个时段的交通接驳(transit)、餐饮建议(meal)、"
        "必要备注(note，如门票/预约/开放时间)，并给出当日合计预算 est_cost（整数, 元）。"
        "保持 day、主题、时段顺序不变。",
        f"目的地：{spec.get('destination', '')}｜预算总额：{spec.get('budget_cny') or '不限'}\n"
        f"当天草稿：{draft_json}\n"
        f"调整提示：{state.get('fix_hint', '') or '（无）'}\n"
        f"输出完善后的第 {day} 天 DayPlan（day={day}）。",
        fallback=fallback,
        max_tokens=2500,
    )
    plan.day = day
    if not plan.theme:
        plan.theme = o.get("theme", "")
    return {"day_plan": plan}


def build_day_graph():
    """编译逐日子图。子图无 interrupt，故不需 checkpointer。"""
    g = StateGraph(DayState)
    g.add_node("sequence_stops", sequence_stops)
    g.add_node("add_logistics", add_logistics)
    g.add_edge(START, "sequence_stops")
    g.add_edge("sequence_stops", "add_logistics")
    g.add_edge("add_logistics", END)
    return g.compile()
