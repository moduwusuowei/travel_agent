# -*- coding: utf-8 -*-
"""旅行行程规划智能体 · 主图。

流程（DESIGN §4）：
  intake → research → outline → [human_review(interrupt)] → plan_days → feasibility → spot_intro → write_report
                          ↑__有修改意见__|                        |__不可行且未到上限__|
四大进阶特性：
  · Human-in-the-loop：human_review 用 interrupt() 抛大纲，CLI 展示后 Command(resume=文本) 回填；
  · 逐日子图：plan_days 逐天 .invoke() 编译好的 day_subgraph；
  · 反思循环：feasibility 结构化判定，不可行且 iteration<MAX_ITER 回 plan_days（带按天 fixes）；
  · Markdown→Word：write_report 出 md，CLI --docx 经 subprocess 调同目录 md2docx.py 导出。
另含景点介绍：spot_intro 为行程内景点并发生成 200~500 字介绍，写入报告「景点介绍」节（--no-intro 关）。
运行：
  python agent.py -q "北京3天，喜欢历史和美食，预算3000，两个人"
  python agent.py                       # 交互模式
  python agent.py -q "上海2天" --no-web   # 离线桩自测
  python agent.py -q "成都3天 美食为主" --docx --auto-approve
"""
from __future__ import annotations
import argparse
import os
import subprocess
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor
from typing import Annotated, TypedDict

from langchain_core.messages import AIMessage, HumanMessage
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langgraph.checkpoint.memory import MemorySaver
from langgraph.types import interrupt, Command

import llm
from schemas import TripSpec, DayOutline, TripOutline, DayPlan, Feasibility
from day_subgraph import build_day_graph
import report as report_mod
from tools import web_search

HERE = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(HERE, "dist")   # 所有生成物（md / docx）统一落到此目录
MAX_ITER = 2                       # 反思回炉硬上限
CATEGORIES = ["景点", "天气", "交通", "美食"]
# 每类给更聚焦的检索词（比一律「XX 推荐」召回更准，也避免「天气 近期天气」这种叠字）。
_QUERY_HINT = {"景点": "必去景点 攻略", "天气": "近期天气 穿衣建议",
               "交通": "市内交通 地铁 打车", "美食": "特色美食 老字号 推荐"}
_DAY_GRAPH = build_day_graph()     # 编译一次，plan_days 反复调用


# ============================ State ============================
class TripState(TypedDict, total=False):
    messages: Annotated[list, add_messages]
    request: str
    spec: TripSpec
    research: dict
    outline: list        # list[DayOutline]
    approved: bool       # 大纲人工审核是否通过
    feedback: str        # 审核意见（空=批准）
    days: list           # list[DayPlan]
    issues: list         # 可行性问题（给人看）
    fixes: list          # list[DayFix]，回炉时按天重排提示
    feasible: bool       # 最近一次反思结论
    iteration: int       # 反思轮次
    spot_intros: dict    # {景点名: 200~500字介绍}
    report_md: str
    use_web: bool
    use_intro: bool      # 是否生成景点介绍（--no-intro 关）


# ============================ 节点 ============================
def intake_node(state: TripState) -> dict:
    req = state["request"]
    spec = llm._structured(
        TripSpec,
        "你是旅行需求解析助手。从用户一句话里抽取结构化出行需求；缺失项留默认/空，别编造预算。",
        f"用户需求：{req}\n抽取 目的地/天数/出发地/预算/兴趣/节奏/人群/出发日期。",
        fallback=TripSpec(destination=(req[:20] or "未知目的地"), days=3),
        max_tokens=1200,
    )
    if spec.days < 1:
        spec.days = 1
    print(f"① 解析需求 → {spec.destination} / {spec.days}天 / 兴趣={spec.interests} / 预算={spec.budget_cny}")
    return {"spec": spec, "approved": False, "feedback": "", "iteration": 0,
            "days": [], "issues": [], "fixes": []}


def research_node(state: TripState) -> dict:
    """四类资料并发检索：各类相互独立，用线程池同时跑，避免 4 次串行阻塞（每次可达 90s 超时）。
    web_search 自带降级不抛异常；ex.map 保序，research 的插入顺序与 CATEGORIES 一致（结果确定）。"""
    spec = state["spec"]
    use_web = state.get("use_web", True)

    def _one(cat):
        q = f"{spec.destination} {_QUERY_HINT.get(cat, cat + ' 推荐')}"
        return cat, web_search(q, use_web=use_web, category=cat)

    research = {}
    with ThreadPoolExecutor(max_workers=len(CATEGORIES)) as ex:
        for cat, text in ex.map(_one, CATEGORIES):
            research[cat] = text
    print(f"② 检索资料（{'联网·并发' if use_web else '离线桩'}）：{', '.join(CATEGORIES)}")
    return {"research": research}


def _outline_user(spec, research, feedback):
    r = "\n".join(f"【{k}】{v}" for k, v in research.items())
    base = (f"目的地：{spec.destination}｜天数：{spec.days}｜兴趣：{'、'.join(spec.interests) or '不限'}｜"
            f"节奏：{spec.pace}｜人群：{spec.party}｜预算：{spec.budget_cny or '不限'}\n"
            f"实时资料：\n{r}\n"
            f"请生成 {spec.days} 天的大纲，每天一个 DayOutline(day/theme/area/candidate_spots)，"
            f"day 从 1 连续到 {spec.days}，候选景点每天 3~5 个且各天不重复。")
    if feedback:
        base += f"\n【上一版的修改意见，请据此调整】：{feedback}"
    return base


def outline_node(state: TripState) -> dict:
    spec = state["spec"]
    fb = state.get("feedback", "")
    fallback = TripOutline(days=[DayOutline(day=i + 1, theme=f"第{i + 1}天", area="",
                                            candidate_spots=[]) for i in range(spec.days)])
    res = llm._structured(
        TripOutline,
        "你是行程大纲规划师。产出逐日主题与候选景点，覆盖用户兴趣、就近成片、节奏合理。",
        _outline_user(spec, state.get("research", {}), fb),
        fallback=fallback,
        max_tokens=2500,
    )
    days = list(res.days or fallback.days)[:spec.days]
    for i, d in enumerate(days):        # 规整 day 编号
        d.day = i + 1
    while len(days) < spec.days:        # 补齐到 spec.days 天
        days.append(DayOutline(day=len(days) + 1, theme=f"第{len(days) + 1}天",
                               area="", candidate_spots=[]))
    print("③ 生成大纲：" + "；".join(f"D{d.day}{d.theme}" for d in days))
    return {"outline": days}


def human_review_node(state: TripState) -> dict:
    """HITL：interrupt 抛出大纲（转成 plain dict 便于序列化），resume 回填的字符串为审核意见。"""
    outline = state.get("outline", [])
    payload = {"outline": [d.model_dump() for d in outline],
               "hint": "回车=批准；或输入修改意见让我重排大纲。"}
    feedback = interrupt(payload)
    fb = (feedback or "").strip()
    if fb:
        print(f"④ 审核 → 收到修改意见：{fb}")
        return {"approved": False, "feedback": fb}
    print("④ 审核 → 已批准")
    return {"approved": True, "feedback": ""}


def route_after_review(state: TripState) -> str:
    return "plan_days" if state.get("approved") else "outline"


def plan_days_node(state: TripState) -> dict:
    """逐天调用子图；顺序执行让后一天看到前一天，避免重复。回炉时按天注入 fix_hint。"""
    spec = state["spec"]
    outline = state["outline"]
    research = state.get("research", {})
    fixes = state.get("fixes", [])
    fix_by_day = {f.day: f.hint for f in fixes}
    spec_json = spec.model_dump()
    done, plans = [], []
    for o in outline:
        sub = _DAY_GRAPH.invoke({
            "day": o.day, "outline": o.model_dump(), "spec": spec_json,
            "research": research, "prev_days": done, "fix_hint": fix_by_day.get(o.day, ""),
        })
        plan = sub["day_plan"]
        plans.append(plan)
        done.append({"day": o.day, "spots": [s.spot for s in plan.slots if s.spot]})
    print(f"⑤ 逐日规划完成：{len(plans)} 天" + ("（含回炉调整）" if fixes else ""))
    return {"days": plans}


def _feas_user(spec, days):
    lines = []
    for d in days:
        lines.append(f"第{d.day}天({d.theme}) 预算{d.est_cost}元：" +
                     "；".join(f"{s.time}{s.spot or s.activity}" for s in d.slots))
    return (f"预算总额：{spec.budget_cny or '不限'}｜节奏：{spec.pace}\n" + "\n".join(lines) +
            "\n判断：单日是否过满/太松、是否超总预算、开放时间/动线是否明显冲突。"
            "有问题就给 fixes（按天 hint）；都合理则 ok=true。")


def feasibility_node(state: TripState) -> dict:
    spec = state["spec"]
    days = state["days"]
    it = state.get("iteration", 0) + 1
    feas = llm._structured(
        Feasibility,
        "你是行程可行性审校员。只挑真正影响体验的问题，别吹毛求疵。",
        _feas_user(spec, days),
        fallback=Feasibility(ok=True, issues=[], fixes=[]),
        max_tokens=1500,
    )
    if feas.ok or it >= MAX_ITER:
        print(f"⑥ 反思 → {'通过' if feas.ok else f'到上限{MAX_ITER}轮'}，去写报告")
    else:
        print(f"⑥ 反思 → 第{it}轮发现问题：{feas.issues}，回炉重排")
    return {"iteration": it, "issues": feas.issues, "fixes": feas.fixes, "feasible": feas.ok}


def route_after_feasibility(state: TripState) -> str:
    if state.get("feasible") or state.get("iteration", 0) >= MAX_ITER:
        return "spot_intro"
    return "plan_days"


def _clip_intro(text, hi=500):
    """介绍过长则在靠后的句末标点处截到 ≤hi 字，保证不超上限；顺带压掉多余空白/换行成一段。"""
    text = " ".join(str(text).split())
    if len(text) <= hi:
        return text
    cut = text[:hi]
    for p in "。！？.!?":
        idx = cut.rfind(p)
        if idx >= 150:                 # 有靠后的句末标点就断在那，避免半句
            return cut[:idx + 1]
    return cut


def _intro_one(spec, spot, research):
    """给单个景点生成 200~500 字中文介绍（free-text，带兜底）；失败/空 → 返回空串。"""
    try:
        msg = llm.get_llm(max_tokens=1200).invoke([
            ("system", "你是熟悉当地的向导。为给定景点写一段 200~500 字的中文介绍，涵盖看点特色、"
                       "适合人群、游玩与注意建议；客观简洁、可读性强，不要编造门票价格与营业时间。"
                       "只输出介绍正文，不要标题、不要分点。"),
            ("user", f"城市：{spec.destination}；景点：{spot}。"
                     f"参考资料（可借鉴，勿照抄）：{str(research.get('景点', ''))[:300]}\n"
                     f"请写 200~500 字介绍。"),
        ])
        return _clip_intro(llm.text_of(msg).strip())
    except Exception:
        return ""


def spot_intro_node(state: TripState) -> dict:
    """为行程里出现的景点并发生成 200~500 字介绍（去重保序）。各景点相互独立，用线程池并发；
    单个失败只降级该条，不影响其余。--no-intro 时整节跳过。"""
    if not state.get("use_intro", True):
        print("⑦ 景点介绍：已跳过（--no-intro）")
        return {"spot_intros": {}}
    spec = state["spec"]
    days = state.get("days", [])
    research = state.get("research", {})
    seen, spots = set(), []            # 去重保序：只收非空 spot
    for d in days:
        for s in d.slots:
            name = (s.spot or "").strip()
            if name and name not in seen:
                seen.add(name)
                spots.append(name)
    intros = {}
    if spots:
        with ThreadPoolExecutor(max_workers=min(5, len(spots))) as ex:
            for name, text in zip(spots, ex.map(lambda n: _intro_one(spec, n, research), spots)):
                if text:
                    intros[name] = text
    print(f"⑦ 景点介绍：{len(intros)}/{len(spots)} 个景点已生成（200~500字/个）")
    return {"spot_intros": intros}


def generate_tips(spec, research, days):
    """一段实用贴士（free-text，带兜底）。失败/空返回空串，交给 report 用默认贴士。"""
    try:
        msg = llm.get_llm(max_tokens=600).invoke([
            ("system", "你是贴心旅行顾问，输出 3~5 条实用贴士，每条一行，以「·」开头，简洁具体。"),
            ("user", f"目的地{spec.destination}，{spec.days}天，人群{spec.party}，"
                     f"兴趣{'、'.join(spec.interests) or '不限'}。"
                     f"天气参考：{research.get('天气', '')}\n只输出贴士本身。"),
        ])
        return llm.text_of(msg).strip()
    except Exception:
        return ""


def write_report_node(state: TripState) -> dict:
    spec = state["spec"]
    days = state["days"]
    research = state.get("research", {})
    tips = generate_tips(spec, research, days)
    md = report_mod.render_markdown(spec, research, days, tips=tips,
                                    spot_intros=state.get("spot_intros", {}))
    print(f"⑧ 生成报告：{len(md)} 字")
    return {"report_md": md,
            "messages": [AIMessage(f"已生成《{spec.destination}·{spec.days}天行程》")]}


# ============================ 组装图 ============================
def build_trip_app(checkpointer=None):
    """编译主图。interrupt 需要 checkpointer，默认给 MemorySaver。"""
    g = StateGraph(TripState)
    g.add_node("intake", intake_node)
    g.add_node("research", research_node)
    g.add_node("outline", outline_node)
    g.add_node("human_review", human_review_node)
    g.add_node("plan_days", plan_days_node)
    g.add_node("feasibility", feasibility_node)
    g.add_node("spot_intro", spot_intro_node)
    g.add_node("write_report", write_report_node)

    g.add_edge(START, "intake")
    g.add_edge("intake", "research")
    g.add_edge("research", "outline")
    g.add_edge("outline", "human_review")
    g.add_conditional_edges("human_review", route_after_review,
                            {"plan_days": "plan_days", "outline": "outline"})
    g.add_edge("plan_days", "feasibility")
    g.add_conditional_edges("feasibility", route_after_feasibility,
                            {"spot_intro": "spot_intro", "plan_days": "plan_days"})
    g.add_edge("spot_intro", "write_report")
    g.add_edge("write_report", END)
    return g.compile(checkpointer=checkpointer or MemorySaver())


# ============================ 运行（流式 + HITL）============================
def _print_outline(payload):
    print("\n===== 请审核行程大纲 =====")
    for d in payload.get("outline", []):
        spots = "、".join(d.get("candidate_spots", []))
        print(f"  第{d.get('day')}天 [{d.get('theme', '')}] {d.get('area', '')}：{spots}")
    print(payload.get("hint", ""))


def run_trip(app, request, thread_id="cli", use_web=True, auto_approve=False, use_intro=True):
    """跑一次并流式展示；遇 interrupt 收一行审核意见后 Command(resume=...) 继续。返回最终 report_md。"""
    config = {"configurable": {"thread_id": thread_id}}
    stream_input = {"messages": [HumanMessage(request)], "request": request,
                    "use_web": use_web, "use_intro": use_intro}
    report_md = ""
    while True:
        interrupted = False
        try:
            for upd in app.stream(stream_input, config, stream_mode="updates"):
                if "__interrupt__" in upd:
                    interrupted = True
                    _print_outline(upd["__interrupt__"][0].value)
                    if auto_approve:
                        fb = ""
                        print("（--auto-approve：自动批准）")
                    else:
                        try:
                            fb = input("审核> 回车=批准 / 输入修改意见：").strip()
                        except (EOFError, KeyboardInterrupt):
                            fb = ""
                    stream_input = Command(resume=fb)
                    break
                for node, out in upd.items():
                    if node == "write_report" and out:
                        report_md = out.get("report_md", report_md)
        except Exception as e:
            if os.environ.get("TRAVEL_DEBUG"):   # 调试时看完整堆栈，平时只提示降级
                traceback.print_exc()
            print(f"（流式降级：{type(e).__name__}）")
            break
        if not interrupted:
            break
    if not report_md:                     # 兜底：从 checkpoint 里取最终态
        try:
            report_md = app.get_state(config).values.get("report_md", "")
        except Exception:
            pass
    return report_md


def export_docx(md_path):
    """复用同目录 md2docx.py 把 md 导出 Word。缺文件/依赖/失败 → 打印提示并返回 False。
    docx 落在 md 同目录（md 已归一到 dist/，故 docx 也在 dist/）。"""
    md2 = os.path.join(HERE, "md2docx.py")
    if not os.path.exists(md2):
        print(f"（跳过 Word 导出：未找到 {md2}）")
        return False
    docx_path = os.path.splitext(md_path)[0] + ".docx"
    try:
        subprocess.run([sys.executable, md2, md_path, "-o", docx_path], check=True)
        print(f"已导出 Word：{docx_path}")
        return True
    except Exception as e:
        print(f"（Word 导出失败：{type(e).__name__}，Markdown 已在 {md_path}）")
        return False


def _resolve_out(path):
    """把输出路径归一到 dist/：纯文件名放进 dist/；绝对路径或自带目录的相对路径按用户给的来。
    顺带把目标目录建好，避免 open() 因目录不存在而失败。"""
    out = path if (os.path.isabs(path) or os.path.dirname(path)) else os.path.join(DIST_DIR, path)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    return out


def _default_out(app, thread_id):
    try:
        spec = app.get_state({"configurable": {"thread_id": thread_id}}).values.get("spec")
        if spec:
            return f"行程_{spec.destination}_{spec.days}天.md"
    except Exception:
        pass
    return "行程.md"


def main():
    ap = argparse.ArgumentParser(description="旅行行程规划智能体")
    ap.add_argument("-q", "--question", default="", help="旅行需求；不给则进交互模式")
    ap.add_argument("--no-web", action="store_true", help="关闭联网，用内置桩离线跑")
    ap.add_argument("--no-intro", action="store_true", help="不生成景点介绍（更快）")
    ap.add_argument("--docx", action="store_true", help="跑完导出 Word（复用同目录 md2docx.py）")
    ap.add_argument("--auto-approve", action="store_true", help="跳过人工审核（脚本化/测试）")
    ap.add_argument("-o", "--out", default="",
                    help="报告输出路径；纯文件名默认落到 dist/（默认 行程_目的地_N天.md）")
    args = ap.parse_args()

    use_web = not args.no_web
    use_intro = not args.no_intro
    app = build_trip_app()          # 内部默认 MemorySaver（interrupt 需要 checkpointer）

    def once(req, tid):
        md = run_trip(app, req, thread_id=tid, use_web=use_web,
                      auto_approve=args.auto_approve, use_intro=use_intro)
        if not md:
            print("（未生成报告）")
            return
        out = _resolve_out(args.out or _default_out(app, tid))
        with open(out, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"\n✅ 已写入：{out}")
        if args.docx:
            export_docx(out)

    if args.question:
        once(args.question, "cli")
    else:
        print("进入交互模式（输入 q 退出）")
        i = 0
        while True:
            try:
                q = input("\n旅行需求> ").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if q.lower() in ("q", "quit", "exit"):
                break
            if q:
                i += 1
                once(q, f"chat-{i}")
    print("[travel_agent 完成]")


if __name__ == "__main__":
    main()
