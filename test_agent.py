# -*- coding: utf-8 -*-
"""travel_agent 全部离线单测（不联网、不调真模型）。
约定：每个测试在函数内部 import 项目模块，避免尚未创建的模块在收集阶段报错。"""


# ============================ Task 1: llm.py ============================
def test_text_of_handles_str_and_blocks():
    import llm
    assert llm.text_of("hi") == "hi"
    blocks = [{"type": "text", "text": "a"},
              {"type": "thinking", "thinking": "x"},
              {"type": "text", "text": "b"}]
    assert llm.text_of(blocks) == "ab"

    class M:
        content = [{"type": "text", "text": "z"}]
    assert llm.text_of(M()) == "z"


def test_structured_returns_fallback_on_error(monkeypatch):
    import llm

    def boom(**k):
        raise RuntimeError("no proxy")

    monkeypatch.setattr(llm, "get_llm", boom)
    fb = {"ok": True}
    assert llm._structured(dict, "sys", "user", fb) is fb


# ============================ Task 2: schemas.py ============================
def test_tripspec_defaults():
    from schemas import TripSpec
    s = TripSpec(destination="北京", days=3)
    assert s.origin is None and s.budget_cny is None and s.start_date is None
    assert s.interests == [] and s.pace == "适中" and s.party == "1人"


def test_nested_models_construct():
    from schemas import DayOutline, TripOutline, Slot, DayPlan, DayFix, Feasibility
    o = TripOutline(days=[DayOutline(day=1, theme="古都", area="东城",
                                     candidate_spots=["故宫", "天安门"])])
    assert o.days[0].candidate_spots == ["故宫", "天安门"]
    p = DayPlan(day=1, theme="古都",
                slots=[Slot(time="上午", activity="逛故宫", spot="故宫")], est_cost=120)
    assert p.slots[0].spot == "故宫" and p.est_cost == 120
    f = Feasibility(ok=False, issues=["太满"], fixes=[DayFix(day=1, hint="减一个点")])
    assert f.ok is False and f.fixes[0].day == 1


def test_slot_and_dayplan_optional_defaults():
    from schemas import Slot, DayPlan
    s = Slot(time="上午", activity="集合")
    assert s.spot == "" and s.transit == "" and s.meal == "" and s.note == ""
    d = DayPlan(day=1)
    assert d.theme == "" and d.slots == [] and d.est_cost == 0


# ============================ Task 3: tools.py ============================
def test_parse_search_json_extracts_lines():
    import tools
    raw = ('{"results":[{"title":"故宫","snippet":"必去","url":"http://x"},'
           '{"title":"","snippet":""}]}\nUpgrade notice trailing line')
    out = tools._parse_search_json(raw)
    assert "故宫" in out and "必去" in out and "http://x" in out
    assert out.count("\n") == 0  # 第二条空标题被过滤，只剩 1 行


def test_parse_search_json_no_json_returns_empty():
    import tools
    assert tools._parse_search_json("no json here") == ""


def test_web_search_stub_offline():
    import tools
    hit = tools.web_search("北京 景点 推荐", use_web=False, category="景点")
    assert "stub" in hit
    assert tools.web_search("任意", use_web=False)  # 无 category 也返回非空说明


def test_web_search_missing_tinyfish_degrades(monkeypatch):
    import tools
    monkeypatch.setattr(tools.shutil, "which", lambda name: None)
    out = tools.web_search("北京 天气", use_web=True, category="天气")
    assert out.startswith("(联网检索不可用")  # 无 tinyfish → 降级说明，不抛异常


# ============================ Task 4: report.py ============================
def _sample_spec_days():
    from schemas import TripSpec, Slot, DayPlan
    spec = TripSpec(destination="北京", days=1, interests=["历史"], budget_cny=1000, party="2人")
    days = [DayPlan(day=1, theme="古都",
                    slots=[Slot(time="上午", activity="逛故宫", spot="故宫", transit="地铁"),
                           Slot(time="午餐", activity="吃烤鸭", spot="全聚德", meal="烤鸭")],
                    est_cost=300)]
    return spec, days


def test_render_contains_all_sections():
    import report
    spec, days = _sample_spec_days()
    md = report.render_markdown(spec, {}, days, tips="· 早晚温差大，带外套")
    assert "# 北京·1天行程" in md
    assert "## 总览" in md and "| 第1天 |" in md
    assert "第1天" in md and "故宫" in md and "全聚德" in md
    assert "```mermaid" in md and "flowchart LR" in md
    assert "全程预计花费：约 300 元" in md
    assert "早晚温差大" in md
    assert "2人" in md  # 人群写进抬头


def test_render_default_tips_and_pipe_escape():
    import report
    from schemas import TripSpec, Slot, DayPlan
    spec = TripSpec(destination="上海", days=1)
    days = [DayPlan(day=1, theme="", slots=[Slot(time="上午", activity="漫步",
                                                 note="含 a|b 竖线")], est_cost=0)]
    md = report.render_markdown(spec, {}, days)  # 不传 tips
    assert "含 a／b 竖线" in md    # 竖线被替换，避免破坏表格
    assert "提前查好" in md         # 默认贴士兜底


# ============================ Task 5: day_subgraph.py ============================
def test_day_graph_returns_dayplan(monkeypatch):
    import llm
    import day_subgraph
    from schemas import DayPlan, Slot

    def fake_structured(model_cls, system, user, fallback, **k):
        # 子图两个节点都用 DayPlan：直接回一个含 1 个 slot 的计划
        return DayPlan(day=1, theme="古都",
                       slots=[Slot(time="上午", activity="逛", spot="故宫",
                                   transit="地铁", meal="", note="预约")],
                       est_cost=200)

    monkeypatch.setattr(llm, "_structured", fake_structured)
    g = day_subgraph.build_day_graph()
    out = g.invoke({
        "day": 1,
        "outline": {"day": 1, "theme": "古都", "area": "东城", "candidate_spots": ["故宫"]},
        "spec": {"destination": "北京", "pace": "适中", "budget_cny": 2000},
        "research": {}, "prev_days": [], "fix_hint": "",
    })
    plan = out["day_plan"]
    assert isinstance(plan, DayPlan)
    assert plan.day == 1 and plan.slots[0].spot == "故宫"


# ============================ Task 6: agent.py（主图）============================
def _install_fakes(monkeypatch, feas_sequence):
    """把 llm._structured 换成按模型类返回罐头对象的假实现；feas_sequence 逐次控制反思结果。
    返回 calls 计数字典，便于断言反思调用次数。"""
    import llm
    from schemas import TripSpec, DayOutline, TripOutline, DayPlan, Slot
    calls = {"feas": 0}

    def fake_structured(model_cls, system, user, fallback, **k):
        name = model_cls.__name__
        if name == "TripSpec":
            return TripSpec(destination="北京", days=2, interests=["历史"], budget_cny=2000)
        if name == "TripOutline":
            return TripOutline(days=[
                DayOutline(day=1, theme="古都", area="东城", candidate_spots=["故宫", "天安门"]),
                DayOutline(day=2, theme="园林", area="海淀", candidate_spots=["颐和园"])])
        if name == "DayPlan":
            return DayPlan(day=1, theme="古都",
                           slots=[Slot(time="上午", activity="逛", spot="故宫",
                                       transit="地铁", meal="", note="预约")], est_cost=300)
        if name == "Feasibility":
            i = calls["feas"]
            calls["feas"] += 1
            return feas_sequence[min(i, len(feas_sequence) - 1)]
        return fallback

    monkeypatch.setattr(llm, "_structured", fake_structured)

    class _FakeMsg:
        content = "· 测试贴士一\n· 测试贴士二"

    class _FakeLLM:
        def invoke(self, msgs):
            return _FakeMsg()

    monkeypatch.setattr(llm, "get_llm", lambda **k: _FakeLLM())  # generate_tips 免联网
    return calls


def _run_to_interrupt(agent, cfg, req="北京2天"):
    from langchain_core.messages import HumanMessage
    app = agent.build_trip_app()
    events = list(app.stream({"messages": [HumanMessage(req)], "request": req,
                              "use_web": False}, cfg, stream_mode="updates"))
    return app, events


def test_graph_reaches_report_and_interrupt_resume(monkeypatch):
    import agent
    from schemas import Feasibility
    from langgraph.types import Command
    _install_fakes(monkeypatch, [Feasibility(ok=True, issues=[], fixes=[])])
    cfg = {"configurable": {"thread_id": "t1"}}
    app, events = _run_to_interrupt(agent, cfg)
    assert any("__interrupt__" in e for e in events)      # 停在人工审核
    list(app.stream(Command(resume=""), cfg, stream_mode="updates"))  # 空串=批准
    st = app.get_state(cfg).values
    assert st.get("approved") is True
    assert st.get("report_md") and "北京" in st["report_md"]


def test_reflection_loops_then_converges(monkeypatch):
    import agent
    from schemas import Feasibility, DayFix
    from langgraph.types import Command
    calls = _install_fakes(monkeypatch, [
        Feasibility(ok=False, issues=["第1天过满"], fixes=[DayFix(day=1, hint="去掉一个点")]),
        Feasibility(ok=True, issues=[], fixes=[])])
    cfg = {"configurable": {"thread_id": "t2"}}
    app, _ = _run_to_interrupt(agent, cfg)
    list(app.stream(Command(resume=""), cfg, stream_mode="updates"))
    st = app.get_state(cfg).values
    assert calls["feas"] == 2        # 反思跑了两次
    assert st["iteration"] == 2      # 回炉一次后收敛
    assert st.get("report_md")


def test_reflection_caps_at_max_iter(monkeypatch):
    import agent
    from schemas import Feasibility, DayFix
    from langgraph.types import Command
    _install_fakes(monkeypatch, [
        Feasibility(ok=False, issues=["总不满意"], fixes=[DayFix(day=1, hint="改")])])
    cfg = {"configurable": {"thread_id": "t3"}}
    app, _ = _run_to_interrupt(agent, cfg)
    list(app.stream(Command(resume=""), cfg, stream_mode="updates"))
    st = app.get_state(cfg).values
    assert st["iteration"] == agent.MAX_ITER  # 到上限即止，不死循环
    assert st.get("report_md")


def test_feedback_reroutes_to_outline(monkeypatch):
    import agent
    from schemas import Feasibility
    from langgraph.types import Command
    _install_fakes(monkeypatch, [Feasibility(ok=True, issues=[], fixes=[])])
    cfg = {"configurable": {"thread_id": "t4"}}
    app, _ = _run_to_interrupt(agent, cfg)
    ev2 = list(app.stream(Command(resume="第一天想更轻松"), cfg, stream_mode="updates"))
    assert any("__interrupt__" in e for e in ev2)   # 给意见 → 回 outline → 二次审核
    assert app.get_state(cfg).values.get("approved") is False
    list(app.stream(Command(resume=""), cfg, stream_mode="updates"))  # 再批准 → 完成
    assert app.get_state(cfg).values.get("report_md")


def test_export_docx_builds_command(monkeypatch, tmp_path):
    import agent
    captured = {}

    def fake_run(cmd, check=True, **k):
        captured["cmd"] = cmd
        class R: pass
        return R()

    monkeypatch.setattr(agent.subprocess, "run", fake_run)
    monkeypatch.setattr(agent.os.path, "exists", lambda p: True)  # 假装 md2docx.py 在
    md = tmp_path / "行程.md"
    md.write_text("# t", encoding="utf-8")
    ok = agent.export_docx(str(md))
    assert ok is True
    assert any("md2docx.py" in str(c) for c in captured["cmd"])
    assert str(md) in captured["cmd"]


# ============================ 优化回归：并发检索 & 报告清洗/资料附录 ============================
def test_research_node_parallel_and_query_hints(monkeypatch):
    """research_node 用线程池并发四类；键序与 CATEGORIES 一致（确定性）；检索词更聚焦（无叠字）。"""
    import agent
    from schemas import TripSpec
    seen = []

    def fake_web_search(query, use_web=True, category=""):
        seen.append((category, query))
        return f"R-{category}"

    monkeypatch.setattr(agent, "web_search", fake_web_search)
    out = agent.research_node({"spec": TripSpec(destination="北京", days=2), "use_web": True})
    research = out["research"]
    assert list(research.keys()) == agent.CATEGORIES     # ex.map 保序 → 键序确定
    assert research["景点"] == "R-景点" and research["美食"] == "R-美食"
    q_by_cat = dict(seen)
    assert "近期天气" in q_by_cat["天气"] and "天气 天气" not in q_by_cat["天气"]  # 聚焦、无叠字


def test_mermaid_label_sanitized():
    """会破坏 D1S0["..."] 语法的字符被清洗，避免下游 md2docx 渲染 mermaid 出错。"""
    import report
    from schemas import TripSpec, DayPlan, Slot
    spec = TripSpec(destination="杭州", days=1)
    days = [DayPlan(day=1, theme="山水",
                    slots=[Slot(time="上午", activity="逛", spot='西湖"断桥"[东]#景（人多）')],
                    est_cost=0)]
    md = report.render_markdown(spec, {}, days)
    node = [l for l in md.splitlines() if l.strip().startswith("D1S0[")][0]
    label = node.split('["', 1)[1].rsplit('"]', 1)[0]   # 取 ["<label>"] 里的标签本体
    for ch in '"[]#':
        assert ch not in label         # 破坏语法的字符已清除
    assert "（" in label               # 圆括号转全角保留可读性


def test_research_appendix_keeps_real_drops_stub_and_degraded():
    """资料参考附录只收真材料；离线桩/降级说明/空串一律不进报告。"""
    import report
    from schemas import TripSpec, DayPlan, Slot
    spec = TripSpec(destination="北京", days=1)
    days = [DayPlan(day=1, theme="古都",
                    slots=[Slot(time="上午", activity="逛", spot="故宫")], est_cost=10)]
    research = {
        "景点": "- 故宫｜必看（http://x）",          # 真材料 → 入报告
        "天气": "- 近期多云（stub）",                 # 离线桩 → 剔除
        "交通": "(联网检索失败：TimeoutError)",       # 降级说明 → 剔除
        "美食": "",                                   # 空 → 剔除
    }
    md = report.render_markdown(spec, research, days)
    assert "## 资料参考" in md and "故宫｜必看" in md
    assert "stub" not in md and "联网检索失败" not in md


def test_research_appendix_absent_when_no_real_content():
    import report
    from schemas import TripSpec, DayPlan, Slot
    spec = TripSpec(destination="北京", days=1)
    days = [DayPlan(day=1, slots=[Slot(time="上午", activity="逛")], est_cost=0)]
    md = report.render_markdown(spec, {"景点": "- x（stub）", "天气": "(联网无结果)"}, days)
    assert "## 资料参考" not in md


def test_resolve_out_routes_bare_name_to_dist(tmp_path):
    """纯文件名归一到 dist/ 且目录自动建好；绝对路径/带目录的相对路径原样尊重。"""
    import os
    import agent
    out = agent._resolve_out("行程.md")
    assert os.path.dirname(out) == agent.DIST_DIR       # 纯文件名 → 落到 dist/
    assert os.path.isdir(agent.DIST_DIR)                 # dist/ 被自动创建
    abs_p = str(tmp_path / "sub" / "行程.md")
    assert agent._resolve_out(abs_p) == abs_p            # 绝对路径不改写
    assert os.path.isdir(os.path.dirname(abs_p))         # 目标目录也顺带建好


# ============================ 景点介绍功能（200~500字） ============================
def test_spot_intros_section_render_and_omit():
    """景点介绍：有内容 → 渲染「## 景点介绍」+ 各景点三级小节；空 → 整节省略。"""
    import report
    from schemas import TripSpec, DayPlan, Slot
    spec = TripSpec(destination="北京", days=1)
    days = [DayPlan(day=1, theme="古都",
                    slots=[Slot(time="上午", activity="逛", spot="故宫")], est_cost=0)]
    md = report.render_markdown(spec, {}, days, spot_intros={"故宫": "明清两代皇宫，红墙金瓦，是必看的地标。"})
    assert "## 景点介绍" in md and "### 故宫" in md and "明清两代皇宫" in md
    md2 = report.render_markdown(spec, {}, days, spot_intros={})
    assert "## 景点介绍" not in md2       # 空 → 不出现


def test_clip_intro_caps_upper_bound():
    """介绍超 500 字被截到 ≤500 且断在句末标点；不超则原样返回。"""
    import agent
    long = "这是一句测试。" * 200            # 远超 500 字
    out = agent._clip_intro(long)
    assert len(out) <= 500 and out.endswith("。")   # 不超上限、断在句末
    assert agent._clip_intro("短介绍。") == "短介绍。"


def test_spot_intro_node_dedup_and_skip(monkeypatch):
    """去重保序、跳过空 spot；--no-intro 时整节跳过、返回空 dict。"""
    import agent
    from schemas import TripSpec, DayPlan, Slot
    monkeypatch.setattr(agent, "_intro_one", lambda spec, spot, research: f"intro-{spot}")
    spec = TripSpec(destination="北京", days=2)
    days = [
        DayPlan(day=1, slots=[Slot(time="上午", activity="逛", spot="故宫"),
                              Slot(time="午餐", activity="吃饭", spot="")], est_cost=0),
        DayPlan(day=2, slots=[Slot(time="上午", activity="逛", spot="故宫"),      # 重复
                              Slot(time="下午", activity="逛", spot="颐和园")], est_cost=0),
    ]
    out = agent.spot_intro_node({"spec": spec, "days": days, "use_intro": True})
    assert list(out["spot_intros"].keys()) == ["故宫", "颐和园"]   # 去重保序、空 spot 跳过
    assert out["spot_intros"]["故宫"] == "intro-故宫"
    out2 = agent.spot_intro_node({"spec": spec, "days": days, "use_intro": False})
    assert out2["spot_intros"] == {}                              # --no-intro → 跳过


