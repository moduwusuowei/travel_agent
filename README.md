# travel_agent —— 旅行行程规划智能体

用一张 LangGraph 有向图，把「自然语言旅行需求 → 结构化逐日行程 → Markdown 报告 →（可选）Word」跑通，
一次演示四大进阶特性：**Human-in-the-loop 审核、逐日子图、可行性反思循环、Markdown→Word 导出**。

```
intake → research → outline → [human_review(interrupt)] → plan_days → feasibility → spot_intro → write_report
                        ↑__有修改意见回炉__|                     |__不可行且未到上限回炉(≤2)__|
```

> `spot_intro`：行程定稿后，为其中每个景点**并发**生成 200~500 字介绍，写入报告「景点介绍」节（`--no-intro` 关闭）。

---

## 环境要点（本机已配好）

| 事项 | 说明 |
|---|---|
| LLM | Claude 兼容 API/代理，配置读本目录 `.env`（模板见 `.env.example`） |
| 模型 | 默认 `claude-opus-4-8`（关 thinking，`.content` 是纯字符串）；`text_of()` 兜底列表 content |
| 实时信息 | TinyFish CLI 搜索（景点/天气/交通/美食）；无网/无 tinyfish 自动降级；`--no-web` 用内置桩 |
| 结构化 | Pydantic + `with_structured_output`，每次调用带 fallback 对象，抽风也不断图 |
| 导出 | 同目录 `md2docx.py`（`--docx`，subprocess 调用；缺 pandoc/python-docx 自动降级/跳过） |
| 依赖 | 见 `requirements.txt`，本机基本已装 |

`ANTHROPIC_AUTH_TOKEN` 是真实密钥：不打印、不写进报告、不提交。

---

## 目录结构

```
travel_agent/
├── DESIGN.md          # 设计文档
├── PLAN.md            # 实现计划（TDD 任务分解）
├── llm.py             # get_llm() / text_of() / _structured()
├── schemas.py         # 7 个 Pydantic 模型
├── tools.py           # TinyFish 检索封装 + 离线桩 + 降级
├── day_subgraph.py    # 逐日规划子图
├── agent.py           # 主图 + 节点 + HITL + 反思 + 流式 + 导出 + CLI
├── report.py          # DayPlan[] → Markdown（含 mermaid）
├── test_agent.py      # 离线单测（不联网、不调真模型）
├── requirements.txt
├── dist/              # 生成的报告 md / docx 统一落到这里（运行后出现）
└── README.md
```

---

## 快速开始（PowerShell，先 cd 进本目录）

```powershell
# 0) 不联网跑（用内置桩数据替代实时检索 + 自动批准）；注意仍需 LLM 代理，最稳的第一步
python agent.py -q "北京3天，喜欢历史和美食，预算3000，两个人" --no-web --auto-approve

# 1) 联网 + 人工审核（默认）：会停在大纲让你回车批准 / 输入修改意见回炉
python agent.py -q "成都3天 美食为主"

# 2) 交互模式（同一进程多轮，各自独立 thread）
python agent.py

# 3) 跑完导出 Word（复用同目录 md2docx.py）
python agent.py -q "上海2天 亲子" --docx --auto-approve

# 3b) 想更快、不要景点介绍
python agent.py -q "上海2天 亲子" --no-intro --auto-approve

# 4) 跑离线单测
python -m pytest test_agent.py -q
```

> 报告和 Word 默认都写到 `dist/`（纯文件名会自动归到这里，目录不存在会自动建）；
> `-o` 传绝对路径或带目录的相对路径时，则按你给的路径来，不再进 `dist/`。

---

## 每个文件学到什么

- **llm.py**：在有坑的真实环境里稳妥拿到 LLM——关 thinking、给足 max_tokens、`text_of()` 兜底、`_structured()` 结构化兜底。
- **schemas.py**：用 Pydantic 给 agent 的每一步定「数据契约」；容器模型绕开 `with_structured_output` 不能绑 `list[X]` 的限制。
- **tools.py**：把外部 CLI（TinyFish）安全地包成工具——Windows `cmd /c`、`raw_decode` 取 JSON、失败降级、离线桩。
- **report.py**：把结构化结果确定性地渲染成带表格与 mermaid 的 Markdown（纯函数、可单测）；含「资料参考」「景点介绍」节。
- **day_subgraph.py**：把「排动线 → 补后勤」做成一张**子图**，被主图当可复用组件逐天调用，子状态隔离。
- **agent.py**：LangGraph 的集大成——
  - `StateGraph` + 带 `add_messages` 的 State；`MemorySaver` + `thread_id` 跨轮记忆
  - **HITL**：`interrupt()` 抛大纲、`Command(resume=...)` 回填；条件边按有无意见路由
  - **反思循环**：`feasibility` 结构化判定 + 条件边回 `plan_days`，`MAX_ITER` 封顶
  - 流式 `stream_mode="updates"` 逐节点展示；`--docx` 图外导出 Word

---

## 四大进阶特性 → 代码落点

| 特性 | 位置 | 机制 |
|---|---|---|
| Human-in-the-loop 审核 | `agent.human_review_node` | `interrupt(大纲)` → CLI 展示 → `Command(resume=文本)`；条件边按 feedback 回 `outline` 或继续 |
| 子图逐日规划 | `day_subgraph.build_day_graph` + `agent.plan_days_node` | 编译成独立图，逐天 `.invoke()`，子状态隔离 |
| 可行性反思循环 | `agent.feasibility_node` + `route_after_feasibility` | 结构化 `ok`；不 ok 且 `iteration<MAX_ITER` 回 `plan_days` 带 `fixes` |
| Markdown→Word 导出 | `report.render_markdown` + `agent.export_docx`（CLI `--docx`） | 带 mermaid 的 md → `md2docx.py` |
| 景点介绍（200~500字） | `agent.spot_intro_node` + `report._spot_intros_section` | 定稿后按景点去重、线程池并发生成；`_clip_intro` 保证 ≤500 字；`--no-intro` 关 |
