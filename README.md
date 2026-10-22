# travel_agent · 旅行行程规划智能体

[English](README.en.md) | **简体中文**

用一张 LangGraph 有向图，把「自然语言旅行需求 → 结构化逐日行程 → Markdown 报告 →（可选）Word」整条链路跑通，一次演示四大进阶特性：**Human-in-the-loop 审核、逐日子图、可行性反思循环、Markdown→Word 导出**。

```mermaid
flowchart LR
    intake["需求解析"] --> research["联网检索"] --> outline["生成大纲"] --> review["人工审核"]
    review -->|"有修改意见"| outline
    review -->|"批准"| plan["逐日规划"]
    plan --> feas["可行性反思"]
    feas -->|"不可行且未到上限"| plan
    feas -->|"通过"| intro["景点介绍"]
    intro --> write["生成报告"]
```

## 主要特性

| 特性 | 说明 |
|---|---|
| Human-in-the-loop 审核 | `interrupt()` 抛出大纲，CLI 人工批准或输入意见回炉重排 |
| 逐日子图 | 「排动线 → 补后勤」封装为独立 StateGraph，子状态隔离、逐天复用 |
| 可行性反思循环 | 结构化判定过载 / 超预算 / 时间冲突，带按天 fixes 回炉，`MAX_ITER` 封顶防死循环 |
| Markdown → Word | 同目录 `md2docx.py` 一键导出（`--docx`） |
| 联网检索（可选） | TinyFish CLI 并发查景点 / 天气 / 交通 / 美食；缺失或超时自动降级，不断图 |
| 景点介绍 | 行程定稿后为每个景点并发生成 200~500 字介绍（`--no-intro` 关闭） |
| 全链路容错 | 所有 LLM 结构化调用带 fallback；`--no-web` 离线桩可完全断网自测 |

## 快速开始

### 环境要求

- Python 3.10+
- 一个 Anthropic 兼容的 API（官方 API 或任意 Claude 兼容代理）

### 安装

```bash
git clone https://github.com/moduwusuowei/travel_agent
cd travel_agent
pip install -r requirements.txt
```

### 配置

```bash
cp .env.example .env        # Windows: copy .env.example .env
```

编辑 `.env` 填入你的配置（`.env` 已被 `.gitignore` 忽略，不会提交）：

| 变量 | 必填 | 说明 |
|---|---|---|
| `ANTHROPIC_API_KEY` 或 `ANTHROPIC_AUTH_TOKEN` | 是 | API Key / 代理 token，二者填其一 |
| `ANTHROPIC_BASE_URL` | 是 | 官方 `https://api.anthropic.com` 或你的代理地址 |
| `VERIFY_MODEL` | 否 | 覆盖默认模型（默认 `claude-opus-4-8`） |

### 运行

```powershell
# 1) 最稳的第一步：离线跑（内置桩数据 + 自动批准审核）
python agent.py -q "北京3天，喜欢历史和美食，预算3000，两个人" --no-web --auto-approve

# 2) 联网 + 人工审核（默认）：停在大纲，回车批准 / 输入修改意见回炉
python agent.py -q "成都3天 美食为主"

# 3) 交互模式（同一进程多轮，各自独立 thread）
python agent.py

# 4) 导出 Word（调用同目录 md2docx.py）
python agent.py -q "上海2天 亲子" --docx --auto-approve

# 5) 不生成景点介绍，更快
python agent.py -q "上海2天 亲子" --no-intro --auto-approve
```

> 报告与 Word 默认写入 `dist/`；`-o` 传绝对路径或带目录的相对路径时按你给的路径输出。

### CLI 参数

| 参数 | 作用 |
|---|---|
| `-q/--question` | 旅行需求；不给则进交互模式 |
| `--no-web` | 关闭联网，用内置桩数据离线跑 |
| `--no-intro` | 不生成景点介绍（更快） |
| `--docx` | 跑完把报告导出 Word |
| `--auto-approve` | 跳过人工审核（脚本化 / 测试用） |
| `-o/--out` | 报告输出路径（默认 `行程_目的地_N天.md`，落到 `dist/`） |

## 项目结构

```
travel_agent/
├── agent.py           # 主图：8 节点 + HITL + 反思 + 流式 + CLI
├── day_subgraph.py    # 逐日规划子图（排动线 → 补后勤）
├── schemas.py         # 7 个 Pydantic 数据契约
├── llm.py             # get_llm() / text_of() / _structured()（结构化兜底）
├── tools.py           # TinyFish 检索封装 + 离线桩 + 降级
├── report.py          # DayPlan[] → Markdown（总览表 / mermaid / 资料附录）
├── md2docx.py         # Markdown → Word 转换器（--docx 调用）
├── test_agent.py      # 25 项离线单测（不联网、不调真模型）
├── .env.example       # 环境变量模板
├── requirements.txt
└── README.md
```

## 测试

```bash
pip install pytest
python -m pytest test_agent.py -q
```

全部离线运行：monkeypatch 隔离网络与真实模型，覆盖 HITL 恢复、反思收敛 / 封顶、报告渲染、导出等路径。

## 每个文件学到什么

- **llm.py**：在有坑的真实环境里稳妥拿到 LLM——关 thinking、给足 max_tokens、`text_of()` 兜底列表 content、`_structured()` 结构化失败返回 fallback。
- **schemas.py**：用 Pydantic 给 agent 每一步定数据契约；容器模型绕开 `with_structured_output` 不能绑 `list[X]` 的限制。
- **tools.py**：把外部 CLI 安全包成工具——Windows `cmd /c`、`raw_decode` 取首个 JSON、失败降级、离线桩。
- **report.py**：把结构化结果确定性渲染成带表格与 mermaid 的 Markdown（纯函数、可单测）。
- **day_subgraph.py**：把多步流程做成子图，被主图当可复用组件逐天调用，子状态隔离。
- **agent.py**：LangGraph 集大成——`StateGraph` + `add_messages`、`MemorySaver` 跨轮记忆、`interrupt`/`Command` HITL、条件边反思循环、`stream_mode="updates"` 流式展示。

## License

[Apache License 2.0](LICENSE)
