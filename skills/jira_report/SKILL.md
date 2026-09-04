---
name: jira-report
description: 生成Jira周报Excel报告。当用户请求"生成Jira报告"、"Jira周报"、"创建Jira Excel报告"、"导出Jira报告"时调用此skill。
---

# Jira Report Skill (Agent Flow)

此skill用于生成Jira周报Excel文件。**AI 总结由调用本 skill 的 agent 自己完成，不调用 DeepSeek**。

## 调用流程

```
┌─────────────────────────────────────────────────────────────┐
│ Step 1: cli.py prepare                                      │
│   → 登录 Jira、拉取 issue、过滤、收集评论、写入 data.json   │
│   → blocked 状态自动写 prefilled_summary                    │
├─────────────────────────────────────────────────────────────┤
│ Step 2: cli.py derive（代码兜底，不调 LLM）                 │
│   → 读 data.json，对每个 prefilled 为空的 issue 确定性提取  │
│     进展（验证/关闭 > 方案/补丁 > 最新评论）                │
│   → 输出 summaries.json 初稿 + 低置信清单                  │
├─────────────────────────────────────────────────────────────┤
│ Step 3: agent 复核/润色 summaries.json                      │
│   → 低置信 issue 用本地 Claude 生成更自然总结               │
│   → 有实质进展的 issue 可直接保留 derive 结果               │
├─────────────────────────────────────────────────────────────┤
│ Step 4: cli.py export                                       │
│   → 读 data.json + summaries.json，导出 report.xlsx        │
└─────────────────────────────────────────────────────────────┘
```

> 分支：若用户选「使用最新评论原文」（问题 3），Step 2/3 跳过 `derive`，改按「agent 总结规则 §6」取当前用户最新评论原文。

## 调用时先出示选项（交互式）

skill 被触发后、执行三步流程**之前**，必须用 `AskUserQuestion` 工具向用户出示选项让用户点选，不要直接套用默认值，也不要让用户用自然语言补述关键参数。

### 问题 1：报告周期（必问）
| 选项 | 含义 |
|------|------|
| 本周（推荐） | 本周一 ~ 周日 |
| 上周 | 上周一 ~ 周日 |
| 本月 | 本月 1 日 ~ 今日 |
| 自定义 | 用户再输入 `YYYY-MM-DD ~ YYYY-MM-DD` |

### 问题 2：状态筛选（必问）
| 选项 | 说明 |
|------|------|
| ALL（全部） | 不筛选 |
| WAIT FAE INFO | 等待 FAE 信息 |
| WORKED AROUND | 已绕行 |
| WORKING | 处理中 |

其余状态（`CLOSED` / `RESOLVED` / `WAIT 3RD PARTY`）由用户通过 "其他" 手动输入。

### 问题 3：进展来源（是否使用 AI summary，必问）
| 选项 | 说明 |
|------|------|
| 使用 AI 总结（推荐） | agent 用本地 Claude 总结评论，消耗 Claude token |
| 使用最新评论原文 | agent 不总结，只取当前用户最新评论原文作进展 |

- 「使用 AI 总结」= **先跑 `cli.py derive` 生成确定性初稿（代码兜底），再由 Claude 复核/润色**（尤其低置信 issue）。即使 Claude 不动手，summaries.json 也有符合 GUI fallback 质量的进展，不会留空、不会大面积「仍在排查中」。
- 「使用最新评论原文」= 不跑 derive、不调 LLM，按「agent 总结规则 §6」取**当前用户**最新评论原文。语义与 GUI 的「fetch latest comment」一致。

### 问题 4：输出路径（必问）
| 选项 | 说明 |
|------|------|
| 默认下载目录（推荐） | `~/Downloads/{username_short}_{start}_{end}_jira_report.xlsx` |
| 自定义 | 用户再输入完整输出路径（含 `.xlsx` 后缀） |

- 默认文件名**必须**与 UI 导出一致：`{username_short}_{start}_{end}_jira_report.xlsx`
  - `username_short` = Jira 用户名 `@` 前部分（如 `saikia.chen`）
  - `start` / `end` = `YYYY-MM-DD`
  - 示例：`saikia.chen_2026-08-31_2026-09-06_jira_report.xlsx`

## 使用条件

当用户请求生成Jira报告时调用。Claude Code 应：
1. 先按上文「调用时先出示选项」用 `AskUserQuestion` 一次性收集：报告周期、状态筛选、进展来源（是否 AI summary）、输出路径
2. 从环境变量读取 Jira 账号密码（`JIRA_USERNAME` / `JIRA_PASSWORD`）；缺失时提示用户提供
3. 执行三步流程

## 必需参数

| 参数 | 说明 |
|------|------|
| `--username/-u` | Jira 用户名（邮箱格式） |
| `--password/-p` | Jira 密码 |
| `--start` | 开始日期 (YYYY-MM-DD) |
| `--end` | 结束日期 (YYYY-MM-DD) |
| `-o/--output` | 输出 Excel 文件路径（export 步骤） |

## 可选参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--status` | ALL | 按状态筛选 |
| `--columns` | 1,2,3,4,5,6,7 | 列顺序 |
| `--header-align` | left | 表头对齐 |
| `--cell-align` | center | 单元格对齐 |
| `--no-key-highlight` | false | 是否标红关键问题 |

## 调用示例

```bash
# Step 1: 拉数据
python cli.py prepare -u username@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 -o /tmp/data.json

# Step 2: 代码兜底，确定性提取进展（不调 LLM），输出初稿 + 低置信清单
python cli.py derive --input /tmp/data.json -o /tmp/summaries.json

# Step 3: agent 读 /tmp/summaries.json 初稿，按下方规则复核/润色低置信 issue

# Step 4: 导出
python cli.py export --input /tmp/data.json \
                     --summaries /tmp/summaries.json \
                     -o ~/Downloads/username_2026-05-01_2026-05-31_jira_report.xlsx
```

## data.json 结构

```json
{
  "metadata": {
    "start_date": "2026-05-01",
    "end_date": "2026-05-31",
    "generated_at": "2026-06-08T10:00:00",
    "status_filter": "ALL",
    "column_order": "1,2,3,4,5,6,7",
    "key_issue_highlight": true,
    "header_align": "left",
    "cell_align": "center",
    "jira_base_url": "https://ticket.quectel.com"
  },
  "issues": [
    {
      "key": "FAE-123",
      "summary": "问题描述",
      "status": "WAIT OFFICIAL RELEASE",
      "customer_name": "佰才邦",
      "model_name": "RG220UAB",
      "key_issue": "是",
      "highlight_key_issue": true,
      "comments": [
        {
          "author": "张三",
          "date": "2026-05-15",
          "body": "已提供新版本，等待客户验证",
          "in_period": true,
          "author_role": "我方"
        }
      ],
      "prefilled_summary": "[问题已解决，等待SPM同步版本] <2026-05-15>"
    }
  ]
}
```

## agent 总结规则（硬约束）

### 1. prefilled_summary 优先
- 如果 `prefilled_summary` 非空 → **跳过此 issue，不要写 summaries.json 条目**。
- `export` 步骤会用 prefilled_summary 覆盖 agent 总结。

### 2. 复核/润色 derive 初稿（对 prefilled 为空的 issue）

`derive` 已用代码（复用 GUI 的 fallback 逻辑）为每个 issue 确定性提取了进展，保证有验证/关闭或方案信号时**不会**被写成「仍在排查中」。Claude 的角色是**复核 + 精修**，不是从零写：

**必须做的：**
1. 读 `cli.py derive` 打印的**低置信清单**（无验证/关闭、无方案信号的 issue）。
2. 对低置信 issue，结合 comments 判断：
   - 评论确有实质进展但 derive 没识别（如口语化表达）→ 用 LLM 改写成具体「动作+结果」。
   - 确实只有等待/通知类评论 → 才保留「仍在排查中」，或写更贴切的「等待XX回复」。
3. 对高置信 issue（derive 已给出具体进展）→ **通常直接保留**；只有明显错误（选错评论、张冠李戴）才修正。

**润色时遵循的成品 prompt（逐字遵循）：**
> 用1~3句话总结技术进展，优先级：验证/恢复/关闭结果 > 当前用户或我方提供的方案/文件/补丁 > 分析结论 > 待确认。必须写成「动作+结果/状态」，不要输出裸路径、文件名、NV/CFUN关键词或日志名。若客户回复验证可以/恢复正常，要明确写验证通过/问题关闭；若无实质进展才回复【仍在排查中】。

**防偷懒硬约束（关键）：**
- 「仍在排查中」**只能**在评论里确实没有任何实质进展时使用。
- 只要评论里出现下列任一**实质进展信号**，就必须写成具体动作，**禁止**输出「仍在排查中」：
  - 验证/关闭信号：验证可以、验证通过、测试通过、恢复正常、问题关闭、此单关闭、解决、closed、验证完成、没有问题
  - 方案信号：提供、替换、修改、配置、方案、补丁、patch、烧写、排查、确认、说明、建议、NV文件
- 违反此约束是**错误**，不是风格问题。

**提取优先级（拿不准时，按此从评论里确定性提取，而不是空泛概括）：**
1. 验证/恢复/关闭结果（如「客户验证通过，问题关闭」）
2. 当前用户或我方方案/文件/补丁（如「提供XX文件替换方案」）
3. 分析结论（如「初步定位为NV配置问题」）
4. 最新评论原文（截断到 100 字）

### 3. 评论处理（润色时参考）
- `data.json` 里的 `comments` 已由 `prepare` 排序+截断：`in_period: true` 评论全部保留（新在前），`in_period: false` 背景评论按进展价值排序、累计 600 字封顶、单条 250 字封顶。**agent 无需再自己做长度截断。**
- `in_period: true` = [本期]（报告周期内），`in_period: false` = [背景]（周期前，最多 60 天）。
- 正文里的裸路径/附件文件名/日志名已由 `prepare` 压缩（路径→「对应路径」）。
- `author_role` 提示：「当前用户」=自己，「我方」=同事，「客户/Reporter」=客户。

### 4. summaries.json 格式
- 扁平对象：`{ "FAE-123": "进展文本", "FAE-124": "进展文本" }`
- 键是 issue_key（不带"FAE-"等前缀？带，issue_key 就是完整 key）。
- 文本 UTF-8，单行，不含 JSON 控制字符。
- 不要包含 `prefilled_summary` 已有内容的 issue。

### 5. export 前自查
执行 export 前，agent 必须校验：
- data.json 中每个 `prefilled_summary` 为空的 issue，在 summaries.json 中都有非空 entry。
- 如有缺失，**报错并要求补充，不要猜测或跳过**。

### 6. 不使用 AI summary 时的进展填充
当用户在「问题 3」选择「使用最新评论原文」时，agent **不总结**，按以下确定性规则填 summaries.json：
- 对每个 `prefilled_summary` 为空的 issue，从 `comments` 中筛出**当前用户**（`author_role == "当前用户"`）的评论。
- 取其中最新一条 `in_period: true` 的 `body` 作为进展原文。
- 若无 `in_period: true` 评论，取当前用户最新一条背景评论（`in_period: false`）的 `body`。
- 若当前用户无任何评论，写「无评论」。
- 评论正文应已由 `prepare` 清洗过；如仍有 HTML 或附件残留，agent 简单去除即可，不要改写语义。

## 报告内容

生成的 Excel 包含以下列：
- 客户名称
- 型号
- 问题描述
- JIRA号（带链接）
- 状态
- 是否为重点问题
- 进展（prefilled_summary 或 agent 总结）

## 环境变量

- `JIRA_USERNAME` - Jira 用户名
- `JIRA_PASSWORD` - Jira 密码

注意：**本 skill 不再需要 `DEEPSEEK_API_KEY`**，AI 总结由 agent 自己完成。

## 向后兼容

老的 `cli.py ... --ai` 一发式命令（调 DeepSeek）仍然可用，路径为 `cli.py run`，但**不推荐** agent 使用——重复调外部 LLM 没有必要。GUI 工具 `python jira_report_generator.py` 完全不受影响。
