---
name: jira-report
description: 生成Jira周报Excel报告。当用户请求"生成Jira报告"、"Jira周报"、"创建Jira Excel报告"、"导出Jira报告"时调用此skill。
---

# Jira Report Skill (Agent Flow)

此skill用于生成Jira周报Excel文件。**AI 总结由调用本 skill 的 agent 自己完成，不调用 DeepSeek**。

## 调用流程（三步）

```
┌─────────────────────────────────────────────────────────────┐
│ Step 1: cli.py prepare                                      │
│   → 登录 Jira、拉取 issue、过滤、收集评论、写入 data.json   │
│   → blocked 状态自动写 prefilled_summary                    │
├─────────────────────────────────────────────────────────────┤
│ Step 2: agent 自己读 data.json                              │
│   → 对 prefilled_summary 为空的 issue 写 1~3 句进展        │
│   → 写入 summaries.json {issue_key: "进展"}                │
├─────────────────────────────────────────────────────────────┤
│ Step 3: cli.py export                                       │
│   → 读 data.json + summaries.json，导出 report.xlsx        │
└─────────────────────────────────────────────────────────────┘
```

## 使用条件

当用户请求生成Jira报告时调用。Claude Code 应：
1. 解析日期范围（如果用户未指定，提示用户输入，默认"本周"或"上周"）
2. 确定输出文件路径（默认 `~/Downloads/jira_report_<start>_<end>.xlsx`）
3. 从环境变量或用户处获取 Jira 账号密码（`JIRA_USERNAME` / `JIRA_PASSWORD`）
4. 执行三步流程

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

# Step 2: agent 读 /tmp/data.json，按下方规则生成 summaries.json

# Step 3: 导出
python cli.py export --input /tmp/data.json \
                     --summaries /tmp/summaries.json \
                     -o ~/Downloads/jira_report_2026-05-01_2026-05-31.xlsx
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

### 2. 进展总结（对 prefilled 为空的 issue）
- 1~3 句话，中文。
- 优先级（从高到低）：
  1. **验证/恢复/关闭结果**（如"客户验证通过，问题关闭"）
  2. **当前用户或我方方案/文件/补丁**（如"提供XXX文件替换方案"）
  3. **分析结论**（如"初步定位为NV配置问题"）
  4. **待确认**（如"等待FAE回复"）
- 写成「动作+结果/状态」，不是「问题描述」。
- 禁止：裸路径、文件名、NV/CFUN 关键词、日志名、英文标点堆砌。
- 若客户回复验证可以/恢复正常，要明确写「验证通过/问题关闭」。
- 若无实质进展（含 in_period 评论但都是等待/通知类），写「仍在排查中」。

### 3. 评论处理
- 优先基于 `in_period: true` 的评论总结。
- `in_period: false` 的评论只作为背景上下文（最多 60 天前）。
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
