---
name: jira-report
description: 生成Jira周报Excel报告。当用户请求"生成Jira报告"、"Jira周报"、"创建Jira Excel报告"、"导出Jira报告"时调用此skill。
---

# Jira Report Skill

此skill用于生成Jira周报Excel文件，支持可选的AI智能总结功能。

## 使用条件

当用户请求生成Jira报告时调用。Claude Code应：
1. 解析日期范围（如果用户未指定，提示用户输入）
2. 确定输出文件路径
3. 从环境变量或用户处获取Jira账号密码
4. 调用CLI生成报告

## 必需参数

| 参数 | 说明 |
|------|------|
| `--username/-u` | Jira用户名（邮箱格式） |
| `--password/-p` | Jira密码 |
| `--start` | 开始日期 (YYYY-MM-DD) |
| `--end` | 结束日期 (YYYY-MM-DD) |
| `-o/--output` | 输出Excel文件路径 |

## 可选参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--status` | ALL | 按状态筛选 |
| `--columns` | 1,2,3,4,5,6,7 | 列顺序 |
| `--header-align` | left | 表头对齐 |
| `--cell-align` | center | 单元格对齐 |
| `--no-key-highlight` | false | 是否标红关键问题 |
| `--ai` | false | 启用AI总结 |
| `--ai-key` | env:DEEPSEEK_API_KEY | DeepSeek API Key |
| `--ai-model` | deepseek-chat | DeepSeek模型 |
| `--batch-mode` | false | 批量AI总结模式 |
| `--batch-size` | 10 | AI批量大小 |
| `--fetch-comment` | false | 获取最新评论作为进展 |
| `--timestamp-prefix` | false | 评论添加时间前缀 |

## 调用示例

```bash
# 基本用法
python cli.py -u username@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx

# 带AI总结
python cli.py -u username@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx --ai

# 批量AI模式
python cli.py -u username@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx --ai --batch-mode --batch-size 20
```

##报告内容

生成的Excel包含以下列：
- 客户名称
- 型号
- 问题描述
- JIRA号（带链接）
- 状态
- 是否为重点问题
- 进展（AI总结或最新评论）

## 环境变量

- `JIRA_USERNAME` - Jira用户名
- `JIRA_PASSWORD` - Jira密码
- `DEEPSEEK_API_KEY` - DeepSeek API Key（AI功能必需）