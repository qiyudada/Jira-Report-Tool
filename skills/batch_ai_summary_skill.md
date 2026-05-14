# Jira Batch AI Summary Skill

## Purpose
用于批量总结 Jira issue 技术进展的 DeepSeek 提示词模板。

## When to Use
- 当 `batch_mode_var.get()` 为 True 时
- 需要一次性总结多个 issue 时调用

## Prompt Template

```
总结以下每个Jira issue技术进展，每项1~2句话。优先级：验证/恢复/关闭结果 > 当前用户或我方方案/文件/补丁 > 分析结论 > 待确认。
必须写成动作+结果/状态，不要输出裸路径、文件名、NV/CFUN关键词或日志名。
格式：issue_key: 总结内容（无实质进展才写：issue_key: 仍在排查中）
```

## Prompt Rules

1. **优先级顺序**：
   - 验证/恢复/关闭结果 > 当前用户或我方方案/文件/补丁 > 分析结论 > 待确认

2. **输出格式**：
   - 每项 1~2 句话
   - 格式：`issue_key: 总结内容`
   - 无实质进展才写：`issue_key: 仍在排查中`

3. **禁止输出**：
   - 裸路径、文件名
   - NV/CFUN 关键词
   - 日志名

4. **标注说明**（当有背景上下文时）：
   - `[本期]` = 报告周期内评论
   - `[背景]` = 周期前背景
   - 优先基于 `[本期]` 总结

## Context Window
- 每个 issue 的评论最多 600 字符
- 整体 prompt max_tokens: 2000