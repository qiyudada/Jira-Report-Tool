"""
AI Summarizer - DeepSeek AI integration for issue summarization
"""
import requests
import re


class AISummarizer:
    def __init__(self, api_key: str = None):
        self.api_key = api_key or ""

    def summarize(self, issue_key: str, summary: str, comments: list, model: str = "deepseek-chat") -> str:
        """Use DeepSeek AI to summarize issue progress from comments"""
        if not self.api_key:
            return "[AI总结] 未配置API Key"

        if not comments:
            return "[AI总结] 无评论"

        comments_text = self._format_comments_for_ai(comments, max_chars=900)

        has_in_period = any(c.get('in_period', True) for c in comments)
        period_hint = (
            "[本期]=本报告周期内评论，[背景]=周期前背景。优先基于[本期]总结，无[本期]时用[背景]。\n"
            if any(not c.get('in_period', True) for c in comments) else ""
        )

        prompt = (
            f"Issue: {summary}\n"
            f"{'标注说明: ' + period_hint if period_hint else ''}\n\n"
            f"{comments_text}\n\n"
            f"用1~3句话总结技术进展，优先级：验证/恢复/关闭结果 > 当前用户或我方提供的方案/文件/补丁 > 分析结论 > 待确认。\n"
            f"必须写成「动作+结果/状态」，不要输出裸路径、文件名、NV/CFUN关键词或日志名。\n"
            f"若客户回复验证可以/恢复正常，要明确写验证通过/问题关闭；若无实质进展才回复【仍在排查中】。"
        ).strip()

        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": model,
                "messages": [
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": 500,
                "temperature": 0.3
            }
            response = requests.post(
                "https://api.deepseek.com/chat/completions",
                headers=headers,
                json=payload,
                timeout=60
            )

            try:
                result = response.json()
            except ValueError:
                return self._fallback_summary(comments, "API返回格式错误")

            if "error" in result:
                return self._fallback_summary(comments, "API返回错误")

            if response.status_code == 200:
                choices = result.get("choices")
                if not choices:
                    return self._fallback_summary(comments, "API返回空")

                message = choices[0].get("message", {})
                summary_text = message.get("content", "").strip()

                if summary_text:
                    return self._sanitize_ai_summary(summary_text, comments)
                else:
                    return self._fallback_summary(comments)
            elif response.status_code == 401:
                return "[AI总结] DeepSeek API Key无效"
            elif response.status_code == 429:
                return self._fallback_summary(comments, "API请求超过限额")
            elif response.status_code == 400:
                return self._fallback_summary(comments, "请求参数错误")
            else:
                return self._fallback_summary(comments, f"API错误: HTTP {response.status_code}")

        except requests.exceptions.Timeout:
            return self._fallback_summary(comments)
        except requests.exceptions.ConnectionError:
            return self._fallback_summary(comments)
        except Exception:
            return self._fallback_summary(comments)

    def batch_summarize(self, items: list, model: str = "deepseek-chat") -> dict:
        """Batch summarize multiple issues with a single API call

        items: list of dicts with {issue_key, summary, comments}
        Returns: dict mapping issue_key to summary
        """
        if not self.api_key:
            return {item['issue_key']: "未配置API Key" for item in items}

        items_with_comments = [item for item in items if item['comments']]
        if not items_with_comments:
            return {item['issue_key']: "无评论" for item in items}

        def batch_fallback():
            return {
                item['issue_key']: self._fallback_summary(item['comments']) if item['comments'] else "无评论"
                for item in items
            }

        # Build combined prompt
        combined_text = []
        has_context_tags = False
        for item in items_with_comments:
            comments_text = self._format_comments_for_ai(item['comments'], max_chars=600)
            title_line = f"【{item['summary']}】" if item.get('summary') else ""
            combined_text.append(f"## {item['issue_key']} {title_line}\n{comments_text}")
            if any(not c.get('in_period', True) for c in item['comments']):
                has_context_tags = True

        period_note = "[本期]=报告周期内, [背景]=周期前背景。优先基于[本期]总结。\n" if has_context_tags else ""
        issues_block = "\n\n".join(combined_text)

        prompt = (
            f"{period_note}"
            f"总结以下每个Jira issue技术进展，每项1~2句话。优先级：验证/恢复/关闭结果 > 当前用户或我方方案/文件/补丁 > 分析结论 > 待确认。\n"
            f"必须写成动作+结果/状态，不要输出裸路径、文件名、NV/CFUN关键词或日志名。\n"
            f"格式：issue_key: 总结内容（无实质进展才写：issue_key: 仍在排查中）\n\n"
            f"{issues_block}"
        )

        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": model,
                "messages": [
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": 2000,
                "temperature": 0.3
            }
            response = requests.post(
                "https://api.deepseek.com/chat/completions",
                headers=headers,
                json=payload,
                timeout=180
            )

            try:
                result = response.json()
            except ValueError:
                return batch_fallback()

            if "error" in result:
                return batch_fallback()

            if response.status_code == 200:
                choices = result.get("choices", [])
                if choices:
                    content = choices[0].get("message", {}).get("content", "").strip()
                    return self._parse_batch_results(content, items_with_comments)

        except Exception:
            return batch_fallback()

        return batch_fallback()

    def _parse_batch_results(self, content: str, items: list) -> dict:
        """Parse batch AI response into individual summaries"""
        results = {}
        lines = content.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue
            line = re.sub(r'^[-*•\d.\s]+', '', line).strip()
            match = re.match(r'([A-Z]+-\d+)\s*[:：]\s*(.+)$', line)
            if match:
                key = match.group(1).strip()
                summary = match.group(2).strip()
                item = next((item for item in items if item['issue_key'] == key), None)
                results[key] = self._sanitize_ai_summary(summary, item['comments']) if item else summary

        for item in items:
            key = item['issue_key']
            if key not in results:
                comments = item['comments']
                if comments:
                    results[key] = self._fallback_summary(comments)
                else:
                    results[key] = "无评论"

        return results

    def _format_comments_for_ai(self, comments: list, max_chars: int = 1500) -> str:
        """Format pre-cleaned comments into a compact string for AI prompt"""
        if not comments:
            return ""

        formatted = []
        total_chars = 0

        ordered_comments = sorted(
            comments,
            key=lambda c: (self._comment_signal_score(c), c.get('date')),
            reverse=True
        )

        for c in ordered_comments:
            body = self._compact_comment_signal(c['body'])
            if not body:
                continue

            date_str = c['date'].strftime("%m-%d") if hasattr(c['date'], 'strftime') else str(c['date'])
            tag = "[本期]" if c.get('in_period', True) else "[背景]"
            role = c.get('author_role', '评论')

            if len(body) > 250:
                body = body[:250] + "…"

            entry = f"{tag}[{role}/{c['author']} {date_str}] {body}"
            if total_chars + len(entry) > max_chars:
                remaining = max_chars - total_chars
                if remaining > 80:
                    formatted.append(entry[:remaining] + "…")
                break

            formatted.append(entry)
            total_chars += len(entry) + 1

        return "\n".join(formatted)

    def _compact_comment_signal(self, body: str) -> str:
        """Reduce noisy technical text before sending to model"""
        if not body:
            return ""

        text = body
        text = re.sub(r'(?:[A-Za-z0-9_.-]+[\\/]){2,}[A-Za-z0-9_.-]+', '[路径]', text)
        text = re.sub(r'(?<![A-Za-z0-9.-])(?:[A-Za-z]:)?(?:[\\/][A-Za-z0-9_. -]+){2,}', '[路径]', text)
        text = re.sub(r'\b[\w.-]+\.(?:zip|rar|7z|tar|gz|log|dmp|txt)\b\s*(?:\([^)]*\))?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s+', ' ', text).strip()

        text = re.sub(r'\b([\w.-]{2,})\s+替换\[路径\]下同文件试下', r'提供\1文件替换方案', text, flags=re.IGNORECASE)
        text = re.sub(r'替换\[路径\]下同文件试下', '替换对应目录下同名文件', text)
        text = text.replace('[路径]下同文件', '对应目录下同名文件')
        text = text.replace('[路径]', '对应路径')
        return text.strip()

    def _comment_signal_score(self, comment: dict) -> int:
        """Rank comments by progress value"""
        body = comment.get('body', '')
        role = comment.get('author_role', '')
        score = 0
        if comment.get('in_period', True):
            score += 2
        if role in ("当前用户", "我方"):
            score += 2
        if re.search(r'验证可以|验证通过|测试通过|恢复正常|解决|关闭|closed|验证完成|没有问题', body, re.IGNORECASE):
            score += 6
        if re.search(r'提供|替换|修改|配置|方案|补丁|patch|disable|disabled|烧写|排查|确认|说明|建议', body, re.IGNORECASE):
            score += 4
        if re.search(r'\b(?:Log|dbg|dump|trace)\b|日志|附件', body, re.IGNORECASE):
            score -= 2
        return score

    def _normalize_progress_text(self, text: str) -> str:
        """Make fallback summaries read like progress"""
        text = self._compact_comment_signal(text)
        text = re.sub(r'^提供([\w.-]{2,})文件替换方案$', r'提供\1文件替换方案', text)
        text = text.replace('验证可以', '验证通过')
        text = text.replace('验证完成，没有问题', '验证无问题')
        text = text.replace('此单关闭', '问题关闭')
        text = re.sub(r'\s+', ' ', text).strip(' ，,。')
        return text

    def _has_resolution_signal(self, text: str) -> bool:
        return bool(re.search(r'验证可以|验证通过|测试通过|恢复正常|问题关闭|此单关闭|解决|closed|验证完成|没有问题', text, re.IGNORECASE))

    def _has_solution_signal(self, text: str) -> bool:
        return bool(re.search(r'提供|替换|修改|配置|方案|补丁|patch|disable|disabled|烧写|排查|确认|说明|建议|NV文件', text, re.IGNORECASE))

    def _is_low_quality_summary(self, summary: str) -> bool:
        """Detect model outputs that are just a token/path/keyword"""
        if not summary:
            return True
        text = summary.strip()
        if "\\" in text or "/" in text:
            return True
        if len(text) <= 18 and re.search(r'[A-Za-z0-9]', text):
            return True
        if re.fullmatch(r'[\w.-]+', text):
            return True
        return False

    def _sanitize_ai_summary(self, summary: str, comments: list) -> str:
        """Replace weak AI output with deterministic comment-derived fallback"""
        summary = (summary or "").strip()
        fallback = self._fallback_summary(comments)
        if self._is_low_quality_summary(summary):
            return fallback
        if re.search(r'无进展|仍在排查|无实质进展', summary) and fallback not in ("无评论", "仍在排查中"):
            if any(self._has_solution_signal(c.get('body', '')) or self._has_resolution_signal(c.get('body', '')) for c in comments):
                return fallback
        return summary

    def _fallback_summary(self, comments: list, reason: str = "") -> str:
        """Deterministic fallback summary"""
        if not comments:
            return "无评论"

        sorted_comments = sorted(comments, key=lambda x: x['date'])
        outcome = None
        solution = None
        latest_signal = None

        for comment in sorted_comments:
            body = self._normalize_progress_text(comment.get('body', ''))
            if not body:
                continue
            if self._has_solution_signal(body):
                solution = body
                latest_signal = body
            if self._has_resolution_signal(body):
                outcome = body
                latest_signal = body

        if outcome:
            if len(outcome) <= 8 and solution:
                return f"{solution}，{outcome}。"
            return outcome[:120] + ("..." if len(outcome) > 120 else "")

        if solution:
            return solution[:120] + ("..." if len(solution) > 120 else "")

        if latest_signal:
            return latest_signal[:120] + ("..." if len(latest_signal) > 120 else "")

        latest = sorted(comments, key=lambda x: x['date'], reverse=True)[0]
        body = self._normalize_progress_text(latest.get('body', ''))
        if len(body) > 100:
            body = body[:100] + "..."

        return body