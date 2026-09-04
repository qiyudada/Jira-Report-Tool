"""
Deterministic progress extraction for Jira issues.

These pure functions mirror the GUI's fallback logic (previously duplicated in
`AISummarizer`) so the skill/agent flow can produce a stable, non-empty 进展 for
every issue without calling any LLM. The rules prefer a real progress signal
(verification/resolution > solution/file/patch > analysis > latest comment)
instead of collapsing to「仍在排查中」.

Single source of truth: `AISummarizer` delegates its private helpers to these
functions, and `cli.py derive` uses them directly. Do not copy these rules again.
"""
import re


def compact_comment_signal(body):
    """Reduce noisy technical text before extracting progress.

    Keep action/result words, but collapse long paths and attachment-like
    fragments so a bare path, file name, CFUN/NV token, or log name is never
    mistaken for progress.
    """
    if not body:
        return ""

    text = body
    text = re.sub(r'(?:[A-Za-z0-9_.-]+[\\/]){2,}[A-Za-z0-9_.-]+', '[路径]', text)
    text = re.sub(r'(?<![A-Za-z0-9_.-])(?:[A-Za-z]:)?(?:[\\/][A-Za-z0-9_. -]+){2,}', '[路径]', text)
    text = re.sub(r'\b[\w.-]+\.(?:zip|rar|7z|tar|gz|log|dmp|txt)\b\s*(?:\([^)]*\))?', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+', ' ', text).strip()

    # Common short-form solution comments: "qlrild 替换[路径]下同文件试下"
    text = re.sub(r'\b([\w.-]{2,})\s+替换\[路径\]下同文件试下', r'提供\1文件替换方案', text, flags=re.IGNORECASE)
    text = re.sub(r'替换\[路径\]下同文件试下', '替换对应目录下同名文件', text)
    text = text.replace('[路径]下同文件', '对应目录下同名文件')
    text = text.replace('[路径]', '对应路径')
    return text.strip()


def normalize_progress_text(text):
    """Make extracted text read like progress, not a copied raw comment."""
    text = compact_comment_signal(text)
    text = re.sub(r'^提供([\w.-]{2,})文件替换方案$', r'提供\1文件替换方案', text)
    text = text.replace('验证可以', '验证通过')
    text = text.replace('验证完成，没有问题', '验证无问题')
    text = text.replace('此单关闭', '问题关闭')
    text = re.sub(r'\s+', ' ', text).strip(' ，,。')
    return text


def has_resolution_signal(text):
    """True if text signals a verification/resolution/closure outcome."""
    return bool(re.search(
        r'验证可以|验证通过|测试通过|恢复正常|问题关闭|此单关闭|解决|closed|验证完成|没有问题',
        text, re.IGNORECASE))


def has_solution_signal(text):
    """True if text signals a concrete solution/file/patch action."""
    return bool(re.search(
        r'提供|替换|修改|配置|方案|补丁|patch|disable|disabled|烧写|排查|确认|说明|建议|NV文件',
        text, re.IGNORECASE))


def comment_signal_score(comment):
    """Rank comments by progress value; used when truncating long threads."""
    body = comment.get('body', '')
    role = comment.get('author_role', '')
    score = 0
    if comment.get('in_period', True):
        score += 2
    if role in ("当前用户", "我方"):
        score += 2
    if has_resolution_signal(body):
        score += 6
    if has_solution_signal(body):
        score += 4
    if re.search(r'\b(?:Log|dbg|dump|trace)\b|日志|附件', body, re.IGNORECASE):
        score -= 2
    return score


def _date_key(comment):
    """Sort key tolerant of both ISO date strings and datetime objects."""
    date = comment.get('date')
    return str(date)


def fallback_summary(comments, reason=""):
    """Deterministic progress extraction that prefers solution + resolution.

    Mirrors the GUI's AI-failure fallback so the skill flow never emits an empty
    or placeholder 进展 when a real signal exists.
    """
    if not comments:
        return "无评论"

    sorted_comments = sorted(comments, key=_date_key)
    outcome = None
    solution = None
    latest_signal = None

    for comment in sorted_comments:
        body = normalize_progress_text(comment.get('body', ''))
        if not body:
            continue
        if has_solution_signal(body):
            solution = body
            latest_signal = body
        if has_resolution_signal(body):
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

    latest = sorted(comments, key=_date_key, reverse=True)[0]
    body = normalize_progress_text(latest.get('body', ''))
    if len(body) > 100:
        body = body[:100] + "..."
    return body


def _cap_body(body, max_chars):
    """Truncate a single comment body to `max_chars`, marking the cut."""
    if not body:
        return body
    if len(body) > max_chars:
        return body[:max_chars] + "…"
    return body


def limit_comments_for_context(comments, max_per_comment=250, max_total=600):
    """Sort and truncate comments to bound the context the agent must read.

    `in_period` comments are the report's core, so all of them are kept (bodies
    capped at `max_per_comment`). Background comments are ranked by progress
    value (`comment_signal_score`) and only fill the remaining `max_total`
    budget, so low-value history is dropped first.

    Used by `prepare` to keep `data.json` small before the agent reads it.
    """
    if not comments:
        return []

    in_period = sorted(
        [c for c in comments if c.get('in_period', True)],
        key=lambda c: str(c.get('date', '')),
        reverse=True,
    )
    background = sorted(
        [c for c in comments if not c.get('in_period', True)],
        key=lambda c: (comment_signal_score(c), str(c.get('date', ''))),
        reverse=True,
    )

    result = []
    total = 0
    for c in in_period:
        out = dict(c)
        out['body'] = _cap_body(c.get('body', ''), max_per_comment)
        result.append(out)
        total += len(out['body'])
    for c in background:
        body = _cap_body(c.get('body', ''), max_per_comment)
        if total + len(body) > max_total:
            break  # budget exhausted; drop remaining (lowest-value) background
        out = dict(c)
        out['body'] = body
        result.append(out)
        total += len(body)
    return result


def derive_summaries(data):
    """Derive a candidate 进展 for every issue with an empty prefilled_summary.

    Returns (summaries, low_conf_keys):
      - summaries: {issue_key: 进展} — deterministic, non-empty for every
        non-prefilled issue (fallback to latest comment, else「无评论」).
      - low_conf_keys: issue keys whose comments carry no resolution/solution
        signal, i.e. where an LLM could still add value or where「仍在排查中」
        may genuinely apply. The agent should review these specifically.
    """
    summaries = {}
    low_conf_keys = []

    for issue in data.get("issues", []):
        key = issue.get("key", "")
        prefilled = (issue.get("prefilled_summary") or "").strip()
        if prefilled:
            continue  # export uses prefilled and ignores summaries for this key

        comments = issue.get("comments", [])
        summaries[key] = fallback_summary(comments)

        has_signal = any(
            has_resolution_signal(c.get('body', '')) or has_solution_signal(c.get('body', ''))
            for c in comments
        )
        if not has_signal:
            low_conf_keys.append(key)

    return summaries, low_conf_keys
