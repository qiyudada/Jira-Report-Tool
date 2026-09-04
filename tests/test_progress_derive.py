"""Tests for the deterministic progress extraction shared by GUI + CLI derive."""
from src.progress_derive import (
    derive_summaries,
    fallback_summary,
    has_resolution_signal,
    has_solution_signal,
    limit_comments_for_context,
)


def _comment(body, date="2026-05-20", role="我方", in_period=True):
    return {"date": date, "body": body, "author_role": role, "in_period": in_period}


def test_fallback_prefers_solution_signal():
    # A concrete action must surface, never collapse to a placeholder.
    summary = fallback_summary([_comment("修改了蓝牙重连参数，提供补丁让客户验证")])
    assert "提供补丁" in summary


def test_fallback_combines_solution_and_resolution():
    comments = [
        _comment("提供新固件替换方案", date="2026-05-18"),
        _comment("验证通过", date="2026-05-19", role="客户/Reporter"),
    ]
    summary = fallback_summary(comments)
    assert "验证通过" in summary
    assert "提供" in summary


def test_fallback_empty_comments():
    assert fallback_summary([]) == "无评论"


def test_fallback_wait_only_returns_latest_not_placeholder():
    # Even a wait-only comment yields its text (not「仍在排查中」) as the floor.
    assert fallback_summary([_comment("等待FAE回复")]) == "等待FAE回复"


def test_signals():
    assert has_solution_signal("提供补丁") is True
    assert has_solution_signal("等待回复") is False
    assert has_resolution_signal("客户验证通过") is True
    assert has_resolution_signal("正在排查") is False


def test_derive_skips_prefilled():
    data = {
        "issues": [
            {"key": "FAE-1", "prefilled_summary": "已预填", "comments": []},
            {"key": "FAE-2", "prefilled_summary": "", "comments": [_comment("提供补丁")]},
        ]
    }
    summaries, low_conf = derive_summaries(data)
    assert "FAE-1" not in summaries
    assert summaries.get("FAE-2")


def test_derive_low_conf_lists_signal_less_issues():
    data = {
        "issues": [
            {"key": "FAE-1", "prefilled_summary": "", "comments": [_comment("等待回复")]},
            {"key": "FAE-2", "prefilled_summary": "", "comments": [_comment("提供补丁")]},
        ]
    }
    _, low_conf = derive_summaries(data)
    assert low_conf == ["FAE-1"]


def test_derive_covers_every_non_prefilled_issue():
    data = {
        "issues": [
            {"key": "FAE-1", "prefilled_summary": "", "comments": []},
            {"key": "FAE-2", "prefilled_summary": "", "comments": [_comment("提供补丁")]},
        ]
    }
    summaries, _ = derive_summaries(data)
    assert summaries["FAE-1"] == "无评论"
    assert "FAE-2" in summaries


def test_limit_comments_keeps_all_in_period():
    comments = [
        _comment("背景1", date="2026-05-01", in_period=False),
        _comment("本期1", date="2026-05-19", in_period=True),
        _comment("本期2", date="2026-05-20", in_period=True),
    ]
    out = limit_comments_for_context(comments)
    assert len(out) == 3
    # in_period kept, newest first
    assert [c["body"] for c in out if c["in_period"]] == ["本期2", "本期1"]


def test_limit_comments_ranks_background_by_signal():
    comments = [
        _comment("等待回复", date="2026-05-01", in_period=False),
        _comment("提供补丁", date="2026-05-02", in_period=False),
    ]
    out = limit_comments_for_context(comments)
    bodies = [c["body"] for c in out]
    # solution-signal comment ranks above the wait-only comment
    assert bodies.index("提供补丁") < bodies.index("等待回复")


def test_limit_comments_caps_single_body():
    comments = [_comment("x" * 300)]
    out = limit_comments_for_context(comments, max_per_comment=250)
    assert len(out[0]["body"]) == 251  # 250 chars + ellipsis
    assert out[0]["body"].endswith("…")


def test_limit_comments_drops_low_value_background_over_budget():
    comments = [
        # two in_period (~6 chars each) leave room; backgrounds fill the rest
        _comment("本期", date="2026-05-20", in_period=True),
        _comment("提供补丁", date="2026-05-02", in_period=False),
        _comment("低价值等待回复", date="2026-05-01", in_period=False),
    ]
    out = limit_comments_for_context(comments, max_per_comment=250, max_total=8)
    bodies = [c["body"] for c in out]
    assert "本期" in bodies
    assert "提供补丁" in bodies
    # budget exhausted before the low-value background comment
    assert "低价值等待回复" not in bodies


def test_limit_comments_empty():
    assert limit_comments_for_context([]) == []
