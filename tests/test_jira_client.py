"""Tests for JiraClient, focused on the per-issue comment cache."""
from src.jira_client import JiraClient


class _FakeResponse:
    def __init__(self, payload):
        self.status_code = 200
        self._payload = payload

    def json(self):
        return self._payload


def test_get_comments_is_cached(monkeypatch):
    client = JiraClient("https://example.com", "user@example.com", "pass")
    calls = {"n": 0}

    def fake_get(url, timeout=30):
        calls["n"] += 1
        return _FakeResponse({"comments": [{"id": "c1"}]})

    monkeypatch.setattr(client.session, "get", fake_get)

    first = client.get_comments("FAE-1")
    second = client.get_comments("FAE-1")

    assert first == second == [{"id": "c1"}]
    assert calls["n"] == 1  # second call served from cache


def test_get_comments_failure_not_cached(monkeypatch):
    client = JiraClient("https://example.com", "user@example.com", "pass")
    calls = {"n": 0}

    class _ErrorResponse:
        status_code = 500

        def json(self):
            raise ValueError("no body")

    def fake_get(url, timeout=30):
        calls["n"] += 1
        return _ErrorResponse()

    monkeypatch.setattr(client.session, "get", fake_get)

    assert client.get_comments("FAE-1") == []
    assert client.get_comments("FAE-1") == []
    # a failed fetch is NOT cached, so the next call retries
    assert calls["n"] == 2
