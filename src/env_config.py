"""
.env file parser and writer.
No external dependencies — pure stdlib implementation.
"""

import os


def load_env(path, defaults=None):
    """Load a .env file into a dict.

    Format: KEY=value per line. # starts a comment. Empty lines skipped.
    Everything after the first = is the value. Trailing whitespace stripped.

    Returns dict with defaults merged underneath the loaded values.
    """
    result = dict(defaults or {})

    if not os.path.exists(path):
        return result

    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip()
            if key:
                result[key] = coerce_bool(value)

    return result


def save_env(path, data):
    """Write a dict to a .env file.

    Existing .env values are preserved — only keys present in *data* are
    updated; every other key in the current .env is kept as-is.  This
    prevents Claude Code compatibility keys (ANTHROPIC_AUTH_TOKEN, …)
    from being silently dropped when the UI saves only canonical keys.

    Keys are written alphabetically grouped by section with comment headers.
    Strings, bools, and other scalars are written as-is (str() conversion).
    """
    # Merge with existing .env so unknown / compat keys survive
    existing = {}
    if os.path.exists(path):
        existing = load_env(path)
    merged = dict(existing)
    merged.update(data)

    lines = []
    # Group keys into sections for readability — mirrors .env layout
    sections = [
        ("# ============================================================", []),
        ("# Jira 认证 (Jira Credentials)", []),
        ("# ============================================================", []),
        ("", ["JIRA_USERNAME", "JIRA_PASSWORD"]),
        ("", []),
        ("# ============================================================", []),
        ("# AI 提供商 - 配置 (Provider)", []),
        ("# ============================================================", []),
        ("", ["AI_PROVIDER", "AI_API_MODE"]),
        ("", []),
        ("# ============================================================", []),
        ("# API 密钥 - 官方接入配置 (Official API Keys)", []),
        ("# ============================================================", []),
        ("", ["DEEPSEEK_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY", "CUSTOM_API_KEY"]),
        ("", []),
        ("# ============================================================", []),
        ("# 模型与端点 - 官方接入配置 (Model & Endpoint)", []),
        ("# ============================================================", []),
        ("", ["AI_MODEL", "CUSTOM_ENDPOINT"]),
        ("", []),
        ("# ============================================================", []),
        ("# 第三方 / Claude Code 兼容配置 (Third-party Proxy)", []),
        ("# ============================================================", []),
        ("# --- 认证令牌 ---", [
            "ANTHROPIC_AUTH_TOKEN", "DEEPSEEK_AUTH_TOKEN", "OPENAI_AUTH_TOKEN",
        ]),
        ("", []),
        ("# --- 第三方 API 端点 ---", [
            "ANTHROPIC_BASE_URL", "OPENAI_BASE_URL",
        ]),
        ("", []),
        ("# --- 模型选择 ---", [
            "ANTHROPIC_MODEL", "ANTHROPIC_DEFAULT_OPUS_MODEL",
            "ANTHROPIC_DEFAULT_SONNET_MODEL", "ANTHROPIC_DEFAULT_HAIKU_MODEL",
            "CLAUDE_CODE_SUBAGENT_MODEL",
        ]),
        ("", []),
        ("# --- 其他 Claude Code 兼容项 ---", [
            "CLAUDE_CODE_EFFORT_LEVEL",
        ]),
        ("", []),
        ("# ============================================================", []),
        ("# 报表格式 (Report Formatting)", []),
        ("# ============================================================", []),
        ("", [
            "COLUMN_ORDER", "KEY_ISSUE_HIGHLIGHT", "COMMENT_TIMESTAMP_PREFIX",
            "THEME", "LANGUAGE",
        ]),
        ("", []),
        ("# ============================================================", []),
        ("# 其他 (Other)", []),
        ("# ============================================================", []),
        ("", ["LAST_SAVE_DIR"]),
    ]
    written = set()

    for header, keys in sections:
        if header:
            lines.append(header)
        elif not keys:
            # Empty header + empty keys = blank line separator
            lines.append("")
        for key in keys:
            if key in merged:
                val = merged[key]
                if isinstance(val, bool):
                    val = str(val).lower()
                lines.append(f"{key}={val}")
                written.add(key)

    # Write any remaining keys that weren't in sections
    remaining = [(k, v) for k, v in merged.items() if k not in written]
    if remaining:
        lines.append("")
        lines.append("# Other")
        for key, val in sorted(remaining):
            if isinstance(val, bool):
                val = str(val).lower()
            lines.append(f"{key}={val}")

    lines.append("")  # trailing newline
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def coerce_bool(value):
    """Convert string 'true'/'false' to Python bool, return string otherwise."""
    if value.lower() == "false":
        return False
    if value.lower() == "true":
        return True
    return value


# .env key to internal field mapping for migration from old .jira_config JSON
JSON_TO_ENV_MAP = {
    "username": "JIRA_USERNAME",
    "password": "JIRA_PASSWORD",
    "ai_provider": "AI_PROVIDER",
    "ai_model": "AI_MODEL",
    "custom_endpoint": "CUSTOM_ENDPOINT",
    "column_order": "COLUMN_ORDER",
    "key_issue_highlight": "KEY_ISSUE_HIGHLIGHT",
    "comment_timestamp_prefix": "COMMENT_TIMESTAMP_PREFIX",
    "theme": "THEME",
    "language": "LANGUAGE",
    "last_save_dir": "LAST_SAVE_DIR",
}

# Old JSON api_keys -> .env key mapping
API_KEYS_ENV_MAP = {
    "deepseek": "DEEPSEEK_API_KEY",
    "openai": "OPENAI_API_KEY",
    "anthropic": "ANTHROPIC_API_KEY",
    "custom": "CUSTOM_API_KEY",
}


def migrate_json_to_env(json_path, env_path):
    """Read old .jira_config JSON and write .env file.
    Returns the loaded env dict.
    """
    import json
    if not os.path.exists(json_path):
        return {}

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    env_data = {}

    # Map simple fields
    for json_key, env_key in JSON_TO_ENV_MAP.items():
        if json_key in data:
            val = data[json_key]
            if isinstance(val, bool):
                val = str(val).lower()
            env_data[env_key] = str(val) if val is not None else ""

    # Map api_keys dict (new format) or old deepseek_api_key
    if "api_keys" in data:
        for provider, env_key in API_KEYS_ENV_MAP.items():
            env_data[env_key] = data["api_keys"].get(provider, "")
    elif "deepseek_api_key" in data:
        env_data["DEEPSEEK_API_KEY"] = data.get("deepseek_api_key", "")

    # Ensure all API key slots exist
    for env_key in API_KEYS_ENV_MAP.values():
        env_data.setdefault(env_key, "")

    # Set defaults for any missing fields
    env_data.setdefault("AI_PROVIDER", "deepseek")
    env_data.setdefault("AI_API_MODE", "official")
    env_data.setdefault("AI_MODEL", "deepseek-chat")
    env_data.setdefault("COLUMN_ORDER", "1,2,3,4,5,6,7")
    env_data.setdefault("KEY_ISSUE_HIGHLIGHT", "true")
    env_data.setdefault("COMMENT_TIMESTAMP_PREFIX", "false")
    env_data.setdefault("THEME", "Geek")
    env_data.setdefault("LANGUAGE", "zh")
    env_data.setdefault("LAST_SAVE_DIR", os.path.expanduser("~"))
    env_data.setdefault("CUSTOM_ENDPOINT", "")

    save_env(env_path, env_data)
    return env_data


def normalize_claude_env_keys(dotenv):
    """Map Claude-style keys in a .env dict to canonical keys.

    This allows .env to use the same key names as Claude Code's
    ``.claude/settings.json`` env block (ANTHROPIC_AUTH_TOKEN,
    ANTHROPIC_BASE_URL, etc.) and have them resolve correctly.

    Returns a new dict — does not mutate the input.
    Canonical keys already present in dotenv take precedence over
    mapped values, so explicit ANTHROPIC_API_KEY beats ANTHROPIC_AUTH_TOKEN.
    """
    result = {}

    # Resolve API keys — Claude-style auth tokens map to canonical API key names
    for canonical, claude_key in (
        ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN"),
        ("DEEPSEEK_API_KEY", "DEEPSEEK_AUTH_TOKEN"),
        ("OPENAI_API_KEY", "OPENAI_AUTH_TOKEN"),
    ):
        claude_val = dotenv.get(claude_key, "")
        if claude_val:
            result[canonical] = str(claude_val)

    # Resolve AI_PROVIDER — default to "anthropic" if ANTHROPIC_AUTH_TOKEN present
    if "AI_PROVIDER" in dotenv:
        result["AI_PROVIDER"] = dotenv["AI_PROVIDER"]
    elif dotenv.get("ANTHROPIC_AUTH_TOKEN") and "AI_PROVIDER" not in result:
        result["AI_PROVIDER"] = "anthropic"

    # Resolve AI_MODEL — Claude model keys
    for claude_key in (
        "ANTHROPIC_MODEL",
        "ANTHROPIC_DEFAULT_SONNET_MODEL",
        "ANTHROPIC_DEFAULT_OPUS_MODEL",
        "ANTHROPIC_DEFAULT_HAIKU_MODEL",
        "CLAUDE_CODE_SUBAGENT_MODEL",
    ):
        val = dotenv.get(claude_key, "")
        if val and "AI_MODEL" not in result:
            result["AI_MODEL"] = str(val)

    if "AI_MODEL" in dotenv:
        result["AI_MODEL"] = dotenv["AI_MODEL"]

    # Resolve custom endpoint
    for endpoint_key in ("ANTHROPIC_BASE_URL", "OPENAI_BASE_URL"):
        val = dotenv.get(endpoint_key, "")
        if val and "CUSTOM_ENDPOINT" not in result:
            result["CUSTOM_ENDPOINT"] = str(val)

    if "CUSTOM_ENDPOINT" in dotenv:
        result["CUSTOM_ENDPOINT"] = dotenv["CUSTOM_ENDPOINT"]

    return result


def load_claude_settings(project_root=None):
    """Read .claude/settings.json and .claude/settings.local.json, extract the
    ``env`` block, and merge into a dict keyed by our canonical .env names.

    * settings.local.json wins over settings.json.
    * Claude-naming conventions are mapped to our canonical keys:

        ANTHROPIC_AUTH_TOKEN     -> ANTHROPIC_API_KEY
        ANTHROPIC_BASE_URL       -> CUSTOM_ENDPOINT  (when AI_PROVIDER is anthropic or custom)
        ANTHROPIC_MODEL          -> AI_MODEL          (when AI_PROVIDER is anthropic)
        OPENAI_BASE_URL          -> CUSTOM_ENDPOINT   (when AI_PROVIDER is openai or custom)
        CLAUDE_CODE_SUBAGENT_MODEL -> AI_MODEL        (generic fallback)

    Returns dict, empty if nothing found.
    """
    if project_root is None:
        import os as _os
        project_root = _os.getcwd()

    claude_dir = os.path.join(project_root, ".claude")
    settings_path = os.path.join(claude_dir, "settings.json")
    local_path = os.path.join(claude_dir, "settings.local.json")

    merged = {}

    def _load_json(path):
        import json
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    # Base settings first, then local overrides
    for path in (settings_path, local_path):
        data = _load_json(path)
        env_block = data.get("env", {})
        if isinstance(env_block, dict):
            merged.update(env_block)

    if not merged:
        return {}

    # Map Claude conventions -> our canonical keys
    result = {}

    for key, value in merged.items():
        val = str(value) if not isinstance(value, (bool, type(None))) else value
        if isinstance(val, bool):
            val = str(val).lower()

        # Direct pass-through for matching key names
        if key in ("DEEPSEEK_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY",
                    "CUSTOM_API_KEY", "AI_PROVIDER", "AI_MODEL",
                    "COLUMN_ORDER", "THEME", "LANGUAGE", "JIRA_USERNAME",
                    "JIRA_PASSWORD", "DEEPSEEK_AUTH_TOKEN"):
            result[key] = str(val) if val is not None else ""
            continue

        # Map Claude-specific names
        if key == "ANTHROPIC_AUTH_TOKEN":
            result["ANTHROPIC_API_KEY"] = str(val) if val is not None else ""
        elif key == "ANTHROPIC_BASE_URL":
            result["ANTHROPIC_BASE_URL"] = str(val) if val is not None else ""
        elif key == "ANTHROPIC_MODEL":
            result["ANTHROPIC_MODEL"] = str(val) if val is not None else ""
        elif key == "ANTHROPIC_DEFAULT_OPUS_MODEL":
            result["ANTHROPIC_DEFAULT_OPUS_MODEL"] = str(val) if val is not None else ""
        elif key == "ANTHROPIC_DEFAULT_SONNET_MODEL":
            result["ANTHROPIC_DEFAULT_SONNET_MODEL"] = str(val) if val is not None else ""
        elif key == "ANTHROPIC_DEFAULT_HAIKU_MODEL":
            result["ANTHROPIC_DEFAULT_HAIKU_MODEL"] = str(val) if val is not None else ""
        elif key == "OPENAI_BASE_URL":
            result["OPENAI_BASE_URL"] = str(val) if val is not None else ""
        elif key == "CLAUDE_CODE_SUBAGENT_MODEL":
            result["CLAUDE_CODE_SUBAGENT_MODEL"] = str(val) if val is not None else ""
        elif key == "CLAUDE_CODE_EFFORT_LEVEL":
            result["CLAUDE_CODE_EFFORT_LEVEL"] = str(val) if val is not None else ""
        else:
            # Unknown key — pass through in case it's a future env var
            result[key] = str(val) if val is not None else ""

    return result
