"""
AI Provider abstraction layer.
Defines provider specs and a unified `call_ai()` that handles
OpenAI Chat Completions and Anthropic Messages formats.
"""
import requests


PROVIDERS = {
    "deepseek": {
        "label": "DeepSeek",
        "endpoint": "https://api.deepseek.com/chat/completions",
        "auth_type": "bearer",
        "api_format": "openai_chat",
        "models": ["deepseek-chat", "deepseek-coder", "deepseek-v4-flash", "deepseek-v4-pro"],
        "default_model": "deepseek-chat",
    },
    "openai": {
        "label": "OpenAI",
        "endpoint": "https://api.openai.com/v1/chat/completions",
        "auth_type": "bearer",
        "api_format": "openai_chat",
        "models": ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo", "o1", "o1-mini", "o3-mini"],
        "default_model": "gpt-4o",
    },
    "anthropic": {
        "label": "Anthropic (Claude)",
        "endpoint": "https://api.anthropic.com/v1/messages",
        "auth_type": "x-api-key",
        "api_format": "anthropic_messages",
        "api_version": "2023-06-01",
        "models": [
            "claude-sonnet-4-6-20250514",
            "claude-opus-4-6-20250624",
            "claude-haiku-4-5-20251001",
        ],
        "default_model": "claude-sonnet-4-6-20250514",
    },
    "custom": {
        "label": "Custom (OpenAI-compatible)",
        "endpoint": None,
        "auth_type": "bearer",
        "api_format": "openai_chat",
        "models": [],
        "default_model": "",
    },
}


def get_provider_spec(provider_id):
    """Return the spec dict for a provider id, or None."""
    return PROVIDERS.get(provider_id)


def get_models(provider_id):
    """Return model list for a provider."""
    spec = PROVIDERS.get(provider_id)
    return spec["models"] if spec else []


def get_default_model(provider_id):
    """Return the default model for a provider."""
    spec = PROVIDERS.get(provider_id)
    return spec["default_model"] if spec else ""


def get_label(provider_id):
    """Return display label for a provider."""
    spec = PROVIDERS.get(provider_id)
    return spec["label"] if spec else provider_id


def provider_env_var(provider_id):
    """Return the environment variable name for a provider's API key."""
    mapping = {
        "deepseek": "DEEPSEEK_API_KEY",
        "openai": "OPENAI_API_KEY",
        "anthropic": "ANTHROPIC_API_KEY",
        "custom": "CUSTOM_API_KEY",
    }
    return mapping.get(provider_id, f"{provider_id.upper()}_API_KEY")


def call_ai(provider_id, api_key, model, prompt, max_tokens=500,
            temperature=0.3, timeout=60, custom_endpoint=None):
    """Make an AI API call and return normalized result.

    Returns:
        {"ok": True, "content": "..."} on success
        {"ok": False, "error": "...", "status_code": int} on failure
    """
    spec = get_provider_spec(provider_id)
    if not spec:
        return {"ok": False, "error": f"Unknown provider: {provider_id}", "status_code": 0}

    # Resolve endpoint — custom_endpoint overrides spec endpoint for any provider
    endpoint = custom_endpoint or spec["endpoint"]
    if not endpoint:
        return {"ok": False, "error": "No endpoint configured for custom provider", "status_code": 0}

    # Auto-append standard API path when the custom endpoint is a bare host.
    # Examples: "https://proxy.example.com"   → ".../v1/chat/completions"
    #           "https://proxy.example.com/v1" → ".../v1/chat/completions"
    #           "https://proxy.example.com/v1/chat/completions" → unchanged
    if custom_endpoint:
        # Already has the full path — don't touch
        if any(p in custom_endpoint for p in ("/chat/completions", "/messages")):
            pass
        # Ends with /v1 — append only the resource path
        elif custom_endpoint.rstrip("/").endswith("/v1"):
            _suffix = "/messages" if spec["api_format"] == "anthropic_messages" else "/chat/completions"
            endpoint = custom_endpoint.rstrip("/") + _suffix
        # Bare host — append /v1/<resource>
        else:
            _suffix = "/v1/messages" if spec["api_format"] == "anthropic_messages" else "/v1/chat/completions"
            endpoint = custom_endpoint.rstrip("/") + _suffix

    # Build headers
    if spec["auth_type"] == "x-api-key":
        headers = {
            "x-api-key": api_key,
            "Content-Type": "application/json",
            "anthropic-version": spec.get("api_version", "2023-06-01"),
        }
    else:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

    # Build payload
    if spec["api_format"] == "anthropic_messages":
        payload = {
            "model": model,
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": prompt}],
        }
    else:
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_tokens,
            "temperature": temperature,
        }

    try:
        response = requests.post(endpoint, headers=headers, json=payload, timeout=timeout)
    except requests.exceptions.Timeout:
        return {"ok": False, "error": "Request timed out", "status_code": 0, "endpoint": endpoint}
    except requests.exceptions.ConnectionError:
        return {"ok": False, "error": f"Connection failed to {endpoint}", "status_code": 0, "endpoint": endpoint}
    except Exception as e:
        return {"ok": False, "error": str(e), "status_code": 0, "endpoint": endpoint}

    # Parse response
    try:
        result = response.json()
    except ValueError:
        return {"ok": False, "error": f"Invalid JSON response from {endpoint}", "status_code": response.status_code, "endpoint": endpoint}

    if "error" in result:
        err_detail = result["error"]
        if isinstance(err_detail, dict):
            err_detail = err_detail.get("message", str(err_detail))
        return {"ok": False, "error": f"[{endpoint}] {err_detail}", "status_code": response.status_code, "endpoint": endpoint}

    if response.status_code == 200:
        # Extract content based on format
        if spec["api_format"] == "anthropic_messages":
            content_blocks = result.get("content", [])
            if content_blocks and isinstance(content_blocks, list):
                content = content_blocks[0].get("text", "")
            else:
                content = ""
        else:
            choices = result.get("choices", [])
            if choices:
                content = choices[0].get("message", {}).get("content", "").strip()
            else:
                content = ""

        if content:
            return {"ok": True, "content": content, "endpoint": endpoint}
        else:
            return {"ok": False, "error": "Empty response content", "status_code": 200, "endpoint": endpoint}

    # HTTP error codes
    error_map = {401: "API Key 无效", 429: "请求超过限额", 400: "请求参数错误"}
    return {
        "ok": False,
        "error": f"[{endpoint}] {error_map.get(response.status_code, f'HTTP {response.status_code}')}",
        "status_code": response.status_code,
        "endpoint": endpoint,
    }
