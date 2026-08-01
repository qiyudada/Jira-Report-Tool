"""
Configuration management for Jira Report Tool
"""
import os


class Config:
    JIRA_BASE_URL = "https://ticket.quectel.com"

    def __init__(self, username: str, password: str,
                 deepseek_api_key: str = None,
                 ai_model: str = "deepseek-chat",
                 ai_provider: str = "deepseek",
                 api_keys: dict = None,
                 custom_endpoint: str = "",
                 column_order: str = "1,2,3,4,5,6,7",
                 key_issue_highlight: bool = True,
                 comment_timestamp_prefix: bool = False,
                 header_align: str = "left",
                 cell_align: str = "center"):
        self.username = username
        self.password = password
        self.ai_model = ai_model
        self.ai_provider = ai_provider
        self.api_keys = api_keys or {}
        # Handle backward-compat deepseek_api_key
        if deepseek_api_key and not self.api_keys.get("deepseek"):
            self.api_keys["deepseek"] = deepseek_api_key
        self.custom_endpoint = custom_endpoint
        self.column_order = column_order
        self.key_issue_highlight = key_issue_highlight
        self.comment_timestamp_prefix = comment_timestamp_prefix
        self.header_align = header_align
        self.cell_align = cell_align

    @property
    def deepseek_api_key(self):
        """Backward-compatible property. Read from api_keys dict."""
        return self.api_keys.get("deepseek", "")

    @deepseek_api_key.setter
    def deepseek_api_key(self, value):
        self.api_keys["deepseek"] = value or ""

    @classmethod
    def from_args(cls, args):
        provider = getattr(args, 'ai_provider', None) or "deepseek"
        api_keys = {
            "deepseek": "",
            "openai": "",
            "anthropic": "",
            "custom": "",
        }
        # --ai-key maps to selected provider
        ai_key = getattr(args, 'deepseek_api_key', None) or getattr(args, 'ai_key', None) or ""
        if ai_key:
            api_keys[provider] = ai_key
        # Also check provider-specific env vars
        _env_map = {
            "deepseek": "DEEPSEEK_API_KEY",
            "openai": "OPENAI_API_KEY",
            "anthropic": "ANTHROPIC_API_KEY",
            "custom": "CUSTOM_API_KEY",
        }
        env_var = _env_map.get(provider, f"{provider.upper()}_API_KEY")
        env_key = os.getenv(env_var, "")
        if not api_keys.get(provider) and env_key:
            api_keys[provider] = env_key
        # Legacy DEEPSEEK_API_KEY fallback for backward compat
        if not api_keys.get("deepseek") and provider == "deepseek":
            api_keys["deepseek"] = os.getenv("DEEPSEEK_API_KEY", "")

        return cls(
            username=getattr(args, 'username', '') or os.getenv('JIRA_USERNAME', ''),
            password=getattr(args, 'password', '') or os.getenv('JIRA_PASSWORD', ''),
            ai_model=getattr(args, 'ai_model', 'deepseek-chat') or 'deepseek-chat',
            ai_provider=provider,
            api_keys=api_keys,
            custom_endpoint=getattr(args, 'custom_endpoint', '') or "",
            column_order=getattr(args, 'column_order', '1,2,3,4,5,6,7') or '1,2,3,4,5,6,7',
            key_issue_highlight=getattr(args, 'key_issue_highlight', True) or True,
            comment_timestamp_prefix=getattr(args, 'comment_timestamp_prefix', False) or False,
            header_align=getattr(args, 'header_align', 'left') or 'left',
            cell_align=getattr(args, 'cell_align', 'center') or 'center',
        )

    @staticmethod
    def normalize_column_order(value):
        default = "1,2,3,4,5,6,7"
        text = str(value or "").strip()
        try:
            nums = [int(x.strip()) for x in text.split(",")]
            if len(nums) != 7:
                return default
            if set(nums) != set(range(1, 8)):
                return default
            return ",".join(str(x) for x in nums)
        except:
            return default
