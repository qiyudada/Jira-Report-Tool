"""
Configuration management for Jira Report Tool
"""
import os


class Config:
    JIRA_BASE_URL = "https://ticket.quectel.com"

    def __init__(self, username: str, password: str,
                 deepseek_api_key: str = None,
                 ai_model: str = "deepseek-chat",
                 column_order: str = "1,2,3,4,5,6,7",
                 key_issue_highlight: bool = True,
                 comment_timestamp_prefix: bool = False,
                 header_align: str = "left",
                 cell_align: str = "center"):
        self.username = username
        self.password = password
        self.deepseek_api_key = deepseek_api_key or ""
        self.ai_model = ai_model
        self.column_order = column_order
        self.key_issue_highlight = key_issue_highlight
        self.comment_timestamp_prefix = comment_timestamp_prefix
        self.header_align = header_align
        self.cell_align = cell_align

    @classmethod
    def from_args(cls, args):
        return cls(
            username=getattr(args, 'username', '') or os.getenv('JIRA_USERNAME', ''),
            password=getattr(args, 'password', '') or os.getenv('JIRA_PASSWORD', ''),
            deepseek_api_key=getattr(args, 'deepseek_api_key', None) or os.getenv('DEEPSEEK_API_KEY', ''),
            ai_model=getattr(args, 'ai_model', 'deepseek-chat') or 'deepseek-chat',
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