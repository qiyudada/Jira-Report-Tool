#!/usr/bin/env python3
"""
Jira Report CLI - Command-line interface for Jira Report Tool
Can be used directly or called by Claude Code via skill.

Subcommands:
  run      End-to-end fetch + filter + AI + Excel (default; legacy).
  prepare  Fetch + filter + collect comments, write data.json (no AI).
  export   Read data.json + summaries.json, write Excel (no AI).

When invoked as a Claude Code skill, agents should use `prepare` and `export`
so the agent itself does the AI summary, with no API call.
"""
import argparse
import sys
import os
import datetime
import json

# Add project root to path so `from src import ...` works regardless of cwd.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src import JiraClient, ReportGenerator, Config
from src.ai_providers import get_models, get_default_model, get_label, provider_env_var
from src.env_config import load_env, load_claude_settings, normalize_claude_env_keys


VALID_STATUSES = ["ALL", "WAIT FAE INFO", "WORKED AROUND", "WORKING",
                  "CLOSED", "RESOLVED", "WAIT 3RD PARTY"]

VALID_PROVIDERS = ["deepseek", "openai", "anthropic", "custom"]


def _add_format_flags(parser):
    """Formatting flags shared across subcommands (column order, alignment, etc.)."""
    parser.add_argument("--status", default="ALL", choices=VALID_STATUSES,
                        help="Filter by Jira status (default: ALL)")
    parser.add_argument("--columns", default="1,2,3,4,5,6,7",
                        help="Column order (default: 1,2,3,4,5,6,7)")
    parser.add_argument("--header-align", default="left", choices=["left", "center", "right"],
                        help="Header alignment (default: left)")
    parser.add_argument("--cell-align", default="center", choices=["left", "center", "right"],
                        help="Cell alignment (default: center)")
    parser.add_argument("--no-key-highlight", action="store_true",
                        help="Disable key issue highlighting in red")


def _parse_dates(start_str, end_str):
    try:
        start_date = datetime.datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        print("Error: Invalid date format. Use YYYY-MM-DD", file=sys.stderr)
        sys.exit(1)
    if end_date < start_date:
        print("Error: End date must be >= start date", file=sys.stderr)
        sys.exit(1)
    return start_date, end_date


def _build_config_from_args(args):
    """Layer config sources: CLI args > os.environ > .claude/settings.json > .env > defaults"""
    project_dir = os.path.dirname(os.path.abspath(__file__))

    # --- Layer 4: .env file (lowest file-based priority) ---
    dotenv_path = os.path.join(project_dir, ".env")
    dotenv = load_env(dotenv_path) if os.path.exists(dotenv_path) else {}
    # Apply Claude-style key mappings so .env can use ANTHROPIC_AUTH_TOKEN,
    # ANTHROPIC_BASE_URL, ANTHROPIC_MODEL etc. like .claude/settings.json
    dotenv_claude = normalize_claude_env_keys(dotenv)
    # Merge Claude-mapped values under existing canonical keys (dotenv wins)
    dotenv = {**dotenv_claude, **dotenv}

    # --- Layer 3: .claude/settings.json env block ---
    claude_env = load_claude_settings(project_dir)
    # Map Claude conventions: ANTHROPIC_AUTH_TOKEN -> ANTHROPIC_API_KEY, etc.
    claude_keys = {
        "deepseek": claude_env.get("DEEPSEEK_API_KEY", "") or claude_env.get("DEEPSEEK_AUTH_TOKEN", ""),
        "openai": claude_env.get("OPENAI_API_KEY", ""),
        "anthropic": claude_env.get("ANTHROPIC_API_KEY", "") or claude_env.get("ANTHROPIC_AUTH_TOKEN", ""),
        "custom": claude_env.get("CUSTOM_API_KEY", ""),
    }
    claude_provider = claude_env.get("AI_PROVIDER", "")
    claude_model = claude_env.get("AI_MODEL", "")
    claude_custom_endpoint = claude_env.get("ANTHROPIC_BASE_URL", "") or claude_env.get("OPENAI_BASE_URL", "")

    # --- Resolve provider ---
    provider = (getattr(args, 'ai_provider', None)
                or os.getenv("AI_PROVIDER", "")
                or claude_provider
                or dotenv.get("AI_PROVIDER")
                or "deepseek")

    # --- Resolve API keys (per-source merging) ---
    api_keys = {
        "deepseek": "",
        "openai": "",
        "anthropic": "",
        "custom": "",
    }
    # .env layer
    for p in api_keys:
        api_keys[p] = dotenv.get(f"{p.upper()}_API_KEY", "")
    # .claude/settings.json layer (overwrites .env)
    for p in api_keys:
        if claude_keys.get(p):
            api_keys[p] = claude_keys[p]
    # os.environ layer (overwrites .claude — Claude Code injects these at runtime)
    for p in api_keys:
        env_val = os.getenv(f"{p.upper()}_API_KEY", "")
        if env_val:
            api_keys[p] = env_val
    # Legacy DEEPSEEK_API_KEY env var fallback for deepseek
    if not api_keys.get("deepseek"):
        legacy = os.getenv("DEEPSEEK_API_KEY", "")
        if legacy:
            api_keys["deepseek"] = legacy
    # --ai-key flag (highest)
    ai_key = getattr(args, 'ai_key', None) or ""
    if ai_key:
        api_keys[provider] = ai_key

    # --- Resolve model ---
    ai_model = (getattr(args, 'ai_model', None)
                or os.getenv("AI_MODEL", "")
                or claude_model
                or dotenv.get("AI_MODEL")
                or get_default_model(provider))

    # --- Resolve custom endpoint ---
    custom_endpoint = (getattr(args, 'custom_endpoint', '')
                       or os.getenv("CUSTOM_ENDPOINT", "")
                       or claude_custom_endpoint
                       or dotenv.get("CUSTOM_ENDPOINT", ""))

    # --- Resolve column order ---
    column_order = (getattr(args, 'columns', None)
                    or dotenv.get("COLUMN_ORDER")
                    or "1,2,3,4,5,6,7")

    return Config(
        username=args.username,
        password=args.password,
        ai_model=ai_model,
        ai_provider=provider,
        api_keys=api_keys,
        custom_endpoint=custom_endpoint,
        column_order=Config.normalize_column_order(column_order),
        key_issue_highlight=not args.no_key_highlight,
        comment_timestamp_prefix=getattr(args, 'timestamp_prefix', False),
        header_align=args.header_align,
        cell_align=args.cell_align,
    )


def _login_client(args):
    print(f"Connecting to Jira as {args.username}...")
    client = JiraClient(Config.JIRA_BASE_URL, args.username, args.password)
    if not client.login():
        print("Error: Jira login failed. Check username and password.", file=sys.stderr)
        sys.exit(1)
    print("Login successful.")
    return client


def cmd_run(args):
    """End-to-end one-shot: fetch + filter + (optional AI) + Excel.
    This is the legacy entry point. For the skill/agent flow, use prepare/export."""
    start_date, end_date = _parse_dates(args.start, args.end)

    output_path = args.output
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    config = _build_config_from_args(args)
    provider = config.ai_provider
    model = config.ai_model
    api_key = config.api_keys.get(provider, "")

    # Validate model for non-custom providers
    if provider != "custom":
        valid_models = get_models(provider)
        if valid_models and model not in valid_models:
            print(f"Error: model '{model}' not valid for provider '{provider}'. "
                  f"Valid models: {', '.join(valid_models)}", file=sys.stderr)
            sys.exit(1)

    client = _login_client(args)

    ai_summarizer = None
    use_ai = False
    fetch_comment = bool(getattr(args, 'fetch_comment', False))
    if args.ai and api_key:
        from src import AISummarizer
        provider_label = get_label(provider)
        print(f"AI summarization enabled (provider: {provider_label}, model: {model})")
        ai_summarizer = AISummarizer(api_key=api_key, provider=provider,
                                      custom_endpoint=config.custom_endpoint)
        use_ai = True
    elif args.ai and not api_key:
        env_var = provider_env_var(provider)
        print(f"Warning: --ai specified but no API key found for provider '{provider}'. "
              f"Set --ai-key or {env_var} env var. AI summarization disabled.",
              file=sys.stderr)

    print(f"Generating report for {start_date} to {end_date}...")
    generator = ReportGenerator(client, config)

    try:
        count = generator.generate(
            start_date=start_date,
            end_date=end_date,
            output_path=output_path,
            status_filter=args.status,
            use_ai_summary=use_ai,
            fetch_comment=fetch_comment,
            batch_mode=args.batch_mode,
            batch_size=args.batch_size,
            ai_summarizer=ai_summarizer,
        )
        print(f"Report generated successfully: {count} issues exported to {output_path}")
        print(f"File: {os.path.abspath(output_path)}")
    except Exception as e:
        print(f"Error generating report: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        client.logout()


def cmd_prepare(args):
    """Fetch + filter + collect comments. Write a data.json the agent can read.
    No AI is called. Blocked-status issues (WAIT FAE INFO, etc.) with no
    in-period activity are auto-marked in `prefilled_summary`."""
    start_date, end_date = _parse_dates(args.start, args.end)

    output_path = args.output
    if not output_path.endswith(".json"):
        output_path += ".json"
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    config = _build_config_from_args(args)
    client = _login_client(args)

    print(f"Collecting issues for {start_date} to {end_date} (status={args.status})...")
    generator = ReportGenerator(client, config)

    try:
        data = generator.collect_issues_data(
            start_date=start_date,
            end_date=end_date,
            status_filter=args.status,
        )
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        prefilled = sum(1 for i in data["issues"] if i.get("prefilled_summary"))
        print(f"Data written: {len(data['issues'])} issues ({prefilled} prefilled as blocked) -> {output_path}")
    except Exception as e:
        print(f"Error preparing data: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        client.logout()


def cmd_export(args):
    """Read data.json + summaries.json, write Excel. No Jira login needed.
    `prefilled_summary` (if present) wins over `summaries` for the same issue.
    Missing/empty entries in summaries leave 进展 blank with a stderr warning."""
    input_path = args.input
    summaries_path = args.summaries

    if not os.path.exists(input_path):
        print(f"Error: data file not found: {input_path}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(summaries_path):
        print(f"Error: summaries file not found: {summaries_path}", file=sys.stderr)
        sys.exit(1)

    output_path = args.output
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    with open(summaries_path, "r", encoding="utf-8") as f:
        summaries = json.load(f)

    if not isinstance(summaries, dict):
        print("Error: summaries.json must be a flat {issue_key: summary_text} object.",
              file=sys.stderr)
        sys.exit(1)

    config = Config(
        username="",
        password="",
        ai_model="deepseek-chat",
        column_order=Config.normalize_column_order(args.columns),
        key_issue_highlight=not args.no_key_highlight,
        comment_timestamp_prefix=False,
        header_align=args.header_align,
        cell_align=args.cell_align,
    )
    client = _StubJiraClient()

    generator = ReportGenerator(client, config)
    try:
        count = generator.export_from_data(data, summaries, output_path)
        print(f"Report exported: {count} issues -> {output_path}")
        print(f"File: {os.path.abspath(output_path)}")
    except Exception as e:
        print(f"Error exporting report: {e}", file=sys.stderr)
        sys.exit(1)


class _StubJiraClient:
    """Minimal JiraClient stub for `export`. ReportGenerator.export_from_data
    does not call any network methods; it only needs `base_url` for hyperlinks
    plus a couple of helpers that the Excel-rendering helpers touch."""
    base_url = Config.JIRA_BASE_URL
    username = ""

    @staticmethod
    def _field_to_text(value):
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        if isinstance(value, dict):
            for key in ("name", "value", "summary", "displayName"):
                if key in value and value[key]:
                    return str(value[key])
        if isinstance(value, list):
            return " ".join(self._field_to_text(v) for v in value if v)
        return str(value)

    def is_current_user(self, user):
        return False


def _route_subcommand(argv):
    """If argv[1] is not a known subcommand, prepend 'run' so the legacy
    `python cli.py -u X -p Y --start ...` form continues to work.

    Top-level --help/-h/-V/--version pass through unchanged so the parent
    parser can list all subcommands instead of leaking one subcommand's help.
    """
    subcommands = {"run", "prepare", "export"}
    help_flags = {"-h", "--help", "-V", "--version"}
    if not argv:
        return ["run"]
    if argv[0] in help_flags:
        return list(argv)
    if argv[0] in subcommands:
        return list(argv)
    if argv[0].startswith("-"):
        return ["run"] + list(argv)
    return ["run"] + list(argv)


def main():
    parser = argparse.ArgumentParser(
        description="Generate Jira Report Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Skill/agent flow (recommended for AI agents):
  python cli.py prepare -u USER -p PASS --start 2026-05-01 --end 2026-05-31 -o data.json
  # ... agent reads data.json, writes summaries.json ...
  python cli.py export --input data.json --summaries summaries.json -o report.xlsx

  # One-shot flow with DeepSeek (default provider):
  python cli.py run -u USER -p PASS --start 2026-05-01 --end 2026-05-31 -o report.xlsx --ai

  # One-shot flow with OpenAI:
  python cli.py run -u USER -p PASS --start 2026-05-01 --end 2026-05-31 -o report.xlsx --ai --ai-provider openai --ai-model gpt-4o

  # (also works without the 'run' subcommand for backward compat)
        """,
    )
    subparsers = parser.add_subparsers(dest="subcommand")
    subparsers.required = False
    subparsers.default = "run"

    # --- run (one-shot, default) ---
    p_run = subparsers.add_parser("run", help="end-to-end with optional AI (default)")
    p_run.add_argument("-u", "--username", required=True)
    p_run.add_argument("-p", "--password", required=True)
    p_run.add_argument("--start", required=True)
    p_run.add_argument("--end", required=True)
    p_run.add_argument("-o", "--output", required=True)
    p_run.add_argument("--ai", action="store_true")
    p_run.add_argument("--ai-key", help="API key for the selected provider")
    p_run.add_argument("--ai-provider", default="deepseek", choices=VALID_PROVIDERS,
                       help="AI provider (default: deepseek)")
    p_run.add_argument("--custom-endpoint",
                       help="Custom OpenAI-compatible endpoint URL (for --ai-provider=custom)")
    p_run.add_argument("--ai-model", default=None,
                       help="AI model name (default: provider's default model)")
    p_run.add_argument("--batch-mode", action="store_true")
    p_run.add_argument("--batch-size", type=int, default=10)
    p_run.add_argument("--fetch-comment", action="store_true")
    p_run.add_argument("--timestamp-prefix", action="store_true")
    _add_format_flags(p_run)

    # --- prepare ---
    p_prep = subparsers.add_parser("prepare", help="fetch + filter + JSON for agent")
    p_prep.add_argument("-u", "--username", required=True)
    p_prep.add_argument("-p", "--password", required=True)
    p_prep.add_argument("--start", required=True)
    p_prep.add_argument("--end", required=True)
    p_prep.add_argument("-o", "--output", required=True,
                        help="Path for data.json (must end in .json or will be appended)")
    _add_format_flags(p_prep)

    # --- export ---
    p_exp = subparsers.add_parser("export", help="data.json + summaries.json -> xlsx")
    p_exp.add_argument("--input", required=True, help="Path to data.json from `prepare`")
    p_exp.add_argument("--summaries", required=True,
                       help="Path to summaries.json (flat {issue_key: summary})")
    p_exp.add_argument("-o", "--output", required=True)
    _add_format_flags(p_exp)

    argv = _route_subcommand(sys.argv[1:])
    args = parser.parse_args(argv)
    {
        "run": cmd_run,
        "prepare": cmd_prepare,
        "export": cmd_export,
    }[args.subcommand](args)


if __name__ == "__main__":
    main()
