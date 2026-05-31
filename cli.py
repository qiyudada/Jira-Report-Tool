#!/usr/bin/env python3
"""
Jira Report CLI - Command-line interface for Jira Report Tool
Can be used directly or called by Claude Code via skill
"""
import argparse
import sys
import os
import datetime

# Add src to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src import JiraClient, ReportGenerator, AISummarizer, Config


def main():
    parser = argparse.ArgumentParser(
        description="Generate Jira Report Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python cli.py -u username@quectel.com -p password --start 2026-05-01 --end 2026-05-31 -o report.xlsx
  python cli.py -u username@quectel.com -p password --start 2026-05-01 --end 2026-05-31 -o report.xlsx --ai
  python cli.py -u username@quectel.com -p password --start 2026-05-01 --end 2026-05-31 -o report.xlsx --ai --batch-size 20
        """
    )

    # Required
    parser.add_argument("-u", "--username", required=True, help="Jira username/email")
    parser.add_argument("-p", "--password", required=True, help="Jira password")
    parser.add_argument("--start", required=True, help="Start date (YYYY-MM-DD)")
    parser.add_argument("--end", required=True, help="End date (YYYY-MM-DD)")
    parser.add_argument("-o", "--output", required=True, help="Output Excel file path")

    # Optional
    parser.add_argument("--status", default="ALL",
                       choices=["ALL", "WAIT FAE INFO", "WORKED AROUND", "WORKING",
                               "CLOSED", "RESOLVED", "WAIT 3RD PARTY"],
                       help="Filter by Jira status (default: ALL)")
    parser.add_argument("--columns", default="1,2,3,4,5,6,7",
                       help="Column order (default: 1,2,3,4,5,6,7)")
    parser.add_argument("--header-align", default="left", choices=["left", "center", "right"],
                       help="Header alignment (default: left)")
    parser.add_argument("--cell-align", default="center", choices=["left", "center", "right"],
                       help="Cell alignment (default: center)")
    parser.add_argument("--no-key-highlight", action="store_true",
                       help="Disable key issue highlighting in red")

    # AI options
    parser.add_argument("--ai", action="store_true", help="Enable AI summarization")
    parser.add_argument("--ai-key", help="DeepSeek API key (or set DEEPSEEK_API_KEY env var)")
    parser.add_argument("--ai-model", default="deepseek-chat",
                       choices=["deepseek-chat", "deepseek-coder", "deepseek-v4-flash", "deepseek-v4-pro"],
                       help="DeepSeek model (default: deepseek-chat)")
    parser.add_argument("--batch-mode", action="store_true", help="Use batch AI summarization")
    parser.add_argument("--batch-size", type=int, default=10, help="Batch size for AI summarization")

    # Comment options
    parser.add_argument("--fetch-comment", action="store_true",
                       help="Fetch latest comment as progress content")
    parser.add_argument("--timestamp-prefix", action="store_true",
                       help="Add timestamp prefix to comments")

    args = parser.parse_args()

    # Validate dates
    try:
        start_date = datetime.datetime.strptime(args.start, "%Y-%m-%d").date()
        end_date = datetime.datetime.strptime(args.end, "%Y-%m-%d").date()
    except ValueError:
        print("Error: Invalid date format. Use YYYY-MM-DD", file=sys.stderr)
        sys.exit(1)

    if end_date < start_date:
        print("Error: End date must be >= start date", file=sys.stderr)
        sys.exit(1)

    # Ensure output has .xlsx extension
    output_path = args.output
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"

    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Get DeepSeek API key
    api_key = args.ai_key or os.getenv("DEEPSEEK_API_KEY", "")

    # Build config
    config = Config(
        username=args.username,
        password=args.password,
        deepseek_api_key=api_key,
        ai_model=args.ai_model,
        column_order=Config.normalize_column_order(args.columns),
        key_issue_highlight=not args.no_key_highlight,
        comment_timestamp_prefix=args.timestamp_prefix,
        header_align=args.header_align,
        cell_align=args.cell_align,
    )

    # Initialize client and login
    print(f"Connecting to Jira as {args.username}...")
    client = JiraClient(Config.JIRA_BASE_URL, args.username, args.password)

    if not client.login():
        print("Error: Jira login failed. Check username and password.", file=sys.stderr)
        sys.exit(1)

    print("Login successful.")

    # Initialize AI summarizer if enabled
    ai_summarizer = None
    if args.ai and api_key:
        print(f"AI summarization enabled (model: {args.ai_model})")
        ai_summarizer = AISummarizer(api_key)
    elif args.ai and not api_key:
        print("Warning: --ai specified but DEEPSEEK_API_KEY not set. AI summarization disabled.")

    # Generate report
    print(f"Generating report for {start_date} to {end_date}...")

    generator = ReportGenerator(client, config)

    try:
        count = generator.generate(
            start_date=start_date,
            end_date=end_date,
            output_path=output_path,
            status_filter=args.status,
            use_ai_summary=args.ai and ai_summarizer is not None,
            fetch_comment=args.fetch_comment,
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

    sys.exit(0)


if __name__ == "__main__":
    main()