# Jira Report

A tool for generating Excel reports from Jira issues. Supports both **Desktop GUI** and **Claude Code skill** invocation.

## Installation

```bash
pip install -r requirements.txt
```

## Two Usage Modes

### 1. Desktop GUI App

```bash
python jira_report_generator.py
```

### 2. Claude Code Skill (Recommended for AI workflow)

After installation, Claude Code can automatically invoke this skill when you ask to generate a Jira report.

**Example prompts:**
- "生成Jira周报"
- "Create a Jira Excel report for May 2026"
- "Generate my weekly Jira report"

Claude Code will guide you through the required parameters and invoke the CLI automatically.

## CLI Usage

```bash
python cli.py -u <username> -p <password> --start <YYYY-MM-DD> --end <YYYY-MM-DD> -o <output.xlsx>
```

**Example:**
```bash
python cli.py -u user@quectel.com -p password123 \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx
```

**With AI summarization:**
```bash
DEEPSEEK_API_KEY=your_api_key python cli.py \
  -u user@quectel.com -p password123 \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx --ai
```

### CLI Options

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `-u/--username` | Yes | - | Jira username (email format) |
| `-p/--password` | Yes | - | Jira password |
| `--start` | Yes | - | Start date (YYYY-MM-DD) |
| `--end` | Yes | - | End date (YYYY-MM-DD) |
| `-o/--output` | Yes | - | Output Excel file path |
| `--status` | No | ALL | Filter by status |
| `--ai` | No | - | Enable AI summarization |
| `--ai-model` | No | deepseek-chat | DeepSeek model |
| `--batch-mode` | No | - | Use batch AI mode |
| `--batch-size` | No | 10 | AI batch size |
| `--fetch-comment` | No | - | Fetch latest comment as progress |
| `--columns` | No | 1,2,3,4,5,6,7 | Column order |

## Pack to EXE (Optional)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed jira_report_generator.py
```

The executable will be generated in the `dist/` folder.

## Configuration (`.jira_config`)

The desktop app stores settings in `.jira_config` (JSON):

```json
{
  "username": "your.name@example.com",
  "password": "your_jira_password",
  "last_save_dir": "C:/Users/YourName/Downloads",
  "deepseek_api_key": "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  "ai_model": "deepseek-chat"
}
```

| Field | Description |
|-------|-------------|
| `username` | Jira login email |
| `password` | Jira login password |
| `last_save_dir` | Default directory for saving the Excel report |
| `deepseek_api_key` | DeepSeek API Key for AI summary feature |
| `ai_model` | DeepSeek model to use |

## AI Summary (Optional)

The app supports AI-powered comment summarization via the [DeepSeek API](https://platform.deepseek.com/).

**Setup:**

1. Register at [platform.deepseek.com](https://platform.deepseek.com/) and create an API Key.
2. Set via environment variable: `export DEEPSEEK_API_KEY=your_key`
3. Or use `--ai-key` flag in CLI

**Supported models:**

| Model | Description |
|-------|-------------|
| `deepseek-chat` | General-purpose chat model (default, recommended) |
| `deepseek-coder` | Optimized for code-related content |
| `deepseek-v4-flash` | Faster, lower-cost variant |
| `deepseek-v4-pro` | Higher capability, higher cost |

## Report Output

Generated Excel contains:

| Column | Description |
|--------|-------------|
| 客户名称 | Customer name |
| 型号 | Model |
| 问题描述 | Issue description (with hyperlink) |
| JIRA号 | Issue key |
| 状态 | Status |
| 是否为重点问题 | Key issue flag (A/B/C/D/E or 是/否) |
| 进展 | Progress (AI summary or latest comment) |

## Quick Start (Desktop App)

1. Run `python jira_report_generator.py`
2. Enter your Jira credentials and click "Login"
3. Select a date range or use "This Week" / "This Month" quick buttons
4. (Optional) Check "Use AI Summary"
5. Choose a save location for the Excel file
6. Click "Generate Report"