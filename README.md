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

The CLI exposes three subcommands. The `run` subcommand is the legacy one-shot
path (still the default if no subcommand is given). The `prepare` / `export`
pair is the recommended path for AI agents: the agent itself writes the AI
summary between them, so no DeepSeek key is needed.

### `prepare` + `export` (recommended for AI agents)

```bash
# Step 1: pull data and write data.json
python cli.py prepare -u user@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 -o /tmp/data.json

# Step 2: agent reads /tmp/data.json, writes /tmp/summaries.json
#   (see skills/jira_report/SKILL.md for the summary rules)

# Step 3: render the Excel
python cli.py export --input /tmp/data.json \
                     --summaries /tmp/summaries.json \
                     -o ~/Downloads/jira_report.xlsx
```

`data.json` includes `prefilled_summary` for issues that are in a blocked
status (WAIT FAE INFO / WORKED AROUND / WAIT OFFICIAL RELEASE) with no
in-period activity. The agent must skip those — `export` uses the prefilled
text and ignores any agent summary for the same key.

### `run` (legacy one-shot, uses DeepSeek for AI)

The same flags as before, just under an explicit `run` subcommand (which is
the default, so old invocations still work).

```bash
DEEPSEEK_API_KEY=your_api_key python cli.py run \
  -u user@quectel.com -p password123 \
  --start 2026-05-01 --end 2026-05-31 \
  -o ~/Downloads/jira_report.xlsx --ai
```

### CLI Options

#### Common (all subcommands)
| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `-u/--username` | Yes | - | Jira username (email format) |
| `-p/--password` | Yes | - | Jira password |
| `--start` | Yes | - | Start date (YYYY-MM-DD) |
| `--end` | Yes | - | End date (YYYY-MM-DD) |
| `--status` | No | ALL | Filter by status |
| `--columns` | No | 1,2,3,4,5,6,7 | Column order |
| `--header-align` | No | left | Header alignment |
| `--cell-align` | No | center | Cell alignment |
| `--no-key-highlight` | No | - | Disable key issue highlighting |

#### `run` only (DeepSeek AI)
| Option | Description |
|--------|-------------|
| `--ai` | Enable AI summarization via DeepSeek |
| `--ai-key` | DeepSeek API key (or set `DEEPSEEK_API_KEY` env var) |
| `--ai-model` | DeepSeek model (default: `deepseek-chat`) |
| `--batch-mode` | Use batch AI mode |
| `--batch-size` | AI batch size (default: 10) |
| `--fetch-comment` | Fetch latest comment as progress |
| `--timestamp-prefix` | Add timestamp prefix to comments |

#### `export` only
| Option | Description |
|--------|-------------|
| `--input` | Path to `data.json` from `prepare` |
| `--summaries` | Path to `summaries.json` (flat `{issue_key: summary}`) |

## Skill Mode vs GUI Mode

| | GUI (`jira_report_generator.py`) | Skill (agent + `cli.py prepare/export`) |
|---|---|---|
| AI summary | DeepSeek API | **Agent itself** (no external LLM) |
| Setup | Needs `DEEPSEEK_API_KEY` | No LLM key needed |
| Use case | Manual weekly report | Programmatic / agent workflow |
| Multi-issue batching | `--batch-mode` | Agent batches naturally |

Both modes produce identical Excel output (columns, hyperlinks, dropdowns,
red highlighting for key issues).

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