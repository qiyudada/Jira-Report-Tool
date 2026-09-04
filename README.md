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

The skill walks you through the required options interactively, then invokes the
CLI automatically. It asks for: **report period**, **status filter**,
**progress source (AI summary vs. latest-comment verbatim)**, and **output
path**.

The skill's AI summary runs on **your local Claude** (Claude Code's own
credential), not on any key in `.env`. The `.env` API keys are only used by the
GUI / `cli.py run --ai` path.

## CLI Usage

The CLI exposes four subcommands. The `run` subcommand is the legacy one-shot
path (still the default if no subcommand is given). The `prepare` / `derive` /
`export` flow is the recommended path for AI agents: `derive` deterministically
extracts progress (no LLM), and the agent only reviews/refines low-confidence
issues, so no DeepSeek key is needed.

### `prepare` + `derive` + `export` (recommended for AI agents)

```bash
# Step 1: pull data and write data.json
python cli.py prepare -u user@quectel.com -p password \
  --start 2026-05-01 --end 2026-05-31 -o /tmp/data.json

# Step 2: deterministic progress extraction (no LLM) — writes a candidate
#         summaries.json plus a low-confidence list on stdout
python cli.py derive --input /tmp/data.json -o /tmp/summaries.json

# Step 3: agent reviews/refines the low-confidence entries in summaries.json
#   (see skills/jira_report/SKILL.md for the review rules)

# Step 4: render the Excel
python cli.py export --input /tmp/data.json \
                     --summaries /tmp/summaries.json \
                     -o ~/Downloads/jira_report.xlsx
```

`data.json` includes `prefilled_summary` for issues that are in a blocked
status (WAIT FAE INFO / WORKED AROUND / WAIT OFFICIAL RELEASE) with no
in-period activity. `derive` skips those — `export` uses the prefilled text
and ignores any agent summary for the same key.

### `run` (legacy one-shot, uses external LLM for AI)

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

#### `run` only (external LLM AI)
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

#### `derive` only
| Option | Description |
|--------|-------------|
| `--input` | Path to `data.json` from `prepare` |
| `-o/--output` | Path for the candidate `summaries.json` |

`derive` deterministically extracts a progress line for every non-prefilled
issue (priority: verification/closure > solution/file/patch > latest comment),
reusing the same fallback logic as the GUI. It prints a list of low-confidence
issues (no resolution/solution signal) that the agent should review.

## Skill Mode vs GUI Mode

| | GUI (`jira_report_generator.py`) | Skill (agent + `cli.py prepare/derive/export`) |
|---|---|---|
| AI summary | Provider API (DeepSeek/OpenAI/Anthropic/custom), key from `.env` | **Local Claude** (Claude Code credential) + `derive` code fallback |
| Progress source | Latest comment or AI summary | AI summary or latest-comment verbatim (current user) |
| Interaction | Tkinter window | Interactive options (`AskUserQuestion`) |
| Setup | Needs a provider API key in `.env` | No LLM key needed |
| Use case | Manual weekly report | Programmatic / agent workflow |
| Output filename | `{user}_{start}_{end}_jira_report.xlsx` | Same as GUI (default `~/Downloads`) |
| Speed | Fast (no LLM inference) | Slower (Claude reads `data.json`) |

Both modes produce identical Excel output (columns, hyperlinks, dropdowns,
red highlighting for key issues).

> **Why the skill is slower than the GUI:** the skill adds a "Claude reads
> `data.json` and reviews/refines summaries" step, which is the cost of using
> local Claude for the AI summary. Jira API calls are cached per issue, and
> `prepare` truncates/ranks comments, to keep that step as small as possible.

## Pack to EXE (Optional)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed jira_report_generator.py
```

The executable will be generated in the `dist/` folder.

## Configuration (`.env`)

All settings are stored in a `.env` file at the project root. Copy `.env.example` to `.env` and fill in your credentials:

```env
# Jira credentials
JIRA_USERNAME=your-email@example.com
JIRA_PASSWORD=your-password

# AI provider: deepseek | openai | anthropic | custom
AI_PROVIDER=deepseek

# Per-provider API keys (fill in the ones you need)
DEEPSEEK_API_KEY=sk-your-deepseek-key
OPENAI_API_KEY=sk-your-openai-key
ANTHROPIC_API_KEY=sk-ant-your-anthropic-key
CUSTOM_API_KEY=

# Settings
AI_MODEL=deepseek-chat
CUSTOM_ENDPOINT=
COLUMN_ORDER=1,2,3,4,5,6,7
KEY_ISSUE_HIGHLIGHT=true
COMMENT_TIMESTAMP_PREFIX=false
THEME=Geek
LANGUAGE=zh
LAST_SAVE_DIR=~/Downloads
```

| Field | Description |
|---|---|
| `JIRA_USERNAME` | Jira login email |
| `JIRA_PASSWORD` | Jira login password |
| `AI_PROVIDER` | AI provider: `deepseek`, `openai`, `anthropic`, or `custom` |
| `DEEPSEEK_API_KEY` | DeepSeek API key |
| `OPENAI_API_KEY` | OpenAI API key |
| `ANTHROPIC_API_KEY` | Anthropic (Claude) API key |
| `CUSTOM_API_KEY` | Custom OpenAI-compatible provider API key |
| `AI_MODEL` | Model name (auto-populated when provider is selected) |
| `CUSTOM_ENDPOINT` | Endpoint URL for custom provider |
| `LAST_SAVE_DIR` | Default directory for saving reports |

## AI Summary (Optional)

Supports **DeepSeek**, **OpenAI**, **Anthropic (Claude)**, and **custom OpenAI-compatible** providers.

**Setup:**
1. Get an API key from your chosen provider
2. Set it in `.env` under the matching key (e.g. `OPENAI_API_KEY=sk-...`)
3. Or set the environment variable (e.g. `export OPENAI_API_KEY=your_key`)
4. Or use `--ai-key` + `--ai-provider` flags in CLI

**Provider model lists:**

| Provider | Models |
|---|---|
| DeepSeek | deepseek-chat, deepseek-coder, deepseek-v4-flash, deepseek-v4-pro |
| OpenAI | gpt-4o, gpt-4o-mini, gpt-4-turbo, gpt-3.5-turbo, o1, o1-mini, o3-mini |
| Anthropic | claude-sonnet-4-6, claude-opus-4-6, claude-haiku-4-5 |
| Custom | User-defined model name |

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

## Windows Scheduled Task (Weekly Automation)

Use `schedule_weekly.bat` to automate weekly report generation via Windows Task Scheduler.

### Script Usage

```batch
schedule_weekly.bat              Generate with AI summarization
schedule_weekly.bat --no-ai      Generate without AI (comment-based progress)
schedule_weekly.bat --help       Show help
```

The script automatically:
- Reads Jira credentials and AI config from `.env`
- Calculates the current ISO week (Monday ~ Sunday)
- Activates the Python virtual environment (`.venv`)
- Runs `cli.py run` and saves the report to `LAST_SAVE_DIR` (from `.env`)
- Bridges `ANTHROPIC_AUTH_TOKEN` for third-party proxy setups when the declared provider's key slot is empty

### Task Scheduler Setup

1. `Win+R` → `taskschd.msc`
2. Create Basic Task → Name: "Jira Weekly Report"
3. Trigger: **Weekly** → Monday, 9:00 AM
4. Action: **Start a program**
   - **Program:** `C:\path\to\Jira-Report\schedule_weekly.bat`
   - **Arguments:** (leave empty for AI, or `--no-ai`)
   - **Start in:** `C:\path\to\Jira-Report`
5. Conditions: uncheck "Start only if on AC power" (for laptops)

### Prerequisites

The `.venv` must already exist before the scheduled task runs:
```batch
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Quick Start (Desktop App)

1. Run `python jira_report_generator.py`
2. Enter your Jira credentials and click "Login"
3. Select a date range or use "This Week" / "This Month" quick buttons
4. (Optional) Check "Use AI Summary"
5. Choose a save location for the Excel file
6. Click "Generate Report"