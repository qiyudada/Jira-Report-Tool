# Jira Report

A desktop application for generating Excel reports from Jira issues.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python jira_report_generator.py
```

## Pack to EXE (Optional)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed jira_report_generator.py
```

The executable will be generated in the `dist/` folder.

## Configuration (`.jira_config`)

The app stores settings in `.jira_config` (JSON) in the same directory as the executable. You can edit it manually:

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
| `deepseek_api_key` | DeepSeek API Key for AI summary feature (see below) |
| `ai_model` | DeepSeek model to use (see below) |

## AI Summary (Optional)

The app supports AI-powered comment summarization via the [DeepSeek API](https://platform.deepseek.com/).

**Setup:**

1. Register at [platform.deepseek.com](https://platform.deepseek.com/) and create an API Key.
2. Fill in `deepseek_api_key` in `.jira_config`, or the app will save it automatically after you use the feature.

**Supported models:**

| Model | Description |
|-------|-------------|
| `deepseek-chat` | General-purpose chat model (default, recommended) |
| `deepseek-coder` | Optimized for code-related content |
| `deepseek-v4-flash` | Faster, lower-cost variant |
| `deepseek-v4-pro` | Higher capability, higher cost |

**Usage:**

1. Check **[x] Use AI Summary** in the app before generating the report.
2. Select a model from the **Model** dropdown.
3. Generate the report — each issue's latest progress column will be filled with an AI-generated summary.

> If `deepseek_api_key` is empty, the AI summary column will show `[AI总结] 未配置DeepSeek API Key`.

## Quick Start

1. Run the application
2. Enter your Jira credentials and click "Login"
3. Select a date range or use "This Week" / "This Month" quick buttons
4. (Optional) Check "Use AI Summary" and configure your DeepSeek API Key in `.jira_config`
5. Choose a save location for the Excel file
6. Click "Generate Report"
