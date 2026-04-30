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

## Quick Start

1. Run the application
2. Enter your Jira credentials and click "Login"
3. Select a date range or use "This Week" / "This Month" quick buttons
4. Choose a save location for the Excel file
5. Click "Generate Report"
