@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ============================================================
::  Jira Weekly Report — Scheduled Task Entry Point
::
::  Usage:
::    schedule_weekly.bat              Generate with AI summarization
::    schedule_weekly.bat --no-ai      Generate without AI (comment-based)
::    schedule_weekly.bat --help       Show help
::
::  Windows Task Scheduler setup:
::    1. Win+R → taskschd.msc
::    2. Create Basic Task → Name: "Jira Weekly Report"
::    3. Trigger: Weekly → Monday, 9:00 AM
::    4. Action: Start a program
::       Program:   C:\path\to\Jira-Report\schedule_weekly.bat
::       Arguments: (leave empty for AI, or --no-ai)
::       Start in:  C:\path\to\Jira-Report
::    5. Conditions: uncheck "Start only if on AC power" (for laptops)
:: ============================================================

:: --- Switch to the directory containing this bat file ---
cd /d "%~dp0"

:: ============================================================
::  Parse command-line arguments
:: ============================================================
set "USE_AI=1"
for %%a in (%*) do (
    if /i "%%a"=="--no-ai"   set "USE_AI=0"
    if /i "%%a"=="-h"        goto :show_help
    if /i "%%a"=="--help"    goto :show_help
)

:: ============================================================
::  Read settings from .env
:: ============================================================
for /f "tokens=1,* delims==" %%a in ('findstr /R "^JIRA_USERNAME="        .env 2^>nul') do set "JIRA_USERNAME=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^JIRA_PASSWORD="        .env 2^>nul') do set "JIRA_PASSWORD=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^LAST_SAVE_DIR="        .env 2^>nul') do set "OUTDIR=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^AI_PROVIDER="          .env 2^>nul') do set "AI_PROVIDER=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^DEEPSEEK_API_KEY="     .env 2^>nul') do set "DS_KEY=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^ANTHROPIC_AUTH_TOKEN=" .env 2^>nul') do set "AUTH_TOKEN=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^ANTHROPIC_API_KEY="    .env 2^>nul') do set "ANTHROPIC_KEY=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^OPENAI_API_KEY="       .env 2^>nul') do set "OPENAI_KEY=%%b"
for /f "tokens=1,* delims==" %%a in ('findstr /R "^CUSTOM_API_KEY="       .env 2^>nul') do set "CUSTOM_KEY=%%b"

:: --- Validate Jira credentials ---
if "%JIRA_USERNAME%"=="" (
    echo [ERROR] Could not read JIRA_USERNAME from .env
    pause
    exit /b 1
)

:: --- Fallback output directory ---
if "%OUTDIR%"=="" set "OUTDIR=%~dp0"

:: ============================================================
::  Calculate current week (ISO week: Monday ~ Sunday)
:: ============================================================
for /f %%i in ('powershell -NoProfile -Command "(Get-Date).AddDays(-((Get-Date).DayOfWeek.value__+6)%%7).ToString('yyyy-MM-dd')"') do set "MON=%%i"
for /f %%j in ('powershell -NoProfile -Command "(Get-Date).AddDays(6-((Get-Date).DayOfWeek.value__+6)%%7).ToString('yyyy-MM-dd')"') do set "SUN=%%j"

set "OUTFILE=%OUTDIR%\jira_weekly_%MON%_%SUN%.xlsx"

:: ============================================================
::  Provider / API key compatibility bridge
::
::  Many third-party proxy setups store the real API token in
::  ANTHROPIC_AUTH_TOKEN (for Claude Code compatibility) while
::  AI_PROVIDER is set to "deepseek".  This section detects the
::  mismatch and patches the missing slot so the CLI can find a key.
:: ============================================================
if "%USE_AI%"=="1" (

    :: --- Determine which provider's key slot is empty ---
    set "NEED_PATCH=0"

    if /i "%AI_PROVIDER%"=="deepseek" (
        if "%DS_KEY%"=="" set "NEED_PATCH=1"
    )
    if /i "%AI_PROVIDER%"=="anthropic" (
        if "%ANTHROPIC_KEY%"=="" set "NEED_PATCH=1"
    )
    if /i "%AI_PROVIDER%"=="openai" (
        if "%OPENAI_KEY%"=="" set "NEED_PATCH=1"
    )
    if /i "%AI_PROVIDER%"=="custom" (
        if "%CUSTOM_KEY%"=="" set "NEED_PATCH=1"
    )

    :: --- Patch: use ANTHROPIC_AUTH_TOKEN as fallback for the declared provider ---
    if "!NEED_PATCH!"=="1" (
        if not "!AUTH_TOKEN!"=="" (
            echo [INFO] %AI_PROVIDER% key slot is empty, bridging from ANTHROPIC_AUTH_TOKEN.
            if /i "!AI_PROVIDER!"=="deepseek"  set "DEEPSEEK_API_KEY=!AUTH_TOKEN!"
            if /i "!AI_PROVIDER!"=="anthropic" set "ANTHROPIC_API_KEY=!AUTH_TOKEN!"
            if /i "!AI_PROVIDER!"=="openai"    set "OPENAI_API_KEY=!AUTH_TOKEN!"
            if /i "!AI_PROVIDER!"=="custom"    set "CUSTOM_API_KEY=!AUTH_TOKEN!"
        ) else (
            echo [WARN] AI is enabled but no API key found for provider '%AI_PROVIDER%'.
            echo        AI summarization will be skipped by the CLI.
        )
    )
)

:: ============================================================
::  Display summary
:: ============================================================
set "AI_LABEL=ON"
if "%USE_AI%"=="0" set "AI_LABEL=OFF"
echo ============================================
echo   Jira Weekly Report Generator
echo   Week   : %MON% ~ %SUN%
echo   AI     : %AI_LABEL%
echo   Output : %OUTFILE%
echo ============================================
echo.

:: ============================================================
::  Activate Python virtual environment
:: ============================================================
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment not found: .venv\Scripts\activate.bat
    echo          Run: python -m venv .venv
    echo          Then: .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

:: ============================================================
::  Build AI flags
:: ============================================================
set "AI_FLAGS="
if "%USE_AI%"=="1" set "AI_FLAGS=--ai --fetch-comment"

:: ============================================================
::  Generate the report
:: ============================================================
python cli.py run ^
    -u "%JIRA_USERNAME%" ^
    -p "%JIRA_PASSWORD%" ^
    --start %MON% ^
    --end %SUN% ^
    -o "%OUTFILE%" ^
    %AI_FLAGS% 2>&1

:: ============================================================
::  Result
:: ============================================================
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Report generation failed (exit code: %errorlevel%).
    pause
    exit /b %errorlevel%
)

echo.
echo [OK] Report saved: %OUTFILE%
exit /b 0

:: ============================================================
::  Help
:: ============================================================
:show_help
echo Usage: schedule_weekly.bat [--no-ai] [--help]
echo.
echo   (no args)   Generate weekly report with AI progress summarization.
echo   --no-ai     Skip AI — use the latest Jira comment as progress instead.
echo   --help      Show this help.
echo.
echo The script reads Jira credentials and AI config from .env automatically.
echo It detects common third-party proxy setups (ANTHROPIC_AUTH_TOKEN bridging)
echo and patches the credential mismatch so the CLI finds a usable API key.
echo.
echo Output file: %LAST_SAVE_DIR%\jira_weekly_{start}_{end}.xlsx
echo              (LAST_SAVE_DIR from .env, or the project directory as fallback)
exit /b 0
