@echo off
REM ============================================================
REM   bjxh-ppt skill installer for Windows
REM
REM   Copies this folder into %USERPROFILE%\.claude\skills\bjxh-ppt
REM   so the AI agent can discover and use the skill.
REM
REM   Usage:  double-click install.bat
REM   Run from a regular user account (no admin needed).
REM ============================================================

setlocal EnableDelayedExpansion

set "SKILL_NAME=bjxh-ppt"
set "DEST_DIR=%USERPROFILE%\.claude\skills\%SKILL_NAME%"
set "SRC_DIR=%~dp0"

echo.
echo === bjxh-ppt skill installer ===
echo.
echo Source : %SRC_DIR%
echo Target : %DEST_DIR%
echo.

REM Sanity check: must contain SKILL.md
if not exist "%SRC_DIR%SKILL.md" (
    echo [ERROR] SKILL.md not found in %SRC_DIR%
    echo         Please run install.bat from inside the skill folder.
    pause
    exit /b 1
)

REM Check Python
where python >nul 2>&1
if errorlevel 1 (
    echo [WARN] Python is not on PATH.
    echo        Install from https://www.python.org/downloads/ ^(3.11+ recommended^)
    echo        Make sure to check "Add Python to PATH" during installation.
    echo.
)

REM Check python-pptx
python -c "import pptx, lxml" >nul 2>&1
if errorlevel 1 (
    echo [WARN] python-pptx or lxml not installed.
    echo        Run: pip install python-pptx lxml
    echo.
)

REM Create destination
if not exist "%USERPROFILE%\.claude\skills" (
    mkdir "%USERPROFILE%\.claude\skills"
)

REM If old version exists, back it up
if exist "%DEST_DIR%" (
    set "BACKUP=%DEST_DIR%.backup.%RANDOM%"
    echo [INFO] Existing installation found, moving to !BACKUP!
    move "%DEST_DIR%" "!BACKUP!" >nul
)

REM Copy
xcopy /E /I /Y /Q "%SRC_DIR%*" "%DEST_DIR%\" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed.
    pause
    exit /b 1
)

echo [OK] Installed to: %DEST_DIR%
echo.
echo Next steps:
echo   1. Make sure Python 3.11+ is installed and on PATH
echo   2. Run: pip install python-pptx lxml
echo   3. Configure your template path — copy scripts\config.example.json
echo      to scripts\config.json and set "template_path", or set the
echo      BJXH_TEMPLATE environment variable.
echo   4. Restart Claude Code / OpenCode to pick up the new skill
echo.
echo To uninstall: run uninstall.bat
echo.
pause
