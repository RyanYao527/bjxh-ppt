@echo off
REM ============================================================
REM   bjxh-ppt skill uninstaller
REM
REM   Removes %USERPROFILE%\.claude\skills\bjxh-ppt
REM   Backs up the current install to bjxh-ppt.backup.<rand>
REM   just in case you want to restore it.
REM ============================================================

setlocal EnableDelayedExpansion

set "SKILL_NAME=bjxh-ppt"
set "DEST_DIR=%USERPROFILE%\.claude\skills\%SKILL_NAME%"

echo.
echo === bjxh-ppt skill uninstaller ===
echo.

if not exist "%DEST_DIR%" (
    echo [INFO] %SKILL_NAME% is not installed at:
    echo        %DEST_DIR%
    echo        Nothing to do.
    pause
    exit /b 0
)

set "BACKUP=%DEST_DIR%.backup.%RANDOM%"
echo [INFO] Backing up to: !BACKUP!
move "%DEST_DIR%" "!BACKUP!" >nul
if errorlevel 1 (
    echo [ERROR] Move failed.
    pause
    exit /b 1
)

echo [OK] Uninstalled. Backup kept at: !BACKUP!
echo.
echo Restart Claude Code / OpenCode to unload the skill.
echo.
pause
