@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo =================================================
echo   StoryBoardStudio updater (Windows)
echo =================================================
echo.

REM NOTE: this file must stay pure ASCII (cmd.exe mis-parses UTF-8 .bat files).
REM All interactive logic lives in sbs_update.py so that macOS and Windows behave
REM exactly the same, and so Chinese messages never get mangled by cmd.exe.

where python >nul 2>nul
if errorlevel 1 (
    echo [X] Python not found. Please install Python 3.10+ first.
    echo     Download: https://www.python.org/downloads/windows/
    echo     During setup, tick "Add Python to PATH".
    pause
    exit /b 1
)

python sbs_update.py
if errorlevel 1 echo [!] The updater stopped with an error - see the messages above.

echo.
pause
