@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PORT=7861"

echo =================================================
echo   StoryBoardStudio (Windows)
echo =================================================

REM NOTE: this file must stay pure ASCII.
REM cmd.exe mis-parses UTF-8 batch files (byte/char offset bug), which turns
REM Chinese lines into "'xxx' is not recognized as an internal or external command".
REM All Chinese messages live in the Python side, which prints them correctly.

REM ---- Startup check: port already taken means the app is already running ----
REM Starting a second instance fails with "errno 10048 address already in use",
REM which looks scary but only means "it is already up".
netstat -ano | findstr ":%PORT%" | findstr "LISTENING" >nul
if not errorlevel 1 (
    echo.
    echo [i] StoryBoardStudio is already running on port %PORT%.
    echo     No need to start it again. Opening http://localhost:%PORT%
    echo.
    echo     To restart: press Ctrl+C in the window that is running it,
    echo     or run: taskkill /F /IM python.exe
    echo     then double-click this script again.
    echo.
    start "" "http://localhost:%PORT%"
    pause
    exit /b 0
)

REM ---- Detect Python ----
where python >nul 2>nul
if errorlevel 1 (
    echo [X] Python not found. Please install Python 3.10+ first.
    echo     Download: https://www.python.org/downloads/windows/
    echo     During setup, tick "Add Python to PATH".
    pause
    exit /b 1
)

REM ---- Virtual environment ----
REM A venv must never be copied across machines or OSes: its scripts contain
REM absolute paths and its binaries are platform specific. This repo ships no
REM venv; it is created on first run. A venv folder that is not a Windows venv
REM is renamed aside (kept for reference), never hard-deleted.
if exist "venv" if not exist "venv\Scripts\python.exe" (
    echo [!] Found a "venv" folder that is not a Windows venv ^(probably copied from a Mac^).
    echo     Renaming it aside and rebuilding ...
    move "venv" "venv_broken_%RANDOM%" >nul
)
if not exist "venv\Scripts\python.exe" (
    echo [*] Creating the virtual environment "venv" ...
    python -m venv venv
)

call "venv\Scripts\activate.bat"

REM Official PyPI source: https://pypi.org/simple
echo [*] Checking environment dependencies ...
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt

echo.
echo [*] Ready. Starting Gradio and opening the browser ...
echo     Stop the service with Ctrl+C in this window.
echo.

REM Wait 3 seconds before opening the browser so the server can boot
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:%PORT%"

python StoryBoardStudio.py

echo.
echo Service stopped.
pause
