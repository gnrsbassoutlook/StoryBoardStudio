@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo =================================================
echo   StoryBoardStudio GitHub push script (Windows)
echo =================================================

REM NOTE: this file must stay pure ASCII (cmd.exe mis-parses UTF-8 .bat files).

REM SSH is used by default; this machine needs a working GitHub SSH key.
REM To use HTTPS instead, replace the line below with:
REM   set "REMOTE_SSH=https://github.com/gnrsbassoutlook/StoryBoardStudio.git"
set "REMOTE_SSH=git@github.com:gnrsbassoutlook/StoryBoardStudio.git"

if not exist ".git" (
    echo [*] Initializing git repository ...
    git init
    git branch -M main
    git remote add origin "%REMOTE_SSH%"
)

set "REMOTE_URL="
for /f "delims=" %%i in ('git remote get-url origin 2^>nul') do set "REMOTE_URL=%%i"
if "%REMOTE_URL%"=="" (
    git remote add origin "%REMOTE_SSH%"
) else (
    if not "%REMOTE_URL%"=="%REMOTE_SSH%" (
        echo [*] Updating origin to: %REMOTE_SSH%
        git remote set-url origin "%REMOTE_SSH%"
    )
)

echo.
echo [*] Working tree status:
git status -s
echo.

set "STATUS_OUT="
for /f "delims=" %%i in ('git status -s') do set "STATUS_OUT=1"
if not defined STATUS_OUT (
    echo [i] Working tree is clean - nothing to commit, nothing to push.
    echo.
    pause
    exit /b 0
)

set "MSG="
set /p "MSG=Commit message [Enter = Auto update]: "
if "%MSG%"=="" set "MSG=Auto update"

echo.
echo [*] Committing and pushing to GitHub ...
git add .
git commit -m "%MSG%" --allow-empty
git push -u origin main
if errorlevel 1 (
    echo.
    echo [!] Push failed; trying to pull --rebase and push again ...
    git pull origin main --rebase
    if errorlevel 1 (
        echo [X] Rebase failed. Please resolve the conflicts manually.
    ) else (
        git push -u origin main
        echo [*] Done. Pushed to GitHub.
    )
) else (
    echo.
    echo [*] Done. Pushed to GitHub.
)

echo.
pause
