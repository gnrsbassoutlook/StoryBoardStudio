@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul 2>&1
cd /d "%~dp0"
title StoryBoardStudio - Push to GitHub

rem ===================== config =====================
set "REPO_SSH=git@github.com:gnrsbassoutlook/StoryBoardStudio.git"
set "REPO_HTTPS=https://github.com/gnrsbassoutlook/StoryBoardStudio.git"
set "BRANCH=main"
rem ===================================================

rem NOTE: this file must stay pure ASCII (cmd.exe mis-parses UTF-8 .bat files).
rem Keep Windows system dirs first in PATH so find/findstr are not hijacked by Git Bash.
set "PATH=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%PATH%"

echo ==========================================
echo   StoryBoardStudio  -  Push to GitHub
echo ==========================================

git --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [ERROR] git not found. Please install Git for Windows.
    echo.
    pause
    exit /b 1
)

rem ---- current branch ----
for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set "CURB=%%b"
if defined CURB if not "!CURB!"=="HEAD" set "BRANCH=!CURB!"

if not exist ".git" (
    echo.
    echo   [i] No git repository found, initializing ...
    git init
    git branch -M !BRANCH!
)

echo.
echo   [1/4] Detecting remote channel ...
git ls-remote "!REPO_SSH!" HEAD >nul 2>&1
if not errorlevel 1 (
    set "REMOTE=!REPO_SSH!"
    echo         SSH is available
) else (
    set "REMOTE=!REPO_HTTPS!"
    echo         SSH unavailable, falling back to HTTPS
)

git remote get-url origin >nul 2>&1
if errorlevel 1 (
    git remote add origin "!REMOTE!"
) else (
    for /f "delims=" %%u in ('git remote get-url origin') do set "CURU=%%u"
    if /i not "!CURU!"=="!REMOTE!" (
        echo         origin URL updated
        git remote set-url origin "!REMOTE!"
    )
)

echo.
echo   [2/4] Fetching remote state ...
git fetch origin !BRANCH! 2>nul
set "BEHIND=0"
for /f "delims=" %%c in ('git rev-list --count HEAD..origin/!BRANCH! 2^>nul') do set "BEHIND=%%c"
if !BEHIND! GTR 0 (
    echo         Local branch is !BEHIND! commit^(s^) behind, rebasing first ...
    git pull --rebase --autostash origin !BRANCH!
    if errorlevel 1 (
        echo.
        echo   [X] Rebase failed. Please resolve conflicts manually and retry.
        pause
        exit /b 1
    )
)

echo.
echo   [3/4] Current changes:
echo   --------------------------------------------------------------
git status -s
echo   --------------------------------------------------------------

set "HASCHG="
for /f "delims=" %%i in ('git status --porcelain') do set "HASCHG=1"
if not defined HASCHG (
    echo.
    echo   [i] Working tree is clean - nothing to commit, nothing to push.
    echo.
    pause
    exit /b 0
)

echo.
set "MSG="
set /p "MSG=  Commit message (Enter = auto timestamp): "
if not defined MSG set "MSG=Auto update: !date! !time!"

echo.
echo   [4/4] Committing and pushing to !BRANCH! ...
git add -A
git commit -m "!MSG!"
if errorlevel 1 (
    echo.
    echo   [X] Commit failed, stopped.
    pause
    exit /b 1
)

git push -u origin !BRANCH!
if errorlevel 1 (
    echo.
    echo   [!] Push rejected, pulling and retrying ...
    git pull --rebase --autostash origin !BRANCH!
    if errorlevel 1 (
        echo   [X] Merge failed. Please resolve conflicts manually.
        pause
        exit /b 1
    )
    git push -u origin !BRANCH!
    if errorlevel 1 (
        echo   [X] Push still failing. Check network or repository permission.
        pause
        exit /b 1
    )
)

echo.
echo   Pushed successfully.
echo.
pause
exit /b 0
