@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo =================================================
echo   StoryBoardStudio GitHub 自动推送脚本 (Windows)
echo =================================================

REM 优先使用 SSH，需要本机已配置 GitHub SSH 密钥。
REM 不想用 SSH 就换成 HTTPS： set "REMOTE_SSH=https://github.com/gnrsbassoutlook/StoryBoardStudio.git"
set "REMOTE_SSH=git@github.com:gnrsbassoutlook/StoryBoardStudio.git"

if not exist ".git" (
    echo [*] 正在初始化 Git 仓库...
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
        echo [*] 更新 origin 远程地址为: %REMOTE_SSH%
        git remote set-url origin "%REMOTE_SSH%"
    )
)

echo.
echo [*] 当前文件改动状态：
git status -s
echo.

set "STATUS_OUT="
for /f "delims=" %%i in ('git status -s') do set "STATUS_OUT=1"
if not defined STATUS_OUT (
    echo [i] 检测到 working tree clean，没有文件改动，无需提交推送。
    echo.
    pause
    exit /b 0
)

set "MSG="
set /p "MSG=请输入 Commit 说明 (直接回车默认: Auto update): "
if "%MSG%"=="" set "MSG=Auto update"

echo.
echo [*] 正在提交并推送到 GitHub ...
git add .
git commit -m "%MSG%" --allow-empty
git push -u origin main
if errorlevel 1 (
    echo.
    echo [!] 推送可能遇到冲突，正在尝试拉取合并后重新推送...
    git pull origin main --rebase
    if errorlevel 1 (
        echo [X] rebase 合并失败，请手动解决冲突！
    ) else (
        git push -u origin main
        echo [*] 推送成功！代码已同步至 GitHub。
    )
) else (
    echo.
    echo [*] 推送成功！代码已同步至 GitHub。
)

echo.
pause
