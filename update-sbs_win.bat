@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo =================================================
echo   StoryBoardStudio GitHub 更新脚本 (Windows)
echo =================================================

if not exist ".git" (
    echo [X] 当前目录尚未关联 Git 仓库，请先 git clone。
    pause
    exit /b 1
)

echo.
echo ==================== 本地最近20个提交记录 ====================
git log --oneline -n 20
echo.
echo ==============================================================
echo.
for /f "delims=" %%i in ('git rev-parse --short HEAD') do echo 当前本地 commit: %%i
echo.

pause
echo [*] 正在从远程仓库拉取最新代码...
git pull origin main
if errorlevel 1 (
    echo.
    echo [X] 更新失败，请检查网络或是否存在本地文件冲突！
    pause
    exit /b 1
)

echo.
echo [*] 更新成功！代码已是最新版本。
for /f "delims=" %%i in ('git rev-parse --short HEAD') do echo [*] 更新后 commit: %%i

if exist "requirements.txt" (
    echo [*] 正在检查并更新依赖包...
    if exist "venv\Scripts\python.exe" (
        call "venv\Scripts\activate.bat"
        python -m pip install -q -r requirements.txt
    ) else (
        echo [!] 未发现 venv，跳过依赖更新（双击 launch_sbs_win.bat 会自动创建）。
    )
)

echo.
pause
