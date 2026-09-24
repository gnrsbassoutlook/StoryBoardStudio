@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PORT=7861"

echo =================================================
echo   StoryBoardStudio (Windows)
echo =================================================

REM ---- 启动前自检：端口被占用 = 程序已经在跑，不重复启动 ----
REM 重复启动会绑定失败并抛 "errno 48 / 10048 address already in use"，看着吓人其实只是「已经开着呢」。
netstat -ano | findstr ":%PORT%" | findstr "LISTENING" >nul
if not errorlevel 1 (
    echo.
    echo [i] 检测到 StoryBoardStudio 已经在运行（端口 %PORT% 被占用）。
    echo     不需要重复启动，已为你打开： http://localhost:%PORT%
    echo.
    echo     · 想重启：先在原来那个运行窗口按 Ctrl+C，
    echo       或在本终端执行： taskkill /F /IM python.exe
    echo       然后再双击本脚本。
    echo.
    start "" "http://localhost:%PORT%"
    pause
    exit /b 0
)

REM ---- 检测 Python ----
where python >nul 2>nul
if errorlevel 1 (
    echo [X] 未检测到 Python，请先安装 Python 3.10+ ^(安装时务必勾选 "Add Python to PATH"^)！
    echo     下载地址： https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

REM ---- 虚拟环境 ----
REM 【重要】venv 不能跨机器 / 跨系统拷贝：Mac 上生成的 venv 在 Windows 上根本用不了
REM （.so 是 Mach-O，Windows 需要 .dll；目录结构也是 Scripts\ 而非 bin/）。
REM 本仓库不含 venv，首次运行会自动创建；若目录里有从别处拷来的 venv，请先整个删掉。
if exist "venv" if not exist "venv\Scripts\python.exe" (
    echo [!] 发现 venv 目录但不是 Windows 虚拟环境（多半是从 Mac 拷来的），改名留档后重建 ...
    move "venv" "venv_broken_%RANDOM%" >nul
)
if not exist "venv\Scripts\python.exe" (
    echo [*] 正在创建专属虚拟环境 venv ...
    python -m venv venv
)

call "venv\Scripts\activate.bat"

REM 走官方 PyPI 源 https://pypi.org/simple
echo [*] 正在检查环境依赖...
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt

echo.
echo [*] 服务已就绪，正在启动 Gradio 并打开浏览器 ...
echo     提示：同局域网设备可通过本机 IP:%PORT% 访问协作
echo     停止服务：在本窗口按 Ctrl+C
echo.

REM 延迟 3 秒再开浏览器，给服务留出启动时间
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:%PORT%"

python StoryBoardStudio.py

echo.
echo 服务已停止。
pause
