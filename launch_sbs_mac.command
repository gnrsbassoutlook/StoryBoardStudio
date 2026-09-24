#!/bin/bash
# 切换到脚本所在绝对目录
cd "$(dirname "$0")" || exit 1

PORT=7861

echo "================================================="
echo "  🎬 正在启动 StoryBoardStudio (Mac)   "
echo "================================================="

# ---- 启动前自检：端口已被占用 = 程序已经在跑了，不重复启动 ----
# 重复启动会让新进程绑定失败并抛 "errno 48 address already in use"，看起来很吓人，
# 其实只是「已经开着呢」。这里直接改成为你打开页面。
if command -v lsof >/dev/null 2>&1 && lsof -nP -iTCP:$PORT -sTCP:LISTEN >/dev/null 2>&1; then
    echo ""
    echo "ℹ️  检测到 StoryBoardStudio 已经在运行（端口 $PORT 被占用）。"
    echo "    不需要重复启动，已为你打开： http://localhost:$PORT"
    echo ""
    echo "    · 想重启：先在原来那个运行窗口按 Ctrl+C，"
    echo "      或在本终端执行： pkill -f StoryBoardStudio.py"
    echo "      然后再双击本脚本。"
    echo ""
    open "http://localhost:$PORT"
    read -p "按回车键关闭本窗口..."
    exit 0
fi

# 检测 Python3
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未检测到 Python3，请先安装 Python 3.10+！"
    read -p "按回车键退出..."
    exit 1
fi

# 检查虚拟环境
# 【重要】venv 不能跨机器 / 跨路径拷贝：shebang 里写死了绝对路径，.so 又是 Mach-O 二进制。
# 检测到 venv 目录存在但不可用（多为跨机器拷来），改名留档后重建，不硬删。
if [ -d "venv" ] && [ ! -x "venv/bin/python" ]; then
    mv venv "venv_broken_$(date +%Y%m%d_%H%M%S)"
    echo "⚠️  原 venv 不可用（多为跨机器拷贝所致），已改名留档，正在重建..."
fi
if [ ! -x "venv/bin/python" ]; then
    echo "⚙️ 首次运行，正在创建专属虚拟环境 venv..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 自动检查并安装依赖（走官方 PyPI 源 https://pypi.org/simple）
echo "📦 正在检查环境依赖..."
pip install -q --upgrade pip
pip install -q -r requirements.txt

echo "🚀 服务已就绪，正在启动 Gradio 并打开浏览器..."
echo "💡 提示：同局域网设备可通过本机 IP:$PORT 访问协作！"

# 在后台延迟 2 秒自动打开默认浏览器
(sleep 2 && open "http://localhost:$PORT") &

# 启动 Python 服务
python3 StoryBoardStudio.py

# 退出提示
echo ""
echo "服务已停止。"
read -p "按回车键关闭窗口..."
