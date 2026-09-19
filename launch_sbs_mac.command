#!/bin/bash
# 切换到脚本所在绝对目录
cd "$(dirname "$0")"

echo "================================================="
echo "  🎬 正在启动 StoryBoardStudio (Mac)   "
echo "================================================="

# 检测 Python3
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未检测到 Python3，请先安装 Python 3.10+！"
    read -p "按回车键退出..."
    exit 1
fi

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "⚙️ 首次运行，正在创建专属虚拟环境 venv..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 自动检查并安装依赖
echo "📦 正在检查环境依赖..."
pip install -q --upgrade pip
pip install -q -r requirements.txt

echo "🚀 服务已就绪，正在启动 Gradio 并打开浏览器..."
echo "💡 提示：同局域网设备可通过本机 IP:7861 访问协作！"

# 在后台延迟 2 秒自动打开默认浏览器
(sleep 2 && open "http://localhost:7861") &

# 启动 Python 服务
python3 app.py

# 退出提示
read -p "服务已停止，按回车键关闭窗口..."