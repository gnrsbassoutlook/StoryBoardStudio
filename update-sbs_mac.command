#!/bin/bash
# StoryBoardStudio macOS 更新脚本
# 交互式更新：列出远程最近 20 个提交，输入序号（回车 = 1 = 最新版）选择要部署到本地的版本。
# 具体逻辑在 sbs_update.py 里，与 Windows 版共用同一份实现。
cd "$(dirname "$0")" || exit 1

# 优先 python3，没有就退回 python
PY=""
if command -v python3 >/dev/null 2>&1; then
    PY="python3"
elif command -v python >/dev/null 2>&1; then
    PY="python"
else
    echo "❌ 未检测到 python3 / python，无法执行更新，请先安装 Python 3.10+。"
    read -n 1 -s -r -p "按任意键退出..."
    echo ""
    exit 1
fi

"$PY" sbs_update.py
RET=$?

echo ""
if [ $RET -ne 0 ]; then
    echo "ℹ️  本次没有更新（返回码 $RET）—— 主动取消或出错都会这样，上面的提示是原因。"
fi
read -n 1 -s -r -p "按任意键关闭窗口..."
echo ""
exit $RET
