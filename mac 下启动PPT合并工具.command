#!/bin/bash

# PPT转PDF合并工具启动脚本（Mac版）
# 使用 conda base 环境启动
# 双击此文件即可在终端中启动程序

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 获取 Python 脚本路径
PYTHON_SCRIPT="$(pwd)/ppt_pdf_merger.py"

# 检查 Python 脚本是否存在
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "❌ 错误：未找到 Python 脚本：$PYTHON_SCRIPT"
    echo ""
    echo "按任意键退出..."
    read -n 1
    exit 1
fi

echo "🚀 正在启动 PPT 转 PDF 合并工具..."
echo "📁 工作目录：$(pwd)"
echo ""

# 查找 conda 命令
CONDA_CMD=""
if command -v conda &> /dev/null; then
    CONDA_CMD="conda"
elif [ -f "$HOME/anaconda3/bin/conda" ]; then
    CONDA_CMD="$HOME/anaconda3/bin/conda"
elif [ -f "$HOME/miniconda3/bin/conda" ]; then
    CONDA_CMD="$HOME/miniconda3/bin/conda"
elif [ -f "/opt/homebrew/Caskroom/miniconda/base/bin/conda" ]; then
    CONDA_CMD="/opt/homebrew/Caskroom/miniconda/base/bin/conda"
fi

# 如果找到 conda，使用 conda run 运行
if [ -n "$CONDA_CMD" ]; then
    echo "✅ 使用 conda base 环境启动..."
    echo ""
    "$CONDA_CMD" run -n base python "$PYTHON_SCRIPT"
    EXIT_CODE=$?
else
    # 如果找不到 conda，尝试初始化 conda 环境
    if [ -f "$HOME/anaconda3/etc/profile.d/conda.sh" ]; then
        source "$HOME/anaconda3/etc/profile.d/conda.sh"
        conda activate base
        python "$PYTHON_SCRIPT"
        EXIT_CODE=$?
    elif [ -f "$HOME/miniconda3/etc/profile.d/conda.sh" ]; then
        source "$HOME/miniconda3/etc/profile.d/conda.sh"
        conda activate base
        python "$PYTHON_SCRIPT"
        EXIT_CODE=$?
    else
        echo "❌ 错误：未找到 conda 环境"
        echo "请确保已安装 conda/anaconda/miniconda"
        echo "或者手动运行：conda activate base && python $PYTHON_SCRIPT"
        echo ""
        echo "按任意键退出..."
        read -n 1
        exit 1
    fi
fi

# 程序运行完成后的处理
echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo "✅ 程序已正常退出"
else
    echo "⚠️  程序退出，退出码：$EXIT_CODE"
fi
echo ""
echo "按任意键关闭窗口..."
read -n 1

