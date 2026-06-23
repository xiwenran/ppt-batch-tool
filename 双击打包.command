#!/bin/bash

set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR"

clear
echo "=========================================="
echo "  PPT / Word 转图片工具 - 双击打包"
echo "=========================================="
echo ""

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3。请先安装 Python 3。"
  echo ""
  read -r -p "按回车关闭窗口..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "首次运行：正在创建本项目专用 Python 环境..."
  python3 -m venv .venv
fi

source ".venv/bin/activate"

echo "正在安装 / 更新打包依赖..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo ""
echo "开始打包..."
bash build_app.sh

echo ""
echo "双击打包流程完成。"
echo "输出位置：$PROJECT_DIR/dist"
echo ""
read -r -p "按回车关闭窗口..."
