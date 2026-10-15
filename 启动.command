#!/bin/bash
# ============================================================
#  非打包模式的一键启动 (首次运行会自动装依赖)
#  用法: 双击 启动.command  或 终端 bash 启动.command
# ============================================================
set -e
cd "$(dirname "$0")"

VENV=".venv_run"
if [ ! -d "$VENV" ]; then
  echo "首次运行, 创建虚拟环境并安装依赖..."
  PY=""
  command -v python3.12 &>/dev/null && PY="python3.12"
  command -v python3.11 &>/dev/null && PY="${PY:-python3.11}"
  command -v python3 &>/dev/null && PY="${PY:-python3}"
  [ -z "$PY" ] && { echo "找不到 python3, 请 brew install python@3.12"; exit 1; }
  $PY -m venv "$VENV"
  source "$VENV/bin/activate"
  pip install --upgrade pip >/dev/null
  pip install -r requirements-mac.txt
else
  source "$VENV/bin/activate"
fi

python MPS-GUI.py
echo ""
echo "App 已退出. 按回车关闭窗口..."
read _
