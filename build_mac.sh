#!/bin/bash
# ============================================================
#  MPS-Capacity.app  一键构建脚本  (macOS Apple Silicon)
#  用法: 双击执行  或  终端  bash build_mac.sh
# ============================================================
set -e
cd "$(dirname "$0")"

echo "=== [1/5] 检查 Python 版本 (需要 >= 3.11) ==="
PY_BIN=""
if command -v python3.12 &>/dev/null; then PY_BIN="python3.12"
elif command -v python3.11 &>/dev/null; then PY_BIN="python3.11"
elif command -v python3 &>/dev/null; then PY_BIN="python3"
else
  echo "[ERROR] 找不到 python3, 请先安装 Homebrew Python:  brew install python@3.12"
  exit 1
fi
echo "使用解释器: $($PY_BIN --version)"

VENV=".venv_build"
echo ""
echo "=== [2/5] 创建虚拟环境 $VENV ==="
rm -rf "$VENV"
$PY_BIN -m venv "$VENV"
source "$VENV/bin/activate"
pip install --upgrade pip >/dev/null

echo ""
echo "=== [3/5] 安装依赖 ==="
pip install -r requirements-mac.txt
pip install pyinstaller

echo ""
echo "=== [4/5] PyInstaller 打包 ==="
rm -rf build dist
pyinstaller MPS-Capacity.app.spec --noconfirm

echo ""
echo "=== [5/5] Ad-hoc 签名 (绕过 Gatekeeper) ==="
if [ -d "dist/MPS-Capacity.app" ]; then
  codesign --force --sign - dist/MPS-Capacity.app || true
  echo ""
  echo "构建完成!"
  echo "  App 位置:  dist/MPS-Capacity.app"
  echo ""
  echo "首次运行前建议执行 (清掉系统隔离属性):"
  echo "  xattr -cr dist/MPS-Capacity.app"
  echo ""
  echo "直接打开:"
  echo "  open dist/MPS-Capacity.app"
else
  echo "[ERROR] dist/MPS-Capacity.app 不存在, 构建可能失败, 检查上方日志"
  exit 1
fi
