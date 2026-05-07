#!/usr/bin/env bash
# 在本机生成可分发目录：macOS 为 dist/智能排款.app；Windows 请在 CMD 中运行 build_release.bat。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi
.venv/bin/python -m pip install -q -r requirements-dev.txt

.venv/bin/python -m PyInstaller --noconfirm 智能排款.spec

# macOS：COLLECT 会多留一份与 .app 等大的「智能排款」文件夹，删除以免误打包两份。
if [[ "$(uname -s)" == "Darwin" ]] && [[ -d "dist/智能排款" ]] && [[ -d "dist/智能排款.app" ]]; then
  rm -rf "dist/智能排款"
fi

rm -rf "dist/AutoPaymentScheduleFile"
cp -R "AutoPaymentScheduleFile" "dist/AutoPaymentScheduleFile"

echo "完成。请将 dist 目录下的应用与 AutoPaymentScheduleFile 一起拷贝给对方（须在同一层目录）。详见 打包与分发说明.txt。"
