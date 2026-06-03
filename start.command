#!/bin/zsh
set -e

cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "没有找到 python3。请先安装 Python 3。"
  read -k 1 "reply?按任意键退出..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

if ! .venv/bin/python3 -c "import streamlit, docx" >/dev/null 2>&1; then
  echo "第一次启动需要安装依赖，请保持网络连接。"
  .venv/bin/python3 -m pip install -r requirements.txt
fi

.venv/bin/python3 -m streamlit run app.py --browser.gatherUsageStats false
