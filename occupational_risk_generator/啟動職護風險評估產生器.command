#!/bin/zsh

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "第一次啟動：建立 Python 虛擬環境..."
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi

echo "啟動特約職護職業安全風險評估產生器..."
echo "瀏覽器開啟後，請使用 http://localhost:8501"
echo "若要停止程式，請在此視窗按 Control + C"

.venv/bin/streamlit run app.py
