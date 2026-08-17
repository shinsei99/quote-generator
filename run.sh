#!/bin/bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q -r requirements.txt
fi
exec .venv/bin/streamlit run app.py --server.port 8503 --server.address 0.0.0.0 --server.headless true

# 注意: launchd（com.shinsei.quote-generator）はこのスクリプトを経由せず、
# /usr/bin/python3 -m streamlit run app.py を直接叩いている（plist 参照）。
# このスクリプトは手動起動・他PC用。venv を作るのでシステム側とは別環境になる。
