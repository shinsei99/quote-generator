#!/bin/bash
# quote-generator 起動ランチャー（launchd から呼ばれる）。
# launchd は最小PATHで実行するので、python3 と claude を確実に見つけられるよう PATH を補う。
export PATH="$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"
export PYTHONUNBUFFERED=1

# このスクリプトの1つ上 = quote-generator/ アプリ本体へ移動
APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$APP_DIR" || exit 1

exec python3 -m streamlit run app.py \
  --server.port 8503 \
  --server.address 0.0.0.0 \
  --server.headless true \
  --browser.gatherUsageStats false
