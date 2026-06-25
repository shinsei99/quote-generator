#!/bin/bash
# 常時起動Mac で quote-generator を「自動起動・自動再起動」サービスとして登録する。
# quote-generator/deploy/ の中で  bash install-service.sh  と実行するだけ。
set -e

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"        # quote-generator/
LABEL="com.daikyo.quote-generator"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
RUNNER="$APP_DIR/deploy/run-quote-generator.sh"
LOG="$APP_DIR/deploy/streamlit.log"

echo "▶ アプリ場所 : $APP_DIR"

# --- 事前チェック ---
command -v python3 >/dev/null || { echo "✗ python3 が見つかりません"; exit 1; }
python3 -c "import streamlit, pandas, openpyxl" 2>/dev/null || {
  echo "… 依存パッケージを導入します (pip3 install --user)"
  python3 -m pip install --user -r "$APP_DIR/requirements.txt"
}
if ! command -v claude >/dev/null; then
  echo "⚠ 警告: 'claude' コマンドが見つかりません。"
  echo "  このアプリはAI解析に Claude Code CLI を使います。"
  echo "  Claude Code をインストールし、ログイン認証してから使ってください。"
fi
[ -f "$APP_DIR/data/issuers.csv" ] || echo "⚠ data/issuers.csv がありません（発行元情報。元のMacからコピーしてください）"

chmod +x "$RUNNER"
mkdir -p "$HOME/Library/LaunchAgents"

# --- LaunchAgent plist を生成 ---
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>/bin/bash</string>
    <string>$RUNNER</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$LOG</string>
  <key>StandardErrorPath</key><string>$LOG</string>
</dict>
</plist>
EOF

# --- 再読み込み ---
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"

IP="$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo '<このMacのIP>')"
echo
echo "✅ 登録完了。Mac起動時に自動立ち上げ＋落ちても自動再起動します。"
echo "   社内LANからのアクセスURL →  http://$IP:8503"
echo
echo "   ログ:        $LOG"
echo "   停止:        launchctl unload \"$PLIST\""
echo "   再開:        launchctl load \"$PLIST\""
