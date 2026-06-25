# quote-generator 社内サーバー化 手順（常時起動Mac向け）

社内LANの全員が、毎回起動せずに `http://<サーバーMacのIP>:8503` で使えるようにする手順。

## 前提
- このアプリはAI解析に **Claude Code CLI（`claude`コマンド）** を内部で呼び出す。
  → サーバーにするMacにも **Claude Code のインストールとログイン認証が必須**。
- 社内LAN内アクセスのみ（インターネット公開はしない）。

---

## 1. アプリ一式をサーバーMacにコピー
元のMacの `~/quote-generator/` フォルダ**まるごと**を、サーバーMacの `~/quote-generator/` に置く。
（AirDrop / USB / 社内ファイル共有など。`data/issuers.csv`＝発行元情報も忘れず含める）

## 2. Claude Code CLI を用意
サーバーMacで Claude Code をインストールし、ログイン認証する。
ターミナルで `claude` が起動し、ログイン済みになっていればOK。

## 3. サービス登録（自動起動）
サーバーMacのターミナルで：
```bash
cd ~/quote-generator/deploy
bash install-service.sh
```
- 依存パッケージ（streamlit/pandas/openpyxl）は自動導入。
- 完了時にアクセスURL（`http://<IP>:8503`）が表示される。
- 以後、Mac起動時に自動立ち上げ＋クラッシュ時も自動再起動。手動起動は不要。

## 4. サーバーMacを「常時稼働」にする
- **スリープ無効**：システム設定 →「ロック画面/バッテリー/電源」でディスプレイは消えてもスリープしない設定に。
  （ターミナルなら `sudo pmset -a sleep 0`）
- **自動ログイン**：再起動後に人手なしで復帰できるよう、システム設定 → ユーザとグループ → 自動ログインをオン。
  （LaunchAgentはログインユーザーのセッションで動くため）
- **ファイアウォール**：オンなら、python（streamlit）の外部接続受け入れを「許可」。

## 5. IP固定（推奨）
サーバーMacのLAN IPが変わるとURLも変わる。ルーターのDHCP予約でMACアドレスにIPを固定するか、
手動でIPアドレスを固定設定にしておくと、URLが常に同じで楽。

## 6. 社内に共有
表示されたURL `http://<サーバーMacのIP>:8503` を社内に周知（ブックマーク推奨）。

---

## 運用メモ
| やりたいこと | コマンド |
|---|---|
| ログ確認 | `tail -f ~/quote-generator/deploy/streamlit.log` |
| 一時停止 | `launchctl unload ~/Library/LaunchAgents/com.daikyo.quote-generator.plist` |
| 再開 | `launchctl load ~/Library/LaunchAgents/com.daikyo.quote-generator.plist` |
| 稼働確認 | ブラウザで `http://<IP>:8503` / `lsof -ti tcp:8503` |
| コード更新時 | フォルダを上書き → 上記 unload→load で再起動 |

> Linux / Windows をサーバーにする場合はこの手順（launchd）は使えない。
> その旨を伝えてくれれば systemd（Linux）/ タスクスケジューラ・NSSM（Windows）版に差し替える。
