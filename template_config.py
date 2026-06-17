"""自社テンプレート（template.xlsx）のセル配置設定。

ユーザー提供の実物の見積書テンプレート
「コピー1F　請求書・見積書　リフォーム工事（内容変更後）.xls」のレイアウト
（タイトル結合・宛先欄・発行元欄・品目表・小計/消費税/合計欄の位置）を
参考に作成しています。実際の自社テンプレートに合わせる場合は、
ここの値だけを書き換えれば app.py 側のロジックは変更不要です。
"""

TEMPLATE_PATH = "template.xlsx"
SHEET_NAME = "見積書"

# ── タイトル・宛先 ───────────────────────────────────────────
CELL_TITLE = "A1"          # 結合 A1:N2
CELL_CLIENT = "A3"         # 宛先（結合 A3:D4）
CELL_CLIENT_SUFFIX = "E3"  # 「様」（結合 E3:E4、固定文言）
CELL_ISSUE_DATE = "K4"     # 発行日（結合 K4:N4）
CELL_GREETING = "B8"       # 固定の挨拶文

# ── 発行元（自社）情報 ───────────────────────────────────────
CELL_ISSUER_LICENSE = "I6"   # 許可番号など（結合 I6:N6、任意）
CELL_ISSUER_NAME = "K8"      # 発行元 会社名（結合 K8:N8）
CELL_ISSUER_ADDRESS = "K9"   # 発行元 住所（結合 K9:N9）
CELL_ISSUER_TEL = "N10"      # 発行元 TEL
CELL_ISSUER_FAX = "N11"      # 発行元 FAX
CELL_ISSUER_REG_NO = "N12"   # インボイス登録番号

# ── 見積金額の上部サマリー表示 ──────────────────────────────
CELL_TOTAL_LABEL = "C13"     # 「見積金額」（結合 C13:D14、固定文言）
CELL_TOTAL_DISPLAY = "E13"   # 合計金額を表示（結合 E13:G14）

# ── 品目テーブルのセル配置 ───────────────────────────────────
COLUMN_HEADER_ROW = 17     # 列見出し（名称・数量…）の行
DATA_START_ROW = 18        # 品目データの開始行
MAX_DATA_ROWS = 12         # テンプレートにあらかじめ罫線・結合を設定する行数

COL_ITEM_NAME = "A"        # 名称（結合 A:C）
COL_SPEC = "D"             # 寸法・規格（結合 D:F、任意）
COL_QTY = "G"              # 数量
COL_UNIT = "H"             # 単位
COL_UNIT_PRICE = "I"       # 単価（上乗せ後・税別）
COL_AMOUNT = "J"           # 金額（結合 J:K、数量×単価）
COL_NOTE = "L"             # 備考（結合 L:N、任意）

# ── 小計・消費税・合計 ───────────────────────────────────────
SUBTOTAL_ROW = 31           # 小計（税別の品目金額合計）
TAX_ROW = 32                # 消費税（小計 × 税率、小数点以下切り捨て）
TOTAL_ROW = 33              # 合計（小計＋消費税、結合 33:34）

LABEL_COL_START = "A"       # 小計・消費税・合計の行ラベル開始列
LABEL_COL_END_SUBTOTAL = "F"   # 小計・消費税のラベル結合終了列
LABEL_COL_END_TOTAL = "I"      # 合計のラベル結合終了列（より広い）
TAX_RATE_COL = "G"          # 消費税率の数値（例: 10）
TAX_RATE_UNIT_COL = "H"     # 「％」
VALUE_COL_START = "J"       # 各行の金額表示開始列（結合 J:K）
VALUE_COL_END = "K"
