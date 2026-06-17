"""自社テンプレート（template.xlsx）のセル配置設定。

実際の自社テンプレートに合わせて、ここの値だけを書き換えれば
app.py 側のロジックは変更しなくても流し込み先が変わります。
"""

TEMPLATE_PATH = "template.xlsx"
SHEET_NAME = "見積書"

# ── ヘッダー情報セル ─────────────────────────────────────────
CELL_TITLE = "A1"
CELL_CLIENT = "A3"        # 宛先
CELL_SUBJECT = "A4"       # 件名
CELL_ISSUE_DATE = "F3"    # 発行日
CELL_QUOTE_NO = "F4"      # 見積番号

# ── 品目テーブルのセル配置 ───────────────────────────────────
COLUMN_HEADER_ROW = 7      # 列見出し（品名・数量…）の行
DATA_START_ROW = 8         # 品目データの開始行
MAX_DATA_ROWS = 20         # テンプレートにあらかじめ罫線を引いておく行数

COL_ITEM_NAME = "A"        # 品名
COL_SPEC = "B"             # 規格・仕様（任意）
COL_UNIT = "C"             # 単位（任意）
COL_QTY = "D"              # 数量
COL_UNIT_PRICE = "E"       # 単価（上乗せ後）
COL_AMOUNT = "F"           # 金額（数量×単価）

TOTAL_LABEL_COL = "E"
TOTAL_VALUE_COL = "F"
