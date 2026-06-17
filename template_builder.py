from __future__ import annotations

import io
import math
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import range_boundaries

import template_config as cfg

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
CENTER = Alignment(horizontal="center", vertical="center")

# 品目データ行の列構成（結合する終了列。Noneは単一セル）
DATA_COL_SPECS = [
    (cfg.COL_ITEM_NAME, "C"), (cfg.COL_SPEC, "F"), (cfg.COL_QTY, None),
    (cfg.COL_UNIT, None), (cfg.COL_UNIT_PRICE, None), (cfg.COL_AMOUNT, "K"), (cfg.COL_NOTE, "N"),
]


def _merge(ws, cell_range: str):
    ws.merge_cells(cell_range)


def _set(ws, cell_range: str, value=None, font=None, fill=None, align=None, border=None, number_format=None):
    """セル（結合の場合は左上セル）に値・書式を設定するヘルパー。"""
    if ":" in cell_range:
        _merge(ws, cell_range)
        top_left = cell_range.split(":")[0]
    else:
        top_left = cell_range
    cell = ws[top_left]
    if value is not None:
        cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if number_format:
        cell.number_format = number_format
    if border:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).border = border
    return cell


NORMAL_FONT = Font(size=11, bold=False)
NORMAL_ALIGN = Alignment(horizontal="general", vertical="bottom")
NO_FILL = PatternFill(fill_type=None)


def _style_data_row(ws, row: int):
    """指定した1行を、品目データ行と同じ結合・罫線・フォントにする（流用元の太字や
    フォントサイズが残らないよう、フォント・配置・塗りつぶしも明示的にリセットする）。
    """
    for start_col, end_col in DATA_COL_SPECS:
        cell_range = f"{start_col}{row}:{end_col}{row}" if end_col else f"{start_col}{row}"
        _set(ws, cell_range, font=NORMAL_FONT, fill=NO_FILL, align=NORMAL_ALIGN, border=BORDER)


def _clear_row_range(ws, row_start: int, row_end: int):
    """指定した行範囲の値・結合をすべて解除する（サマリー欄を移動する前の下準備）。"""
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= row_start and m.max_row <= row_end:
            ws.unmerge_cells(str(m))
    for row in range(row_start, row_end + 1):
        for col in range(1, 15):
            ws.cell(row=row, column=col).value = None


def _build_summary_block(ws, subtotal_row: int, tax_row: int, total_row: int, tax_rate: float = 10):
    """小計・消費税・合計の3行ブロックを指定した行位置に作成する。"""
    header_font = Font(bold=True)

    _set(ws, f"A{subtotal_row}:{cfg.LABEL_COL_END_SUBTOTAL}{subtotal_row}", "小　計",
         font=header_font, align=Alignment(horizontal="center"), border=BORDER)
    _set(ws, f"{cfg.VALUE_COL_START}{subtotal_row}:{cfg.VALUE_COL_END}{subtotal_row}", 0,
         align=Alignment(horizontal="right"), border=BORDER, number_format="#,##0")

    _set(ws, f"A{tax_row}:{cfg.LABEL_COL_END_SUBTOTAL}{tax_row}", "消　費　税",
         font=header_font, align=Alignment(horizontal="center"), border=BORDER)
    _set(ws, f"{cfg.TAX_RATE_COL}{tax_row}", tax_rate, align=CENTER, border=BORDER)
    _set(ws, f"{cfg.TAX_RATE_UNIT_COL}{tax_row}", "％", align=CENTER, border=BORDER)
    _set(ws, f"{cfg.VALUE_COL_START}{tax_row}:{cfg.VALUE_COL_END}{tax_row}", 0,
         align=Alignment(horizontal="right"), border=BORDER, number_format="#,##0")

    total_row_end = total_row + 1
    _set(ws, f"A{total_row}:{cfg.LABEL_COL_END_TOTAL}{total_row_end}", "合　計",
         font=Font(bold=True, size=13), align=CENTER, border=BORDER)
    _set(ws, f"{cfg.VALUE_COL_START}{total_row}:{cfg.VALUE_COL_END}{total_row_end}", 0,
         font=Font(bold=True, size=13), align=Alignment(horizontal="right", vertical="center"),
         border=BORDER, number_format="#,##0")


def ensure_template_exists(path: Path | None = None) -> Path:
    """template.xlsx が無ければ、ユーザー提供の見積書テンプレートを参考にしたサンプルを新規作成する。"""
    path = Path(path) if path else Path(cfg.TEMPLATE_PATH)
    if path.exists():
        return path

    wb = Workbook()
    ws = wb.active
    ws.title = cfg.SHEET_NAME

    title_font = Font(size=20, bold=True)
    header_font = Font(bold=True)
    small_font = Font(size=9)
    bold_font = Font(bold=True)

    # タイトル
    _set(ws, "A1:N2", "御　見　積　書", font=title_font, align=Alignment(horizontal="center", vertical="center"))

    # 宛先
    _set(ws, f"{cfg.CELL_CLIENT}:D4", "（宛先を入力）", font=Font(size=14),
         align=Alignment(horizontal="left", vertical="bottom"), border=BORDER)
    _set(ws, f"{cfg.CELL_CLIENT_SUFFIX}:E4", "様", font=Font(size=14), align=CENTER)
    _set(ws, f"{cfg.CELL_ISSUE_DATE}:N4", None, align=Alignment(horizontal="right"), number_format="yyyy年mm月dd日")

    # 許可番号など（任意）
    _set(ws, f"{cfg.CELL_ISSUER_LICENSE}:N6", "", font=small_font, align=Alignment(horizontal="right"))

    # 挨拶文
    _set(ws, cfg.CELL_GREETING, "下記の通りお見積申し上げますので何卒宜しくお願い致します。")

    # 発行元情報
    _set(ws, f"{cfg.CELL_ISSUER_NAME}:N8", "", font=bold_font, align=Alignment(horizontal="right"))
    _set(ws, f"{cfg.CELL_ISSUER_ADDRESS}:N9", "", align=Alignment(horizontal="right"))
    _set(ws, cfg.CELL_ISSUER_TEL, "", align=Alignment(horizontal="right"))
    _set(ws, cfg.CELL_ISSUER_FAX, "", align=Alignment(horizontal="right"))
    _set(ws, cfg.CELL_ISSUER_REG_NO, "", align=Alignment(horizontal="right"))

    # 見積金額サマリー
    _set(ws, f"{cfg.CELL_TOTAL_LABEL}:D14", "見積金額", font=header_font, align=CENTER, border=BORDER)
    _set(ws, f"{cfg.CELL_TOTAL_DISPLAY}:G14", 0, font=Font(size=16, bold=True),
         align=Alignment(horizontal="right", vertical="center"), border=BORDER, number_format='#,##0"円（税込）"')

    # 検印欄
    _set(ws, "L13:N13", "検　印", font=header_font, align=CENTER, border=BORDER)
    for col in ("L", "M", "N"):
        _set(ws, f"{col}14:{col}16", "", border=BORDER)

    # 品目テーブルのヘッダー
    headers = {
        f"{cfg.COL_ITEM_NAME}:C": "名称",
        f"{cfg.COL_SPEC}:F": "寸法・規格",
        cfg.COL_QTY: "数量",
        cfg.COL_UNIT: "単位",
        cfg.COL_UNIT_PRICE: "単価",
        f"{cfg.COL_AMOUNT}:K": "金額",
        f"{cfg.COL_NOTE}:N": "備考",
    }
    header_row = cfg.COLUMN_HEADER_ROW
    for col_spec, label in headers.items():
        if ":" in col_spec:
            start_col, end_col = col_spec.split(":")
            cell_range = f"{start_col}{header_row}:{end_col}{header_row}"
        else:
            cell_range = f"{col_spec}{header_row}"
        _set(ws, cell_range, label, font=header_font, fill=HEADER_FILL, align=CENTER, border=BORDER)

    # 品目データ行（罫線・結合のみ用意）
    for r in range(cfg.DATA_START_ROW, cfg.DATA_START_ROW + cfg.MAX_DATA_ROWS):
        _style_data_row(ws, r)

    # 小計・消費税・合計
    _build_summary_block(ws, cfg.SUBTOTAL_ROW, cfg.TAX_ROW, cfg.TOTAL_ROW)

    widths = {
        "A": 10, "B": 8, "C": 8, "D": 8, "E": 8, "F": 8,
        "G": 8, "H": 6, "I": 10, "J": 8, "K": 8, "L": 8, "M": 8, "N": 10,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def fill_template(items: list[dict], header_info: dict, issuer_info: dict,
                   tax_rate: float, template_path: Path | None = None) -> bytes:
    """items: [{"品名","数量","上乗せ後単価"}]（税別）を template に流し込み、
    小計・消費税（小数点以下切り捨て）・合計まで計算して bytes を返す。

    品目数があらかじめ用意された行数（MAX_DATA_ROWS）を超える場合は、
    小計・消費税・合計欄を必要な分だけ下にずらして衝突を避ける。
    """
    template_path = Path(template_path) if template_path else Path(cfg.TEMPLATE_PATH)
    wb = load_workbook(template_path)
    ws = wb[cfg.SHEET_NAME] if cfg.SHEET_NAME in wb.sheetnames else wb.active

    if header_info.get("client"):
        ws[cfg.CELL_CLIENT] = header_info["client"]
    if header_info.get("issue_date"):
        ws[cfg.CELL_ISSUE_DATE] = header_info["issue_date"]

    if issuer_info.get("license"):
        ws[cfg.CELL_ISSUER_LICENSE] = issuer_info["license"]
    if issuer_info.get("name"):
        ws[cfg.CELL_ISSUER_NAME] = issuer_info["name"]
    if issuer_info.get("address"):
        ws[cfg.CELL_ISSUER_ADDRESS] = issuer_info["address"]
    if issuer_info.get("tel"):
        ws[cfg.CELL_ISSUER_TEL] = f"TEL　{issuer_info['tel']}"
    if issuer_info.get("fax"):
        ws[cfg.CELL_ISSUER_FAX] = f"FAX　{issuer_info['fax']}"
    if issuer_info.get("registration_no"):
        ws[cfg.CELL_ISSUER_REG_NO] = f"登録番号：{issuer_info['registration_no']}"

    subtotal_row, tax_row, total_row = cfg.SUBTOTAL_ROW, cfg.TAX_ROW, cfg.TOTAL_ROW
    overflow = max(0, len(items) - cfg.MAX_DATA_ROWS)

    if overflow > 0:
        # 元の小計〜合計ブロック（結合済み）を解除し、通常のデータ行として再利用する
        _clear_row_range(ws, cfg.SUBTOTAL_ROW, cfg.TOTAL_ROW + 1)
        for row in range(cfg.SUBTOTAL_ROW, cfg.TOTAL_ROW + 2):
            _style_data_row(ws, row)

        subtotal_row = cfg.SUBTOTAL_ROW + overflow
        tax_row = cfg.TAX_ROW + overflow
        total_row = cfg.TOTAL_ROW + overflow

        # 元のブロックより下、新しい小計欄より上の行にもスタイルを適用
        for row in range(cfg.TOTAL_ROW + 2, subtotal_row):
            _style_data_row(ws, row)

        _build_summary_block(ws, subtotal_row, tax_row, total_row, tax_rate=tax_rate)

    subtotal = 0
    for i, item in enumerate(items):
        row = cfg.DATA_START_ROW + i
        qty = item.get("数量", 0) or 0
        price = math.floor(item.get("上乗せ後単価", 0) or 0)
        amount = math.floor(qty * price)
        subtotal += amount
        ws[f"{cfg.COL_ITEM_NAME}{row}"] = item.get("品名", "")
        ws[f"{cfg.COL_QTY}{row}"] = qty
        ws[f"{cfg.COL_UNIT_PRICE}{row}"] = price
        ws[f"{cfg.COL_AMOUNT}{row}"] = amount

    tax_amount = math.floor(subtotal * tax_rate / 100)
    grand_total = subtotal + tax_amount

    ws[f"{cfg.VALUE_COL_START}{subtotal_row}"] = subtotal
    ws[f"{cfg.TAX_RATE_COL}{tax_row}"] = tax_rate
    ws[f"{cfg.VALUE_COL_START}{tax_row}"] = tax_amount
    ws[f"{cfg.VALUE_COL_START}{total_row}"] = grand_total
    ws[cfg.CELL_TOTAL_DISPLAY] = grand_total

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
