from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import template_config as cfg


def ensure_template_exists(path: Path | None = None) -> Path:
    """template.xlsx が無ければサンプルを新規作成する。"""
    path = Path(path) if path else Path(cfg.TEMPLATE_PATH)
    if path.exists():
        return path

    wb = Workbook()
    ws = wb.active
    ws.title = cfg.SHEET_NAME

    title_font = Font(size=20, bold=True)
    header_font = Font(bold=True)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E2F3")

    ws[cfg.CELL_TITLE] = "御 見 積 書"
    ws[cfg.CELL_TITLE].font = title_font

    ws[cfg.CELL_CLIENT] = "○○○○ 様"
    ws[cfg.CELL_SUBJECT] = "件名："
    ws[cfg.CELL_ISSUE_DATE] = "発行日："
    ws[cfg.CELL_QUOTE_NO] = "見積番号："

    headers = {
        cfg.COL_ITEM_NAME: "品名",
        cfg.COL_SPEC: "規格・仕様",
        cfg.COL_UNIT: "単位",
        cfg.COL_QTY: "数量",
        cfg.COL_UNIT_PRICE: "単価",
        cfg.COL_AMOUNT: "金額",
    }
    for col, label in headers.items():
        cell = ws[f"{col}{cfg.COLUMN_HEADER_ROW}"]
        cell.value = label
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for r in range(cfg.DATA_START_ROW, cfg.DATA_START_ROW + cfg.MAX_DATA_ROWS):
        for col in headers:
            ws[f"{col}{r}"].border = border

    total_row = cfg.DATA_START_ROW + cfg.MAX_DATA_ROWS
    ws[f"{cfg.TOTAL_LABEL_COL}{total_row}"] = "合計"
    ws[f"{cfg.TOTAL_LABEL_COL}{total_row}"].font = header_font

    widths = {
        cfg.COL_ITEM_NAME: 28, cfg.COL_SPEC: 16, cfg.COL_UNIT: 8,
        cfg.COL_QTY: 8, cfg.COL_UNIT_PRICE: 12, cfg.COL_AMOUNT: 14,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def fill_template(items: list[dict], header_info: dict, template_path: Path | None = None) -> bytes:
    """items: [{"品名","数量","上乗せ後単価","単位"(任意)}] を template に流し込み bytes を返す。"""
    template_path = Path(template_path) if template_path else Path(cfg.TEMPLATE_PATH)
    wb = load_workbook(template_path)
    ws = wb[cfg.SHEET_NAME] if cfg.SHEET_NAME in wb.sheetnames else wb.active

    if header_info.get("client"):
        ws[cfg.CELL_CLIENT] = f"{header_info['client']} 様"
    if header_info.get("subject"):
        ws[cfg.CELL_SUBJECT] = f"件名：{header_info['subject']}"
    if header_info.get("issue_date"):
        ws[cfg.CELL_ISSUE_DATE] = f"発行日：{header_info['issue_date']}"
    if header_info.get("quote_no"):
        ws[cfg.CELL_QUOTE_NO] = f"見積番号：{header_info['quote_no']}"

    grand_total = 0
    for i, item in enumerate(items):
        row = cfg.DATA_START_ROW + i
        qty = item.get("数量", 0) or 0
        price = item.get("上乗せ後単価", 0) or 0
        amount = qty * price
        grand_total += amount
        ws[f"{cfg.COL_ITEM_NAME}{row}"] = item.get("品名", "")
        ws[f"{cfg.COL_UNIT}{row}"] = item.get("単位", "")
        ws[f"{cfg.COL_QTY}{row}"] = qty
        ws[f"{cfg.COL_UNIT_PRICE}{row}"] = price
        ws[f"{cfg.COL_AMOUNT}{row}"] = amount

    total_row = cfg.DATA_START_ROW + max(cfg.MAX_DATA_ROWS, len(items))
    ws[f"{cfg.TOTAL_LABEL_COL}{total_row}"] = "合計"
    ws[f"{cfg.TOTAL_VALUE_COL}{total_row}"] = grand_total

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
