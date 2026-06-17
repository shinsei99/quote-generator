"""見積書PDF/Excelからの品目（品名・数量・単価）抽出ロジック。

抽出は次の優先順位で試みます。
  1. pdfplumber によるテーブル抽出（デジタルPDFで罫線がある場合に最も正確）
  2. pdfplumber によるテキスト抽出 + 行パターンの解析（罫線が無いデジタルPDF）
  3. PyMuPDF でページを画像化 + EasyOCR（スキャン画像PDFのフォールバック）

抽出精度には限界があるため、UI側で必ず編集・追加できるようにしてください。
"""
from __future__ import annotations

import io
import re

import pdfplumber
from openpyxl import load_workbook

ITEM_NAME_KEYWORDS = ["品名", "品目", "名称", "商品名", "item", "description"]
QTY_KEYWORDS = ["数量", "数 量", "個数", "qty", "quantity"]
PRICE_KEYWORDS = ["単価", "価格", "unit price", "price"]

# 行内の「金額っぽい数字トークン」を抽出する正規表現（カンマ区切り・円・¥対応）
NUM_RE = re.compile(r"[¥￥]?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*円?")

# 品目候補として扱わない行のキーワード（日付欄・合計行・連絡先など）
EXCLUDE_LINE_KEYWORDS = [
    "見積番号", "発行日", "御中", "TEL", "FAX", "〒", "合計", "小計",
    "消費税", "税込", "税抜", "振込先", "備考", "有効期限",
]

_easyocr_reader = None


def to_number(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    m = NUM_RE.search(s)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def find_col(header_texts: list[str], keywords: list[str]) -> int | None:
    for i, h in enumerate(header_texts):
        h_lower = h.lower()
        for kw in keywords:
            if kw.lower() in h_lower:
                return i
    return None


def parse_table(table: list[list]) -> list[dict]:
    if not table or len(table) < 2:
        return []
    header = [str(c or "").strip() for c in table[0]]
    col_name = find_col(header, ITEM_NAME_KEYWORDS)
    col_qty = find_col(header, QTY_KEYWORDS)
    col_price = find_col(header, PRICE_KEYWORDS)
    if col_name is None or (col_qty is None and col_price is None):
        return []

    rows = []
    for row in table[1:]:
        if col_name >= len(row):
            continue
        name = str(row[col_name] or "").strip()
        if not name:
            continue
        qty = to_number(row[col_qty]) if col_qty is not None and col_qty < len(row) else None
        price = to_number(row[col_price]) if col_price is not None and col_price < len(row) else None
        rows.append({"品名": name, "数量": qty if qty is not None else 1, "元単価": price if price is not None else 0})
    return rows


def parse_line_to_item(line: str) -> dict | None:
    line = line.strip()
    if not line or len(line) < 2:
        return None
    if any(kw in line for kw in EXCLUDE_LINE_KEYWORDS):
        return None

    numbers = list(NUM_RE.finditer(line))
    if len(numbers) < 2:
        return None

    name = line[: numbers[0].start()].strip(" \t-:|　")
    if not name:
        return None

    values = [float(m.group(1).replace(",", "")) for m in numbers]
    if len(values) >= 3:
        qty, price = values[-3], values[-2]
    else:
        price = values[-2]
        amount = values[-1]
        qty = round(amount / price) if price else 1

    if price <= 0:
        return None

    return {"品名": name, "数量": qty if qty > 0 else 1, "元単価": price}


def extract_table_items_from_pdf(file_bytes: bytes) -> list[dict]:
    items = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                items.extend(parse_table(table))
    return items


def extract_text_items_from_pdf(file_bytes: bytes) -> list[dict]:
    items = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                parsed = parse_line_to_item(line)
                if parsed:
                    items.append(parsed)
    return items


def pdf_has_extractable_text(file_bytes: bytes) -> bool:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            if (page.extract_text() or "").strip():
                return True
    return False


def get_easyocr_reader():
    global _easyocr_reader
    if _easyocr_reader is None:
        import easyocr
        _easyocr_reader = easyocr.Reader(["ja", "en"], gpu=False)
    return _easyocr_reader


_OCR_DIGIT_FIX = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "B": "8"})
_NUMLIKE_RE = re.compile(r"[0-9OolIZzSsBb,]{2,}")


def _fix_ocr_digit_confusions(text: str) -> str:
    """OCRが数字を似た形のアルファベットと誤認識しやすい問題（O⇔0等）を補正する。"""
    def repl(m: re.Match) -> str:
        token = m.group(0)
        if not any(c.isdigit() for c in token):
            return token
        return token.translate(_OCR_DIGIT_FIX)
    return _NUMLIKE_RE.sub(repl, text)


def _group_ocr_results_into_lines(results, y_tolerance: float = 15.0) -> list[str]:
    """EasyOCRの検出結果(座標+テキスト)を、同じ行（Y座標が近い）ごとに
    X座標順で連結し、表の1行のようなテキストに再構成する。
    """
    boxes = []
    for bbox, text, _conf in results:
        ys = [p[1] for p in bbox]
        xs = [p[0] for p in bbox]
        boxes.append((sum(ys) / len(ys), min(xs), text))
    boxes.sort(key=lambda b: (b[0], b[1]))

    rows: list[list[tuple[float, str]]] = []
    current_row: list[tuple[float, str]] = []
    current_y = None
    for y, x, text in boxes:
        if current_y is None or abs(y - current_y) <= y_tolerance:
            current_row.append((x, text))
            current_y = y if current_y is None else (current_y + y) / 2
        else:
            rows.append(current_row)
            current_row = [(x, text)]
            current_y = y
    if current_row:
        rows.append(current_row)

    lines = []
    for row in rows:
        row.sort(key=lambda t: t[0])
        lines.append("  ".join(t[1] for t in row))
    return lines


def extract_items_via_ocr(file_bytes: bytes) -> list[dict]:
    import fitz  # PyMuPDF

    reader = get_easyocr_reader()
    items = []
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
        img_bytes = pix.tobytes("png")
        results = reader.readtext(img_bytes, detail=1, paragraph=False)
        for line in _group_ocr_results_into_lines(results):
            parsed = parse_line_to_item(_fix_ocr_digit_confusions(line))
            if parsed:
                items.append(parsed)
    return items


def extract_items_from_pdf(file_bytes: bytes) -> tuple[list[dict], str]:
    """戻り値: (品目リスト, 使用した抽出方式)"""
    items = extract_table_items_from_pdf(file_bytes)
    if items:
        return items, "表形式抽出"

    items = extract_text_items_from_pdf(file_bytes)
    if items:
        return items, "テキスト抽出"

    items = extract_items_via_ocr(file_bytes)
    return items, "OCR（画像認識）"


def extract_items_from_excel(file_bytes: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_idx = col_name = col_qty = col_price = None
    for i, row in enumerate(rows[:15]):
        texts = [str(c).strip() if c is not None else "" for c in row]
        cn = find_col(texts, ITEM_NAME_KEYWORDS)
        cq = find_col(texts, QTY_KEYWORDS)
        cp = find_col(texts, PRICE_KEYWORDS)
        if cn is not None and (cq is not None or cp is not None):
            header_idx, col_name, col_qty, col_price = i, cn, cq, cp
            break

    if header_idx is None:
        return []

    items = []
    for row in rows[header_idx + 1:]:
        if col_name >= len(row):
            continue
        name = row[col_name]
        if name is None or str(name).strip() == "":
            continue
        qty = to_number(row[col_qty]) if col_qty is not None and col_qty < len(row) else None
        price = to_number(row[col_price]) if col_price is not None and col_price < len(row) else None
        items.append({
            "品名": str(name).strip(),
            "数量": qty if qty is not None else 1,
            "元単価": price if price is not None else 0,
        })
    return items
