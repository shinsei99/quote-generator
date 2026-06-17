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
QTY_KEYWORDS = ["数量", "個数", "qty", "quantity"]
UNIT_KEYWORDS = ["単位", "unit"]
SPEC_KEYWORDS = ["規格", "仕様", "寸法", "spec"]
PRICE_KEYWORDS = ["単価", "価格", "unit price", "price"]
AMOUNT_KEYWORDS = ["金額", "amount", "total"]

# 行内の「金額っぽい数字トークン」を抽出する正規表現（カンマ区切り・円・¥対応）
NUM_RE = re.compile(r"[¥￥]?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*円?")

# 品目候補として扱わない行のキーワード（日付欄・合計行・連絡先など）
EXCLUDE_LINE_KEYWORDS = [
    "見積番号", "発行日", "御中", "様", "TEL", "FAX", "〒", "合計", "小計",
    "消費税", "税込", "税抜", "振込先", "備考", "有効期限", "登録番号",
    "工事件名", "工事場所", "工事概要", "工事期間", "支払条件", "次頁",
    "Page", "page", "@", "印",
]


def _normalize_text(s: str) -> str:
    """全角/半角スペースを除去して比較しやすくする（「名　称」のような
    見出しの中に挿入されたスペースでキーワード一致が崩れるのを防ぐ）。
    """
    return re.sub(r"[\s　]+", "", s).lower()

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
    norm_headers = [_normalize_text(h) for h in header_texts]
    norm_keywords = [_normalize_text(k) for k in keywords]
    for i, h in enumerate(norm_headers):
        for kw in norm_keywords:
            if kw and kw in h:
                return i
    return None


# 「小計」「合計」などの集計行とみなすキーワード（実際の金額は新しい上乗せ率
# で再計算するため、元の数字は使わず「ここでグループを区切る」印として扱う）
_SUBTOTAL_MARKERS = ("小計", "合計", "【", "】")
# 品目表の中に現れる、品目でも区切りでもない行（ページ送り表記など）
_IGNORE_MARKERS = ("次頁", "前頁")


def parse_table(table: list[list]) -> list[dict]:
    """テーブルを解析し、品目に加えて「工事区分」の見出し行・グループの
    区切り（元の【小計】の位置）も種別付きで返す。

    元の見積書が工事箇所ごとに見出し＋小計で区切られている場合、その構造を
    できるだけ保つため、
      - 数量・単価・金額がすべて空の行 → 「工事区分」の見出し
      - 「小計」「合計」を含む行 → 「区切り」（その時点までの品目で
        小計を作る合図。名前付きの見出しが無くても区切りだけのグループに
        なる場合がある＝単独項目だけのグループなど、業者によって書き方が
        違うケースに対応するため）
    として扱う。
    """
    if not table or len(table) < 2:
        return []
    header = [str(c or "").strip() for c in table[0]]
    col_name = find_col(header, ITEM_NAME_KEYWORDS)
    col_qty = find_col(header, QTY_KEYWORDS)
    col_unit = find_col(header, UNIT_KEYWORDS)
    col_spec = find_col(header, SPEC_KEYWORDS)
    col_price = find_col(header, PRICE_KEYWORDS)
    col_amount = find_col(header, AMOUNT_KEYWORDS)
    if col_name is None or (col_qty is None and col_price is None and col_amount is None):
        return []

    rows = []
    for row in table[1:]:
        if col_name >= len(row):
            continue
        name = str(row[col_name] or "").strip()
        if not name or any(m in name for m in _IGNORE_MARKERS):
            continue
        if any(m in name for m in _SUBTOTAL_MARKERS):
            rows.append({"種別": "小計区切り"})
            continue

        qty = to_number(row[col_qty]) if col_qty is not None and col_qty < len(row) else None
        unit = str(row[col_unit] or "").strip() if col_unit is not None and col_unit < len(row) else ""
        spec = str(row[col_spec] or "").strip() if col_spec is not None and col_spec < len(row) else ""
        price = to_number(row[col_price]) if col_price is not None and col_price < len(row) else None
        amount = to_number(row[col_amount]) if col_amount is not None and col_amount < len(row) else None

        if qty is None and price is None and amount is None:
            # 数量・単価・金額がすべて空＝工事区分の見出し行
            rows.append({"種別": "工事区分", "品名": name})
            continue

        eff_qty = qty if qty else 1
        if amount:
            # 単価が空欄でも金額は入っていることがあるため、金額を優先して
            # 単価を逆算する（金額が最終的に正しい数字であるケースが多いため）
            eff_price = amount / eff_qty if eff_qty else amount
        elif price is not None:
            eff_price = price
        else:
            continue

        rows.append({
            "種別": "品目", "品名": name, "規格": spec, "単位": unit,
            "数量": eff_qty, "元単価": round(eff_price, 2),
        })
    return rows


_LEADING_LIST_NUMBER_RE = re.compile(r"^[0-9]+[\.\)、]\s*")

# 「式」「台」など、数量1件分をまとめて計上する単位（単価が空欄で、
# 数字が1つ（金額のみ）しか出てこない行を救済するために使う）
_LUMP_SUM_UNIT_WORDS = ("式", "台", "ｾｯﾄ", "セット", "本", "枚", "箇所", "か所", "件")
_LUMP_SUM_UNIT_RE = re.compile("(" + "|".join(_LUMP_SUM_UNIT_WORDS) + r")\s*$")

# 値引・割引行（マイナス金額として品目に取り込む）
_DISCOUNT_KEYWORDS = ("値引", "割引")
_NEGATIVE_NUM_RE = re.compile(r"[−\-－]\s*([0-9][0-9,]*(?:\.[0-9]+)?)")


def parse_line_to_item(line: str) -> dict | None:
    line = line.strip()
    if not line or len(line) < 2:
        return None
    if any(kw in line for kw in _DISCOUNT_KEYWORDS):
        m = _NEGATIVE_NUM_RE.search(line) or NUM_RE.search(line)
        if not m:
            return None
        amount = float(m.group(1).replace(",", ""))
        if amount <= 0:
            return None
        return {"種別": "品目", "品名": "値引", "規格": "", "数量": 1, "単位": "", "元単価": -amount}
    if any(kw in line for kw in EXCLUDE_LINE_KEYWORDS):
        return None

    # 「1.玄関」のような先頭の箇条書き番号は、品名の前の「最初の数字」として
    # 誤検出され品名抽出を壊すため、品名探索の対象からは取り除いておく
    list_prefix_match = _LEADING_LIST_NUMBER_RE.match(line)
    search_line = line[list_prefix_match.end():] if list_prefix_match else line

    numbers = list(NUM_RE.finditer(search_line))

    if len(numbers) == 1:
        # 数字が金額しか無くても、直前に「式」「台」等の単位らしき語があれば
        # 数量1の一括計上とみなして救済する（OCRで数量の"1"だけ読み落とした
        # 場合などに有効）
        before = search_line[: numbers[0].start()]
        if not _LUMP_SUM_UNIT_RE.search(before.strip()):
            return None
        name = before.strip(" \t-:|　")
        # 末尾の単位語を品名から切り離す
        m = _LUMP_SUM_UNIT_RE.search(name)
        if m:
            name = name[: m.start()].strip(" \t-:|　")
        if not name:
            return None
        amount = float(numbers[0].group(1).replace(",", ""))
        if amount <= 0:
            return None
        return {"種別": "品目", "品名": name, "規格": "", "数量": 1, "単位": "", "元単価": amount}

    if len(numbers) < 2:
        return None

    name = search_line[: numbers[0].start()].strip(" \t-:|　")
    if not name:
        return None

    values = [float(m.group(1).replace(",", "")) for m in numbers]
    if len(values) >= 3:
        qty, price = values[-3], values[-2]
    else:
        # 数字が2つだけの場合、単価欄が空欄で「数量・金額」の2列しか
        # 無いケースが多い（例：「1　式　1000」）。数量を先頭、金額を
        # 末尾として扱い、単価は金額÷数量で逆算する。
        qty = values[-2]
        amount = values[-1]
        price = amount / qty if qty else amount

    if price <= 0:
        return None

    return {"種別": "品目", "品名": name, "規格": "", "数量": qty if qty > 0 else 1, "単位": "", "元単価": price}


# OCRがリフォーム見積書でよく誤読する単語の補正辞書（パラメータ調整では
# 改善しない、フォント固有の字形誤認識を補う最後の手段）。
# あくまで既知の頻出パターンのみを対象とし、過剰な書き換えは避ける。
_OCR_TERM_FIXES: list[tuple[re.Pattern, str]] = [
    (re.compile(r"[鮎船則脂刈]替"), "貼替"),          # 貼替 ⇔ 鮎替/船替/則替/脂替/刈替
    (re.compile(r"クス(?=貼替)"), "クロス"),           # クス貼替 → クロス貼替
    (re.compile(r"(ツ|ゆ|Y|ソ)ト木"), "ソフト巾木"),    # ソフト巾木の誤読バリエーション
    (re.compile(r"諸経[葵萎]"), "諸経費"),
    (re.compile(r"^遊面所"), "洗面所"),
    (re.compile(r"^ト[ルI](?=\s|$|　)"), "トイレ"),
]

# リフォーム工事の見積書でよく使われる定型語彙。OCRの誤読を曖昧一致で
# 正規化するために使う（語彙はある程度決まっているため、近い文字列が
# 見つかれば誤読とみなして正規の表記に寄せる）。
COMMON_REFORM_TERMS: list[str] = [
    # 内装・床・壁
    "クロス貼替", "壁紙", "フローリング", "クッションフロア", "畳表替え", "カーペット",
    "タイル張替", "巾木", "ソフト巾木", "廻り縁", "見切り材", "床下地補修", "天井",
    # 建具
    "ふすま", "障子", "扉", "ドア", "サッシ", "窓",
    # 水回り・設備
    "キッチン", "流し台", "レンジフード", "換気扇", "洗面所", "洗面台", "洗面化粧台", "浴槽",
    "ユニットバス", "給湯器", "トイレ", "便器", "エアコン", "リモコン", "キッチンパネル",
    # 工事種別
    "解体", "撤去", "設置", "取付", "交換", "補修", "張替", "貼替", "塗装", "防水",
    "クリーニング", "コーキング",
    # 諸経費系
    "諸経費", "運搬費", "交通費", "廃材処分費", "発生材処分費", "養生費",
]


def fix_common_ocr_terms(text: str, custom_corrections: dict[str, str] | None = None) -> str:
    """OCR由来の品名に対して、頻出する字形誤認識を補正する。
    デジタルPDF/Excelから取得したテキストには適用しない（必要ないため）。

    1. ユーザー登録の辞書（完全一致の文字列置換）を、生のOCRテキストに適用
    2. 既知の頻出パターン（正規表現）で補正
    3. ユーザー登録の辞書を再度適用（手順2で文字列が変わり、画面で見た
       「誤った表記」と一致しなくなるケースを取りこぼさないため）
    4. 定型語彙との曖昧一致で、近い文字列を正規の表記に寄せる
    """
    def _apply_custom(t: str) -> str:
        if not custom_corrections:
            return t
        for wrong, correct in custom_corrections.items():
            if wrong:
                t = t.replace(wrong, correct)
        return t

    text = _apply_custom(text)
    for pattern, repl in _OCR_TERM_FIXES:
        text = pattern.sub(repl, text)
    text = _apply_custom(text)
    text = _fuzzy_correct_against_vocabulary(text, COMMON_REFORM_TERMS)
    return text


def _fuzzy_correct_against_vocabulary(text: str, vocabulary: list[str], threshold: float = 0.78) -> str:
    """text内の部分文字列を、語彙リストとの類似度でスキャンし、十分近ければ
    正規の表記に置き換える（長い語彙から先に処理する）。
    閾値は高め（0.78）にしてあり、誤って「既に正しい別の語彙」を別の語彙に
    書き換えてしまう誤検出（例：クリーニング→コーキング）を避ける。
    どの語彙にも完全一致する部分文字列は、対象外として一切書き換えない。
    """
    import difflib

    vocab_set = set(vocabulary)
    result = text
    for term in sorted(vocabulary, key=len, reverse=True):
        n = len(term)
        if n < 3 or n > len(result):
            continue
        best_ratio = 0.0
        best_start = -1
        for start in range(len(result) - n + 1):
            window = result[start:start + n]
            if window == term:
                best_start = -1
                break  # 既に正しいので置換不要
            if window in vocab_set:
                continue  # 既に別の正しい語彙と完全一致するので触らない
            ratio = difflib.SequenceMatcher(None, window, term).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_start = start
        if best_start >= 0 and best_ratio >= threshold:
            result = result[:best_start] + term + result[best_start + n:]
    return result


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


def _group_ocr_results_into_lines(results, y_tolerance: float = 15.0) -> list[tuple[float, str]]:
    """EasyOCRの検出結果(座標+テキスト)を、同じ行（Y座標が近い）ごとに
    X座標順で連結し、表の1行のようなテキストに再構成する。
    戻り値は (その行のY座標, 連結したテキスト) のリスト（Y座標で判定し、
    宛先・住所など表より上にある行を除外できるようにするため）。
    """
    boxes = []
    for bbox, text, _conf in results:
        ys = [p[1] for p in bbox]
        xs = [p[0] for p in bbox]
        boxes.append((sum(ys) / len(ys), min(xs), text))
    boxes.sort(key=lambda b: (b[0], b[1]))

    rows: list[list[tuple[float, str]]] = []
    row_ys: list[list[float]] = []
    current_row: list[tuple[float, str]] = []
    current_row_ys: list[float] = []
    current_y = None
    for y, x, text in boxes:
        if current_y is None or abs(y - current_y) <= y_tolerance:
            current_row.append((x, text))
            current_row_ys.append(y)
            current_y = y if current_y is None else (current_y + y) / 2
        else:
            rows.append(current_row)
            row_ys.append(current_row_ys)
            current_row = [(x, text)]
            current_row_ys = [y]
            current_y = y
    if current_row:
        rows.append(current_row)
        row_ys.append(current_row_ys)

    lines = []
    for row, ys in zip(rows, row_ys):
        row.sort(key=lambda t: t[0])
        text = "  ".join(t[1] for t in row)
        lines.append((sum(ys) / len(ys), text))
    return lines


OCR_ZOOM = 4.0  # 高いほど文字認識精度は上がるが処理時間も伸びる（4倍で約300dpi相当）


# 品目テーブルの列見出し行を見つけるためのキーワード（2つ以上一致した行を
# 見出しとみなす）。これより前（宛先・住所・郵便番号・日付など）は対象外にする。
_HEADER_LINE_KEYWORDS = ["品名", "名称", "数量", "単価", "金額"]


def _looks_like_table_header(text: str) -> bool:
    norm = _normalize_text(text)
    hits = sum(1 for kw in _HEADER_LINE_KEYWORDS if _normalize_text(kw) in norm)
    return hits >= 2


def extract_items_via_ocr(file_bytes: bytes, custom_corrections: dict[str, str] | None = None) -> list[dict]:
    import fitz  # PyMuPDF

    reader = get_easyocr_reader()
    items = []
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    header_found = False
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(OCR_ZOOM, OCR_ZOOM))
        img_bytes = pix.tobytes("png")
        results = reader.readtext(img_bytes, detail=1, paragraph=False)
        for _y, text in _group_ocr_results_into_lines(results, y_tolerance=15.0 * OCR_ZOOM / 2.5):
            if not header_found:
                if _looks_like_table_header(text):
                    header_found = True
                continue
            parsed = parse_line_to_item(_fix_ocr_digit_confusions(text))
            if parsed:
                parsed["品名"] = fix_common_ocr_terms(parsed["品名"], custom_corrections)
                items.append(parsed)
    return items


def extract_items_from_pdf(file_bytes: bytes, custom_corrections: dict[str, str] | None = None) -> tuple[list[dict], str]:
    """戻り値: (品目リスト, 使用した抽出方式)"""
    items = extract_table_items_from_pdf(file_bytes)
    if items:
        return items, "表形式抽出"

    items = extract_text_items_from_pdf(file_bytes)
    if items:
        return items, "テキスト抽出"

    items = extract_items_via_ocr(file_bytes, custom_corrections)
    return items, "OCR（画像認識）"


def extract_items_from_excel(file_bytes: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_idx = col_name = col_qty = col_unit = col_spec = col_price = None
    for i, row in enumerate(rows[:15]):
        texts = [str(c).strip() if c is not None else "" for c in row]
        cn = find_col(texts, ITEM_NAME_KEYWORDS)
        cq = find_col(texts, QTY_KEYWORDS)
        cu = find_col(texts, UNIT_KEYWORDS)
        cs = find_col(texts, SPEC_KEYWORDS)
        cp = find_col(texts, PRICE_KEYWORDS)
        if cn is not None and (cq is not None or cp is not None):
            header_idx, col_name, col_qty, col_unit, col_spec, col_price = i, cn, cq, cu, cs, cp
            break

    if header_idx is None:
        return []

    def _cell_text(row, col):
        if col is None or col >= len(row) or row[col] is None:
            return ""
        return str(row[col]).strip()

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
            "種別": "品目",
            "品名": str(name).strip(),
            "規格": _cell_text(row, col_spec),
            "単位": _cell_text(row, col_unit),
            "数量": qty if qty is not None else 1,
            "元単価": price if price is not None else 0,
        })
    return items
