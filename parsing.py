"""見積書PDF/Excelからの品目（品名・数量・単価）抽出ロジック。

PDFは Claude Code CLI（`claude` コマンド）にファイルをそのまま読み込ませて
内容を理解させ、構造化JSONとして品目データを受け取る方式にしている。
pdfplumberでの表/テキスト抽出やOCR、正規表現によるフォーマット別の解析は
行わない（フォーマットが見積書ごとに異なっていても、Claude自身が文書を
理解して抽出するため、形式ごとの分岐が不要になる）。

Excel（.xlsx）はもともとフォーマットが扱いやすい構造化データのため、
従来どおり openpyxl でヘッダー列を検出して読み込む。

Anthropic APIキーは使用しない。ユーザーのClaude Code CLIログイン
（Claude Pro/Max 等のサブスクリプション）のみで動作する。
"""
from __future__ import annotations

import io
import json
import re
import subprocess
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ITEM_NAME_KEYWORDS = ["品名", "品目", "名称", "商品名", "item", "description"]
QTY_KEYWORDS = ["数量", "個数", "qty", "quantity"]
UNIT_KEYWORDS = ["単位", "unit"]
SPEC_KEYWORDS = ["規格", "仕様", "寸法", "spec"]
PRICE_KEYWORDS = ["単価", "価格", "unit price", "price"]
AMOUNT_KEYWORDS = ["金額", "amount", "total"]

NUM_RE = re.compile(r"[¥￥]?\s*(-?[0-9][0-9,]*(?:\.[0-9]+)?)\s*円?")


def _normalize_text(s: str) -> str:
    """全角/半角スペースを除去して比較しやすくする（「名　称」のような
    見出しの中に挿入されたスペースでキーワード一致が崩れるのを防ぐ）。
    """
    return re.sub(r"[\s　]+", "", s).lower()


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


# ── PDF: Claude Code CLI による品目抽出 ──────────────────────────────

CLAUDE_BIN = "claude"
CLAUDE_TIMEOUT_SEC = 180
CLAUDE_MODEL = "sonnet"

_ITEM_JSON_SCHEMA = {
    "type": "object",
    "properties": {
        "items": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "種別": {"type": "string"},
                    "品名": {"type": "string"},
                    "規格仕様": {"type": "string"},
                    "備考": {"type": "string"},
                    "数量": {"type": "string"},
                    "単位": {"type": "string"},
                    "単価": {"type": "string"},
                    "金額": {"type": "string"},
                },
                "required": ["品名"],
            },
        },
    },
    "required": ["items"],
}

_CLAUDE_PROMPT = (
    "日本語の見積書PDF（{filename}、このディレクトリ内）を読み込み、内容を理解してください。"
    "見積書ごとにレイアウトは統一されていません。書かれている品目データを漏れなく抽出し、"
    "指定したJSON Schemaに従ってJSONのみを返してください。\n"
    "ルール:\n"
    "・各品目の「種別」には、見積書内で実際にグルーピングされている工事区分・カテゴリ名を"
    "入れてください（例：＜給湯室改修工事＞のような大きな見出し、または同じ製品群でまとめられた"
    "小見出し）。明確なグルーピングが無い品目は種別を空文字にしてください。\n"
    "・小計・合計・中計・大計・消費税など、品目自体ではなく集計を表す行は絶対に items に"
    "含めないでください（【小計】のような記号つきの行も同様です）。\n"
    "・値引き・割引があれば、品名を「値引」とし、金額をマイナス値にして items に含めてください。\n"
    "・各品目に紐づく備考（注記・補足説明・＠単価などの参考表記）があれば「備考」に入れてください。\n"
    "・単価が見積書に書かれておらず金額のみ分かる場合は、単価を空文字のままにして構いません。\n"
    "・数量・単価・金額は数字のみ（カンマや円マークを含めない）にしてください。"
)

# Claudeが誤って集計行を品目として返してしまった場合に備えた、コード側での防御フィルター
_SUMMARY_ROW_MARKERS = ("小計", "合計", "中計", "大計", "【", "】")


class ClaudeExtractionError(RuntimeError):
    pass


def extract_items_from_pdf(file_bytes: bytes) -> tuple[list[dict], str]:
    """PDFをClaude Code CLIに読み込ませ、品目データを抽出する。
    戻り値: (品目リスト, 使用した抽出方式の説明)
    """
    with tempfile.TemporaryDirectory(prefix="quote_pdf_") as tmp_dir:
        tmp_path = Path(tmp_dir) / "quote.pdf"
        tmp_path.write_bytes(file_bytes)

        cmd = [
            CLAUDE_BIN, "-p", _CLAUDE_PROMPT.format(filename=tmp_path.name),
            "--output-format", "json",
            "--json-schema", json.dumps(_ITEM_JSON_SCHEMA, ensure_ascii=False),
            "--tools", "Read",
            "--add-dir", tmp_dir,
            "--dangerously-skip-permissions",
            "--model", CLAUDE_MODEL,
        ]
        try:
            proc = subprocess.run(
                cmd, cwd=tmp_dir, capture_output=True, text=True, timeout=CLAUDE_TIMEOUT_SEC,
            )
        except FileNotFoundError as e:
            raise ClaudeExtractionError(
                "`claude` コマンドが見つかりません。Claude Code CLI がインストールされ、"
                "PATH が通っていることを確認してください。"
            ) from e
        except subprocess.TimeoutExpired as e:
            raise ClaudeExtractionError(
                f"Claude による解析が{CLAUDE_TIMEOUT_SEC}秒を超えたため中断しました。"
            ) from e

        if proc.returncode != 0:
            raise ClaudeExtractionError(
                f"Claude の呼び出しに失敗しました（終了コード {proc.returncode}）。"
                f"\n{proc.stderr.strip()[:500]}"
            )

        try:
            result = json.loads(proc.stdout)
        except json.JSONDecodeError as e:
            raise ClaudeExtractionError("Claude の応答をJSONとして解釈できませんでした。") from e

        if result.get("is_error"):
            raise ClaudeExtractionError(f"Claude がエラーを返しました: {result.get('result')}")

        structured = result.get("structured_output")
        if not structured or "items" not in structured:
            raise ClaudeExtractionError("Claude の応答に品目データ（items）が含まれていませんでした。")

        cost = result.get("total_cost_usd")
        raw_items = structured["items"]

    entries = _build_entries_from_claude_items(raw_items)
    method = "Claude Code 解析" + (f"（コスト: ${cost:.3f}）" if cost is not None else "")
    return entries, method


def _build_entries_from_claude_items(raw_items: list[dict]) -> list[dict]:
    """Claudeが返した（種別=カテゴリ名のフラットな）品目リストを、
    アプリ内部の表現（品目／工事区分／小計区切り）に変換する。
    種別が前の行と変わるたびに区切りを入れ、空でなければ見出し行も追加する
    （これにより、元の見積書の工事箇所ごとの区切り・小計を再現する）。
    """
    entries: list[dict] = []
    prev_section: str | None = None

    for raw in raw_items:
        name = str(raw.get("品名", "")).strip()
        if not name:
            continue
        if name != "値引" and any(m in name for m in _SUMMARY_ROW_MARKERS):
            # Claudeが指示に反して小計・合計などの集計行を返した場合の防御フィルター
            continue

        section = str(raw.get("種別", "")).strip()
        if prev_section is not None and section != prev_section:
            entries.append({"種別": "小計区切り"})
        if section and section != prev_section:
            entries.append({"種別": "工事区分", "品名": section})
        prev_section = section

        qty = to_number(raw.get("数量"))
        price = to_number(raw.get("単価"))
        amount = to_number(raw.get("金額"))

        eff_qty = qty if qty else 1
        if amount is not None:
            eff_price = amount / eff_qty if eff_qty else amount
        elif price is not None:
            eff_price = price
        else:
            continue

        entries.append({
            "種別": "品目",
            "品名": name,
            "規格": str(raw.get("規格仕様", "")).strip(),
            "備考": str(raw.get("備考", "")).strip(),
            "単位": str(raw.get("単位", "")).strip(),
            "数量": eff_qty,
            "元単価": round(eff_price, 2),
        })

    return entries


# ── Excel: openpyxl による品目抽出（構造化データなので従来どおり） ─────

def extract_items_from_excel(file_bytes: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_idx = col_name = col_qty = col_unit = col_spec = col_price = col_note = None
    for i, row in enumerate(rows[:15]):
        texts = [str(c).strip() if c is not None else "" for c in row]
        cn = find_col(texts, ITEM_NAME_KEYWORDS)
        cq = find_col(texts, QTY_KEYWORDS)
        cu = find_col(texts, UNIT_KEYWORDS)
        cs = find_col(texts, SPEC_KEYWORDS)
        cp = find_col(texts, PRICE_KEYWORDS)
        cnote = find_col(texts, ["備考", "note", "remarks"])
        if cn is not None and (cq is not None or cp is not None):
            header_idx, col_name, col_qty, col_unit, col_spec, col_price, col_note = (
                i, cn, cq, cu, cs, cp, cnote,
            )
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
            "備考": _cell_text(row, col_note),
            "単位": _cell_text(row, col_unit),
            "数量": qty if qty is not None else 1,
            "元単価": price if price is not None else 0,
        })
    return items
