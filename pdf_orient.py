# -*- coding: utf-8 -*-
"""PDFの向き（縦横・回転）を解析前に自動補正する共有ユーティリティ。

スキャン謄本・スキャン書類は、ページ回転フラグと埋め込み画像の向きが食い違い、
横向き・逆さのまま AI（Claude のビジョン）に渡ると遅く・不正確になる。
本モジュールは「①各ページを画像化 → ②軽量モデル(haiku)で正立に必要な回転角を判定
→ ③正立画像に直す」を行い、正立済みの PDF / 画像を返す。

依存: PyMuPDF(fitz), Pillow, ローカルの `claude` CLI（APIキー不要）。
いずれも無い場合は安全に元データをそのまま返す（例外を投げない）。

使い方（最小変更でのドロップイン）:
    from pdf_orient import ensure_upright_pdf
    pdf_bytes = ensure_upright_pdf(pdf_bytes)   # ← 既存のClaude読み取りの直前に1行足すだけ
"""

import io
import json
import os
import shutil
import subprocess
import tempfile

ORIENT_MODEL = "haiku"     # 向き判定は軽量・高速モデル
RENDER_DPI = 170           # 画像化の解像度
THUMB_MAX = 760            # 向き判定用サムネイルの最大辺
MAX_PAGES = 20             # 補正対象ページ数の上限
ORIENT_TIMEOUT = 120       # 向き判定のタイムアウト（秒）
TEXT_MIN_CHARS = 40        # PDF全体でこれ以上テキストがあれば「テキストPDF」とみなし補正不要


def _claude_bin():
    """claude CLI のパスを解決する。見つからなければ None。"""
    p = shutil.which("claude")
    if p:
        return p
    cand = os.path.expanduser("~/.local/bin/claude")
    return cand if os.path.exists(cand) else None


def _pdf_text_chars(pdf_bytes: bytes) -> int:
    """PDFのテキスト層の文字数（概算）。fitz が無ければ -1。"""
    try:
        import fitz
    except ImportError:
        return -1
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        n = 0
        for page in doc:
            n += len((page.get_text() or "").strip())
            if n >= TEXT_MIN_CHARS:
                break
        doc.close()
        return n
    except Exception:
        return -1


def _detect_angle(thumb_name: str, cwd: str, claude_bin: str) -> int:
    """サムネイル画像が正立する時計回り回転角（0/90/180/270）を haiku で判定。"""
    prompt = (
        "この画像は日本語の書類をスキャンしたものです。"
        "文字が正しく正立して読める向きにするために、画像を時計回りに何度回転させればよいですか。"
        f"0・90・180・270 のいずれかの数字のみを答えてください。画像ファイル: {thumb_name}"
    )
    cmd = [
        claude_bin, "-p", prompt,
        "--output-format", "json",
        "--tools", "Read",
        "--add-dir", cwd,
        "--dangerously-skip-permissions",
        "--model", ORIENT_MODEL,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              timeout=ORIENT_TIMEOUT, cwd=cwd)
        if proc.returncode != 0:
            return 0
        text = json.loads(proc.stdout).get("result", "") or ""
    except (subprocess.TimeoutExpired, FileNotFoundError, json.JSONDecodeError):
        return 0
    for a in ("270", "180", "90", "0"):
        if a in text:
            return int(a)
    return 0


def _upright_images(pdf_bytes: bytes, claude_bin: str):
    """PDFを各ページ正立PIL画像のリストにして返す。失敗時は空リスト。

    returns: (list[PIL.Image], list[int] 検出角度)
    """
    try:
        import fitz
        from PIL import Image
    except ImportError:
        return [], []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return [], []

    images, angles = [], []
    with tempfile.TemporaryDirectory(prefix="orient_") as td:
        for i, page in enumerate(doc):
            if i >= MAX_PAGES:
                break
            pix = page.get_pixmap(dpi=RENDER_DPI)
            im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            ang = 0
            if claude_bin:
                th = im.copy()
                th.thumbnail((THUMB_MAX, THUMB_MAX))
                thumb_name = f"thumb_{i + 1}.png"
                th.save(os.path.join(td, thumb_name))
                ang = _detect_angle(thumb_name, td, claude_bin)
            up = im.rotate(-ang, expand=True) if ang else im  # 時計回りang → PILは-ang
            images.append(up)
            angles.append(ang)
    doc.close()
    return images, angles


def upright_page_images(pdf_bytes: bytes, out_dir: str):
    """正立補正したページ画像を out_dir に PNG 保存し、ファイル名リストを返す。

    画像でAIに渡すアプリ向け。失敗時は空リスト。
    """
    claude_bin = _claude_bin()
    images, _ = _upright_images(pdf_bytes, claude_bin)
    names = []
    for i, im in enumerate(images):
        fn = f"page_{i + 1}.png"
        im.save(os.path.join(out_dir, fn))
        names.append(fn)
    return names


def ensure_upright_image(image_bytes: bytes, force: bool = True):
    """1枚の画像の向きを補正したバイト列(PNG)を返す。失敗時は元のまま。

    写真・スキャン画像を Claude/AI のビジョンに渡すアプリ向けのドロップイン。
    """
    if not image_bytes:
        return image_bytes
    claude_bin = _claude_bin()
    if not claude_bin:
        return image_bytes
    try:
        from PIL import Image
    except ImportError:
        return image_bytes
    try:
        im = Image.open(io.BytesIO(image_bytes))
        im = im.convert("RGB")
    except Exception:
        return image_bytes
    with tempfile.TemporaryDirectory(prefix="orient_img_") as td:
        th = im.copy()
        th.thumbnail((THUMB_MAX, THUMB_MAX))
        thumb_name = "thumb.png"
        th.save(os.path.join(td, thumb_name))
        ang = _detect_angle(thumb_name, td, claude_bin)
    if not ang:
        return image_bytes
    try:
        up = im.rotate(-ang, expand=True)
        buf = io.BytesIO()
        up.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return image_bytes


def ensure_upright_bytes(data: bytes, filename: str = "", force_pdf: bool = False):
    """拡張子で PDF / 画像を判別し、向き補正したバイト列を返す。

    PDF -> ensure_upright_pdf, 画像 -> ensure_upright_image。判別不能時は PDF 扱い。
    """
    ext = os.path.splitext(filename or "")[1].lower()
    if ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"):
        return ensure_upright_image(data)
    return ensure_upright_pdf(data, force=force_pdf)


def ensure_upright_pdf(pdf_bytes: bytes, force: bool = False):
    """向きを補正した PDF バイト列を返す（テキストPDFや補正不能時は元のまま）。

    force=False のときテキスト層が十分にあるPDFはそのまま返す（補正不要）。
    PDFをClaudeに直接読ませるアプリ向けのドロップイン。
    """
    if not pdf_bytes:
        return pdf_bytes
    if not force:
        nchars = _pdf_text_chars(pdf_bytes)
        if nchars >= TEXT_MIN_CHARS:
            return pdf_bytes  # テキストPDFは向き関係なし
    claude_bin = _claude_bin()
    if not claude_bin:
        return pdf_bytes
    try:
        import fitz
    except ImportError:
        return pdf_bytes

    images, angles = _upright_images(pdf_bytes, claude_bin)
    if not images or not any(angles):
        # 補正不要（全ページ0度）or 失敗 → 元のまま返す（無駄な再エンコードを避ける）
        return pdf_bytes
    try:
        out = fitz.open()
        for im in images:
            buf = io.BytesIO()
            im.save(buf, format="JPEG", quality=88)  # サイズ抑制（スキャン書類に十分な画質）
            data = buf.getvalue()
            rect = fitz.Rect(0, 0, im.width, im.height)
            page = out.new_page(width=im.width, height=im.height)
            page.insert_image(rect, stream=data)
        result = out.tobytes()
        out.close()
        return result
    except Exception:
        return pdf_bytes
