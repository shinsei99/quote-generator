import datetime
import math
from pathlib import Path

import pandas as pd
import streamlit as st

import template_config as cfg
from parsing import ClaudeExtractionError, extract_items_from_excel, extract_items_from_pdf
from template_builder import ensure_template_exists, fill_combined_document

st.set_page_config(page_title="見積書自動生成ツール", page_icon="🧾", layout="wide")

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / cfg.TEMPLATE_PATH
ensure_template_exists(TEMPLATE_PATH)

DATA_DIR = BASE_DIR / "data"
ISSUERS_CSV = DATA_DIR / "issuers.csv"
DATA_DIR.mkdir(exist_ok=True)

ISSUER_FIELDS = ["name", "address", "tel", "fax", "registration_no", "license", "bank_info"]
EMPTY_ISSUER = {f: "" for f in ISSUER_FIELDS}


def load_issuers() -> pd.DataFrame:
    if ISSUERS_CSV.exists():
        return pd.read_csv(ISSUERS_CSV, dtype=str).fillna("")
    return pd.DataFrame(columns=ISSUER_FIELDS)


def save_issuer(data: dict):
    df = load_issuers()
    df = df[df["name"] != data["name"]]
    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_csv(ISSUERS_CSV, index=False)


st.markdown(
    """
    <style>
    .stApp { background-color: #f4f8fb; }
    .info-card {
        background: white; border-radius: 14px; padding: 16px 22px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 14px;
    }
    .total-box {
        background: linear-gradient(90deg,#0f172a,#1e293b); color: white;
        padding: 18px 24px; border-radius: 12px; font-size: 1.4rem; font-weight: 700;
        margin: 12px 0;
    }
    h1, h2, h3 { color: #0f172a; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 見積書自動生成ツール")
st.caption("取引先から受け取ったPDF/Excelの見積書を読み取り、上乗せ率を反映して自社テンプレートに転記します")

ITEM_COLUMNS = ["no", "種別", "品名", "規格", "備考", "数量", "単位", "元単価"]

if "items_df" not in st.session_state:
    st.session_state.items_df = pd.DataFrame(columns=ITEM_COLUMNS)
if "last_uploaded_name" not in st.session_state:
    st.session_state.last_uploaded_name = None
if "notes_text" not in st.session_state:
    st.session_state.notes_text = ""
if "pdf_nos" not in st.session_state:
    st.session_state.pdf_nos = []

uploaded = st.file_uploader("見積書をアップロード（PDF または Excel）", type=["pdf", "xlsx"])

if uploaded is not None and uploaded.name != st.session_state.last_uploaded_name:
    file_bytes = uploaded.read()
    items, method = [], ""
    if uploaded.name.lower().endswith(".pdf"):
        with st.spinner("Claude AIが見積書を解析しています…（通常1〜3分、APIが混雑時は最大10分かかります）"):
            try:
                items, method, notes = extract_items_from_pdf(file_bytes)
                st.session_state.notes_text = notes
            except ClaudeExtractionError as e:
                st.error(f"PDFの解析に失敗しました: {e}")
    else:
        with st.spinner("Excelを解析しています…"):
            items = extract_items_from_excel(file_bytes)
            method = "Excel読み込み"
            st.session_state.notes_text = ""

    if items:
        df = pd.DataFrame(items)
        for col in ITEM_COLUMNS:
            if col not in df.columns:
                if col == "種別":
                    df[col] = "品目"
                elif col in ("数量", "元単価"):
                    df[col] = 0
                else:
                    df[col] = ""
        df["no"] = df["no"].fillna("").astype(str).str.strip()
        df["品名"] = df["品名"].fillna("")
        df["規格"] = df["規格"].fillna("")
        df["備考"] = df["備考"].fillna("")
        df["単位"] = df["単位"].fillna("")
        df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0)
        df["元単価"] = pd.to_numeric(df["元単価"], errors="coerce").fillna(0)
        # PDF由来の番号リストを保存（チェック機能用）
        if uploaded.name.lower().endswith(".pdf"):
            st.session_state.pdf_nos = [str(n) for n in df["no"].tolist() if str(n).strip()]

        # 値引行は表には出さず、別途「値引きを反映する」設定で扱う
        is_discount = (df["種別"] == "品目") & (df["品名"] == "値引")
        detected_discount = float(-df.loc[is_discount, "元単価"].sum()) if is_discount.any() else 0.0
        st.session_state.detected_discount = detected_discount
        df = df[~is_discount]

        st.session_state.items_df = df[ITEM_COLUMNS]
        n_daikomi = (df["種別"] == "大項目").sum()
        n_sections = (df["種別"] == "工事区分").sum()
        n_breaks = (df["種別"] == "小計区切り").sum()
        note_parts = []
        if n_daikomi:
            note_parts.append(f"大項目 {n_daikomi} 件")
        if n_sections:
            note_parts.append(f"工事区分の見出し {n_sections} 件")
        if n_breaks:
            note_parts.append(f"小計区切り {n_breaks} 件")
        if detected_discount:
            note_parts.append(f"値引き {detected_discount:,.0f}円を検出（反映する場合は下のチェックを入れてください）")
        note = f"（{'、'.join(note_parts)}）" if note_parts else ""
        st.success(f"{len(items)} 件読み取りました（方式: {method}）{note}。内容を確認・修正してください。")
    else:
        st.warning("品目を自動検出できませんでした。下の表に手動で入力してください。")
    st.session_state.last_uploaded_name = uploaded.name

if "detected_discount" not in st.session_state:
    st.session_state.detected_discount = 0.0

st.subheader("📋 品目データ（編集可能）")
st.caption(
    "自動読み取りに誤りがある場合は、表のセルを直接クリックして修正・行の追加/削除ができます。"
    "「種別」が工事区分の行は品名だけの見出しとして、小計区切りの行はそこまでの品目で"
    "小計を作る区切りとして出力されます（元の見積書が工事箇所ごとに分かれている場合、"
    "できるだけ同じ区切り・見出しで出力されます）。"
)

bc1, bc2, bc3, bc4 = st.columns(4)
with bc1:
    if st.button("＋ 品目行を追加"):
        new_row = pd.DataFrame([{"no": "", "種別": "品目", "品名": "", "規格": "", "備考": "", "数量": 1, "単位": "", "元単価": 0}])
        st.session_state.items_df = pd.concat([st.session_state.items_df, new_row], ignore_index=True)
with bc2:
    if st.button("＋ 大項目を追加"):
        new_row = pd.DataFrame([{"no": "", "種別": "大項目", "品名": "", "規格": "", "備考": "", "数量": 0, "単位": "", "元単価": 0}])
        st.session_state.items_df = pd.concat([st.session_state.items_df, new_row], ignore_index=True)
with bc3:
    if st.button("＋ 工事区分の見出しを追加"):
        new_row = pd.DataFrame([{"no": "", "種別": "工事区分", "品名": "", "規格": "", "備考": "", "数量": 0, "単位": "", "元単価": 0}])
        st.session_state.items_df = pd.concat([st.session_state.items_df, new_row], ignore_index=True)
with bc4:
    if st.button("＋ 小計区切りを追加"):
        new_row = pd.DataFrame([{"no": "", "種別": "小計区切り", "品名": "", "規格": "", "備考": "", "数量": 0, "単位": "", "元単価": 0}])
        st.session_state.items_df = pd.concat([st.session_state.items_df, new_row], ignore_index=True)

edited_df = st.data_editor(
    st.session_state.items_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "no": st.column_config.TextColumn("No.", width="small", help="元PDFの番号。変更すると出力Excelの番号もこの値になります"),
        "種別": st.column_config.SelectboxColumn("種別", options=["品目", "大項目", "工事区分", "小計区切り"], width="small"),
        "品名": st.column_config.TextColumn("品名", width="large"),
        "規格": st.column_config.TextColumn("規格・仕様", width="medium"),
        "備考": st.column_config.TextColumn("備考", width="medium"),
        "数量": st.column_config.NumberColumn("数量", min_value=0, step=1),
        "単位": st.column_config.TextColumn("単位", width="small"),
        "元単価": st.column_config.NumberColumn("元単価（円）", min_value=0, step=1),
    },
    key="items_editor",
)
st.session_state.items_df = edited_df

st.divider()

c_rate1, c_rate2 = st.columns(2)
with c_rate1:
    markup_rate = st.number_input(
        "上乗せ率（％）", min_value=0.0, max_value=500.0, value=10.0, step=0.5,
        help="例：10%なら元の単価を1.1倍にします（本文の単価・金額は消費税別で計算します）",
    )
with c_rate2:
    tax_rate = st.number_input(
        "消費税率（％）", min_value=0.0, max_value=100.0, value=10.0, step=0.5,
        help="小計に対してこの税率で消費税を計算し、合計の直前に表示します",
    )

dc1, dc2, dc3 = st.columns([1, 2, 2])
with dc1:
    apply_discount = st.checkbox(
        "値引きを反映する", value=False,
        help="デフォルトでは反映しません。元の見積書から値引きが検出された場合、ここにチェックを入れると小計から差し引きます（上乗せの対象外）。",
    )
with dc2:
    discount_amount = st.number_input(
        "値引き額（円）", min_value=0.0, value=st.session_state.detected_discount, step=100.0,
        disabled=not apply_discount,
    )
with dc3:
    skip_round10 = st.checkbox(
        "1円単位で計算する（切り捨てなし）", value=False,
        help="デフォルトでは上乗せ後単価の9円以下を切り捨て（10円単位）します。チェックすると1円単位のまま計算します。",
    )

calc_df = edited_df.copy()
calc_df["種別"] = calc_df["種別"].fillna("品目") if "種別" in calc_df.columns else "品目"
calc_df.loc[calc_df["種別"] == "", "種別"] = "品目"
calc_df["規格"] = calc_df["規格"].fillna("") if "規格" in calc_df.columns else ""
calc_df["備考"] = calc_df["備考"].fillna("") if "備考" in calc_df.columns else ""
calc_df["単位"] = calc_df["単位"].fillna("") if "単位" in calc_df.columns else ""
calc_df["数量"] = pd.to_numeric(calc_df["数量"], errors="coerce").fillna(0)
calc_df["元単価"] = pd.to_numeric(calc_df["元単価"], errors="coerce").fillna(0)
is_item = calc_df["種別"] == "品目"
calc_df["上乗せ後単価"] = 0
calc_df["金額"] = 0
if skip_round10:
    # 1円単位（切り捨てなし）
    calc_df.loc[is_item, "上乗せ後単価"] = (calc_df.loc[is_item, "元単価"] * (1 + markup_rate / 100)).apply(math.floor)
else:
    # デフォルト: 10円以下切り捨て（9円以下切り捨て）
    calc_df.loc[is_item, "上乗せ後単価"] = (calc_df.loc[is_item, "元単価"] * (1 + markup_rate / 100)).apply(
        lambda x: math.floor(x / 10) * 10
    )
calc_df.loc[is_item, "金額"] = (calc_df.loc[is_item, "数量"] * calc_df.loc[is_item, "上乗せ後単価"]).apply(math.floor)
calc_df = calc_df[
    (calc_df["品名"].astype(str).str.strip() != "") | (calc_df["種別"] == "小計区切り")
].reset_index(drop=True)

if apply_discount and discount_amount > 0:
    # 値引きは上乗せ率の対象外（元の値引き額をそのまま差し引く）
    discount_row = pd.DataFrame([{
        "種別": "品目", "品名": "値引", "規格": "", "備考": "", "数量": 1, "単位": "",
        "元単価": -discount_amount, "上乗せ後単価": -discount_amount, "金額": -discount_amount,
    }])
    calc_df = pd.concat([calc_df, discount_row], ignore_index=True)

st.subheader("💰 上乗せ後の計算結果（税別）")
if calc_df.empty:
    st.info("品目を入力すると、上乗せ後の金額がここに表示されます")
    subtotal = tax_amount = grand_total = 0
    result_df = calc_df
else:
    st.caption("「上乗せ後単価」列のセルをクリックすると手動で金額を修正できます。他の列は読み取り専用です。")
    result_df = st.data_editor(
        calc_df[["no", "種別", "品名", "規格", "備考", "数量", "単位", "元単価", "上乗せ後単価"]],
        use_container_width=True,
        disabled=["no", "種別", "品名", "規格", "備考", "数量", "単位", "元単価"],
        column_config={
            "no": st.column_config.TextColumn("No.", width="small"),
            "上乗せ後単価": st.column_config.NumberColumn("上乗せ後単価（手動修正可）", min_value=0, step=1),
        },
        key="calc_editor",
    )
    # 編集後の上乗せ後単価から金額を再計算
    result_df = result_df.copy()
    result_df["金額"] = 0
    is_item_r = result_df["種別"] == "品目"
    result_df.loc[is_item_r, "金額"] = (
        result_df.loc[is_item_r, "数量"] * result_df.loc[is_item_r, "上乗せ後単価"]
    ).apply(math.floor)
    subtotal = int(result_df.loc[result_df["種別"] == "品目", "金額"].sum())
    tax_amount = math.floor(subtotal * tax_rate / 100)
    grand_total = subtotal + tax_amount
    st.markdown(
        f"<div class='info-card'>"
        f"小計（税別）：¥{subtotal:,.0f}　／　"
        f"消費税（{tax_rate:g}%）：¥{tax_amount:,.0f}"
        f"</div>",
        unsafe_allow_html=True,
    )
    st.markdown(f"<div class='total-box'>合計金額（税込）：¥{grand_total:,.0f}</div>", unsafe_allow_html=True)

st.divider()
st.subheader("📄 宛先・発行日")
c1, c2, c3 = st.columns(3)
with c1:
    client = st.text_input("宛先（会社名・氏名）")
with c2:
    quote_issue_date = st.date_input("見積書 発行日", value=datetime.date.today())
with c3:
    invoice_issue_date = st.date_input("請求書 発行日", value=datetime.date.today())

st.subheader("🏢 発行元（自社）情報")

if "issuer_form_version" not in st.session_state:
    st.session_state.issuer_form_version = 0
if "issuer_prefill" not in st.session_state:
    st.session_state.issuer_prefill = dict(EMPTY_ISSUER)

issuers_df = load_issuers()
NEW_ISSUER_LABEL = "＋ 新規入力"
issuer_options = [NEW_ISSUER_LABEL] + issuers_df["name"].tolist()

selected_issuer = st.selectbox(
    "会社名から呼び出し（保存済みの発行元）", issuer_options, key="issuer_select",
)

if selected_issuer != NEW_ISSUER_LABEL:
    row = issuers_df[issuers_df["name"] == selected_issuer].iloc[0]
    candidate_prefill = {f: row.get(f, "") for f in ISSUER_FIELDS}
else:
    candidate_prefill = dict(EMPTY_ISSUER)

if candidate_prefill != st.session_state.issuer_prefill:
    st.session_state.issuer_prefill = candidate_prefill
    st.session_state.issuer_form_version += 1

v = st.session_state.issuer_form_version
p = st.session_state.issuer_prefill

c3, c4 = st.columns(2)
with c3:
    issuer_name = st.text_input("発行元 会社名", value=p["name"], key=f"issuer_name_{v}")
    issuer_address = st.text_input("発行元 住所", value=p["address"], key=f"issuer_address_{v}")
    issuer_license = st.text_input("許可番号など（任意）", value=p["license"], key=f"issuer_license_{v}")
with c4:
    issuer_tel = st.text_input("発行元 TEL", value=p["tel"], key=f"issuer_tel_{v}")
    issuer_fax = st.text_input("発行元 FAX", value=p["fax"], key=f"issuer_fax_{v}")
    issuer_reg_no = st.text_input("インボイス登録番号", value=p["registration_no"], key=f"issuer_reg_no_{v}")

issuer_bank_info = st.text_input(
    "振込先（請求書に記載・任意）", value=p["bank_info"], key=f"issuer_bank_info_{v}",
    placeholder="例：○○銀行　○○支店　普通　1234567　カ）○○商事",
)

if st.button("💾 この発行元情報を保存"):
    if not issuer_name.strip():
        st.error("会社名を入力してください")
    else:
        save_issuer({
            "name": issuer_name.strip(), "address": issuer_address, "tel": issuer_tel,
            "fax": issuer_fax, "registration_no": issuer_reg_no, "license": issuer_license,
            "bank_info": issuer_bank_info,
        })
        st.success(f"「{issuer_name}」の発行元情報を保存しました。次回から会社名で呼び出せます。")
        st.rerun()

st.divider()
st.subheader("📝 備考・特記事項")
st.caption("PDFから自動抽出された【別途見積】【見積条件】などの内容です。内容を確認・編集してください。Excel出力の末尾に追記されます。")
notes_text = st.text_area(
    "備考テキスト（Excel末尾に出力）",
    value=st.session_state.notes_text,
    height=160,
    placeholder="例：【別途見積】\nセキュリティ工事\n\n【見積条件】\n・補修等は追加になります。",
    label_visibility="collapsed",
)
st.session_state.notes_text = notes_text

st.divider()

show_section_subtotals = st.checkbox(
    "グループごとの小計を表示する",
    value=False,
    help="工事区分（グループ）ごとに【小計】行を入れます。元の見積書に小計がこまめに入っている場合にチェックしてください。",
)

if st.session_state.pdf_nos:
    with st.expander("🔍 元PDF番号チェック", expanded=False):
        pdf_nos_list = st.session_state.pdf_nos
        current_nos_list = [
            str(row.get("no", "")).strip()
            for _, row in edited_df.iterrows()
            if str(row.get("no", "")).strip()
        ]

        def _sort_key(x):
            return (0, int(x)) if x.isdigit() else (1, x)

        pdf_set = set(pdf_nos_list)
        current_set = set(current_nos_list)
        missing = sorted(pdf_set - current_set, key=_sort_key)
        added = sorted(current_set - pdf_set, key=_sort_key)

        pdf_counts = {}
        for n in pdf_nos_list:
            pdf_counts[n] = pdf_counts.get(n, 0) + 1
        current_counts = {}
        for n in current_nos_list:
            current_counts[n] = current_counts.get(n, 0) + 1
        dup_pdf = {n: c for n, c in pdf_counts.items() if c > 1}
        dup_current = {n: c for n, c in current_counts.items() if c > 1}

        col_a, col_b = st.columns(2)
        with col_a:
            sorted_pdf = sorted(pdf_nos_list, key=_sort_key)
            st.write(f"**元PDF番号** ({len(pdf_nos_list)}件): {', '.join(sorted_pdf)}")
        with col_b:
            sorted_cur = sorted(current_nos_list, key=_sort_key)
            st.write(f"**現在の番号** ({len(current_nos_list)}件): {', '.join(sorted_cur)}")

        if dup_pdf:
            dup_str = "、".join(f"{n}（{c}回）" for n, c in sorted(dup_pdf.items(), key=lambda kv: _sort_key(kv[0])))
            st.info(f"元PDFで番号が重複しています（PDF側の誤りの可能性）: {dup_str}")
        if missing:
            st.warning(f"元PDFにあるが現在にない番号: **{', '.join(missing)}**")
        if added:
            st.warning(f"元PDFにない番号が追加されています: **{', '.join(added)}**")
        if dup_current:
            dup_str = "、".join(f"{n}（{c}回）" for n, c in sorted(dup_current.items(), key=lambda kv: _sort_key(kv[0])))
            st.warning(f"現在の番号に重複があります: {dup_str}")
        if not missing and not added and not dup_current:
            st.success("✓ 番号は元PDFと完全に一致しています")

if calc_df.empty:
    st.button("📥 見積書・請求書をダウンロード（1冊にまとめて）", disabled=True, use_container_width=True)
else:
    items_payload = result_df[["no", "種別", "品名", "規格", "備考", "数量", "単位", "上乗せ後単価"]].to_dict("records")
    header_info = {
        "client": client,
        "issue_date": quote_issue_date if quote_issue_date else None,
    }
    invoice_header_info = {
        "client": client,
        "issue_date": invoice_issue_date if invoice_issue_date else None,
    }
    issuer_info = {
        "name": issuer_name,
        "address": issuer_address,
        "tel": issuer_tel,
        "fax": issuer_fax,
        "registration_no": issuer_reg_no,
        "license": issuer_license,
    }

    combined_bytes = fill_combined_document(
        items_payload, header_info, issuer_info, tax_rate, TEMPLATE_PATH,
        bank_info=issuer_bank_info,
        notes=notes_text,
        show_section_subtotals=show_section_subtotals,
        invoice_header_info=invoice_header_info,
    )
    st.download_button(
        "📥 見積書・請求書をダウンロード（1冊にまとめて）",
        data=combined_bytes,
        file_name=f"見積書_請求書_{datetime.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        help="「見積書」「請求書」の2シートが1冊のExcelファイルに入っています",
    )
