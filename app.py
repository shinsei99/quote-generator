import datetime
import math
from pathlib import Path

import pandas as pd
import streamlit as st

import template_config as cfg
from parsing import extract_items_from_excel, extract_items_from_pdf
from template_builder import ensure_template_exists, fill_template

st.set_page_config(page_title="見積書自動生成ツール", page_icon="🧾", layout="wide")

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / cfg.TEMPLATE_PATH
ensure_template_exists(TEMPLATE_PATH)

DATA_DIR = BASE_DIR / "data"
ISSUERS_CSV = DATA_DIR / "issuers.csv"
DATA_DIR.mkdir(exist_ok=True)

ISSUER_FIELDS = ["name", "address", "tel", "fax", "registration_no", "license"]
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

if "items_df" not in st.session_state:
    st.session_state.items_df = pd.DataFrame(columns=["品名", "数量", "単位", "元単価"])
if "last_uploaded_name" not in st.session_state:
    st.session_state.last_uploaded_name = None

uploaded = st.file_uploader("見積書をアップロード（PDF または Excel）", type=["pdf", "xlsx"])

if uploaded is not None and uploaded.name != st.session_state.last_uploaded_name:
    file_bytes = uploaded.read()
    if uploaded.name.lower().endswith(".pdf"):
        with st.spinner("PDFを解析しています…（スキャン画像の場合はOCRのため時間がかかることがあります）"):
            items, method = extract_items_from_pdf(file_bytes)
    else:
        with st.spinner("Excelを解析しています…"):
            items = extract_items_from_excel(file_bytes)
            method = "Excel読み込み"

    if items:
        df = pd.DataFrame(items)
        if "単位" not in df.columns:
            df["単位"] = ""
        st.session_state.items_df = df[["品名", "数量", "単位", "元単価"]]
        st.success(f"{len(items)} 件の品目を読み取りました（方式: {method}）。内容を確認・修正してください。")
    else:
        st.warning("品目を自動検出できませんでした。下の表に手動で入力してください。")
    st.session_state.last_uploaded_name = uploaded.name

st.subheader("📋 品目データ（編集可能）")
st.caption("自動読み取りに誤りがある場合は、表のセルを直接クリックして修正・行の追加/削除ができます")

if st.button("＋ 行を追加"):
    new_row = pd.DataFrame([{"品名": "", "数量": 1, "単位": "", "元単価": 0}])
    st.session_state.items_df = pd.concat([st.session_state.items_df, new_row], ignore_index=True)

edited_df = st.data_editor(
    st.session_state.items_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "品名": st.column_config.TextColumn("品名", width="large"),
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

calc_df = edited_df.copy()
calc_df["数量"] = pd.to_numeric(calc_df["数量"], errors="coerce").fillna(0)
calc_df["単位"] = calc_df["単位"].fillna("") if "単位" in calc_df.columns else ""
calc_df["元単価"] = pd.to_numeric(calc_df["元単価"], errors="coerce").fillna(0)
# 上乗せ後単価・金額は消費税を含まない（税別）。小数点以下は切り捨て。
calc_df["上乗せ後単価"] = (calc_df["元単価"] * (1 + markup_rate / 100)).apply(math.floor)
calc_df["金額"] = (calc_df["数量"] * calc_df["上乗せ後単価"]).apply(math.floor)
calc_df = calc_df[calc_df["品名"].astype(str).str.strip() != ""].reset_index(drop=True)

st.subheader("💰 上乗せ後の計算結果（税別）")
if calc_df.empty:
    st.info("品目を入力すると、上乗せ後の金額がここに表示されます")
    subtotal = tax_amount = grand_total = 0
else:
    st.dataframe(
        calc_df[["品名", "数量", "単位", "元単価", "上乗せ後単価", "金額"]],
        use_container_width=True,
    )
    subtotal = int(calc_df["金額"].sum())
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
c1, c2 = st.columns(2)
with c1:
    client = st.text_input("宛先（会社名・氏名）")
with c2:
    issue_date = st.date_input("発行日", value=datetime.date.today())

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

if st.button("💾 この発行元情報を保存"):
    if not issuer_name.strip():
        st.error("会社名を入力してください")
    else:
        save_issuer({
            "name": issuer_name.strip(), "address": issuer_address, "tel": issuer_tel,
            "fax": issuer_fax, "registration_no": issuer_reg_no, "license": issuer_license,
        })
        st.success(f"「{issuer_name}」の発行元情報を保存しました。次回から会社名で呼び出せます。")
        st.rerun()

st.divider()

if calc_df.empty:
    st.button("📥 編集済み見積書をダウンロード", disabled=True, use_container_width=True)
else:
    items_payload = calc_df[["品名", "数量", "単位", "上乗せ後単価"]].to_dict("records")
    header_info = {
        "client": client,
        "issue_date": issue_date if issue_date else None,
    }
    issuer_info = {
        "name": issuer_name,
        "address": issuer_address,
        "tel": issuer_tel,
        "fax": issuer_fax,
        "registration_no": issuer_reg_no,
        "license": issuer_license,
    }
    output_bytes = fill_template(items_payload, header_info, issuer_info, tax_rate, TEMPLATE_PATH)
    st.download_button(
        "📥 編集済み見積書をダウンロード",
        data=output_bytes,
        file_name=f"見積書_{datetime.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
