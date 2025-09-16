import streamlit as st
import pandas as pd
import os
import json

# ファイル保存用ディレクトリと状態ファイルの設定
SAVE_DIR = "uploaded_files"
STATE_FILE = os.path.join(SAVE_DIR, "state.json")
os.makedirs(SAVE_DIR, exist_ok=True)

# ---------------------------- 状態ファイル管理 ----------------------------
def load_state():
    """状態ファイルを読み込み、存在しない場合は初期状態を返します。"""
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"prev_file": None, "curr_file": None, "helper_file": None}

def save_state(state):
    """状態ファイルを保存します。"""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

# ---------------------------- ヘルパー関数 ----------------------------

def read_uploaded_file(uploaded_file):
    """
    アップロードされたExcelファイルを読み込み、シート名をキー、DataFrameを値とする辞書を返します。
    ファイルがNoneの場合は空の辞書を返します。
    """
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file, sheet_name=None, header=None)
    return {}

def save_file_and_update_state(uploaded_file, file_key):
    """
    アップロードされたファイルをローカルに保存し、状態を更新します。
    """
    if uploaded_file:
        filepath = os.path.join(SAVE_DIR, uploaded_file.name)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.session_state.state[file_key] = {"path": filepath, "name": uploaded_file.name}
        save_state(st.session_state.state)

def extract_mapping(helper_sheets):
    """
    補助データシートから、除外コード、売上修正マップ、カテゴリマップを抽出します。
    """
    exclude_codes = []
    if "削除依頼" in helper_sheets:
        codes = helper_sheets["削除依頼"].iloc[:, 0].dropna()
        codes = codes.astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(4)
        exclude_codes = codes.tolist()

    fix_sales_map = {}
    if "計算修正" in helper_sheets:
        sheet = helper_sheets["計算修正"].iloc[:, :2].dropna(how="all")
        for _, row in sheet.iterrows():
            try:
                code = str(int(row[0])).zfill(4)
                factor = float(row[1])
                fix_sales_map[code] = factor
            except ValueError:
                st.warning(f"「計算修正」シートのデータ形式が不正です: {row.tolist()}")
                continue

    category_map = {}
    if "取引先リスト" in helper_sheets:
        sheet = helper_sheets["取引先リスト"].iloc[:, [0, 2]].dropna(how="all")
        for _, row in sheet.iterrows():
            try:
                code = str(int(row[0])).zfill(4)
                category = str(row[2]).strip()
                category_map[code] = category
            except ValueError:
                st.warning(f"「取引先リスト」シートのデータ形式が不正です: {row.tolist()}")
                continue

    return exclude_codes, fix_sales_map, category_map

def clean_sheet(df, exclude_codes, fix_sales_map, category_map):
    """
    アップロードされた売上データをクリーニングし、必要な列を整形します。
    """
    header_idx = df[df.apply(lambda r: r.astype(str).str.contains("得意先コード", na=False)).any(axis=1)].index
    if len(header_idx) == 0:
        return pd.DataFrame()
    header = header_idx[0]

    df.columns = df.iloc[header]
    df = df[(header + 1):].reset_index(drop=True)

    required_columns = {"得意先コード", "得意先名", "純売上額"}
    if not required_columns.issubset(df.columns):
        return pd.DataFrame()

    df["得意先コード"] = df["得意先コード"].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(4)
    df = df[~df["得意先コード"].isin(exclude_codes)]

    df["純売上額"] = df.apply(
        lambda r: r["純売上額"] * fix_sales_map.get(r["得意先コード"], 1.0),
        axis=1
    )
    df["大分類"] = df["得意先コード"].map(category_map).fillna("未分類")

    total_sales = df["純売上額"].sum()
    df["構成比"] = (df["純売上額"] / total_sales * 100).round(2) if total_sales != 0 else 0.0

    grouped = (
        df.groupby(["得意先コード", "得意先名", "大分類"], as_index=False)
        .agg({"純売上額": "sum", "構成比": "sum"})
        .sort_values("純売上額", ascending=False)
    )

    return grouped

def compare_years(prev_df, curr_df):
    """
    前年データと今年データを比較し、差額と前年比を計算します。
    """
    merged = pd.merge(
        prev_df,
        curr_df,
        on=["得意先コード", "得意先名", "大分類"],
        how="outer",
        suffixes=("_前年", "_今年"),
    )

    for col in ["純売上額_前年", "純売上額_今年", "構成比_前年", "構成比_今年"]:
        merged[col] = merged[col].fillna(0)

    merged["純売上額_前年"] = (merged["純売上額_前年"] / 1000).round().astype("Int64")
    merged["純売上額_今年"] = (merged["純売上額_今年"] / 1000).round().astype("Int64")

    merged["差額"] = merged["純売上額_今年"] - merged["純売上額_前年"]
    merged["前年比(%)"] = merged.apply(
        lambda row: round(row["純売上額_今年"] / row["純売上額_前年"] * 100, 1)
        if row["純売上額_前年"] != 0 else (100.0 if row["純売上額_今年"] != 0 else 0.0),
        axis=1
    )

    ordered_cols = [
        "得意先コード", "得意先名", "大分類",
        "純売上額_今年", "構成比_今年",
        "純売上額_前年", "構成比_前年",
        "前年比(%)", "差額"
    ]
    return merged[ordered_cols]

def summarize_by_category(comp_df):
    """
    カテゴリ別に売上データを集計します。
    """
    cat = comp_df.groupby("大分類", as_index=False).agg({
        "純売上額_前年": "sum",
        "純売上額_今年": "sum",
        "差額": "sum"
    })
    cat["前年比(%)"] = cat.apply(
        lambda r: round(r["純売上額_今年"] / r["純売上額_前年"] * 100, 1)
        if r["純売上額_前年"] != 0 else (100.0 if r["純売上額_今年"] != 0 else 0.0),
        axis=1
    )
    return cat

# ---------------------------- Streamlit アプリ ----------------------------

st.set_page_config(page_title="卸営業数値分析システム", layout="wide")
st.title("📊 卸営業数値分析システム")

# 状態のロード
if "state" not in st.session_state:
    st.session_state.state = load_state()

# 左側にファイル状況、右側にアップロードUIを配置
left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("📂 現在のファイル状況")
    file_status_html = "<div style='background-color:#f9f5e9; padding:15px; border-radius:8px; border:1px solid #ddd;'>"
    
    status_map = {
        "prev_file": "前年データ",
        "curr_file": "今年データ",
        "helper_file": "補助データ"
    }
    
    for key, label in status_map.items():
        info = st.session_state.state.get(key)
        if info and "name" in info:
            file_status_html += f"<p><strong>{label}</strong>: ✅ {info['name']}</p>"
        else:
            file_status_html += f"<p><strong>{label}</strong>: ❌ 未設定</p>"
    file_status_html += "</div>"
    st.markdown(file_status_html, unsafe_allow_html=True)
    
with right_col:
    st.header("📁 ファイルアップロード")
    prev_file_uploader = st.file_uploader("前年データ (Excel)", type=["xlsx"], key="prev_file_uploader")
    curr_file_uploader = st.file_uploader("今年データ (Excel)", type=["xlsx"], key="curr_file_uploader")
    helper_file_uploader = st.file_uploader("補助データ (データ整理.xlsx)", type=["xlsx"], key="helper_file_uploader")

    # アップロードされたファイルを保存
    save_file_and_update_state(prev_file_uploader, "prev_file")
    save_file_and_update_state(curr_file_uploader, "curr_file")
    save_file_and_update_state(helper_file_uploader, "helper_file")

# 🚀 分析実行ボタンを押した際の処理
if st.button("🚀 分析実行"):
    # 状態ファイルからパスを取得
    prev_file_path = st.session_state.state.get("prev_file", {}).get("path")
    curr_file_path = st.session_state.state.get("curr_file", {}).get("path")
    helper_file_path = st.session_state.state.get("helper_file", {}).get("path")
    
    if prev_file_path and curr_file_path and helper_file_path:
        try:
            prev_sheets = pd.read_excel(prev_file_path, sheet_name=None, header=None)
            curr_sheets = pd.read_excel(curr_file_path, sheet_name=None, header=None)
            helper_sheets = pd.read_excel(helper_file_path, sheet_name=None, header=None)
            
            exclude_codes, fix_sales_map, category_map = extract_mapping(helper_sheets)

            # データフレームのクリーニング
            if not prev_sheets:
                st.error("前年データファイルにシートが見つかりません。")
                st.stop()
            prev_sheet_df = list(prev_sheets.values())[0]

            if not curr_sheets:
                st.error("今年データファイルにシートが見つかりません。")
                st.stop()
            curr_sheet_df = list(curr_sheets.values())[0]

            prev_clean = clean_sheet(prev_sheet_df, exclude_codes, fix_sales_map, category_map)
            curr_clean = clean_sheet(curr_sheet_df, exclude_codes, fix_sales_map, category_map)

            if prev_clean.empty or curr_clean.empty:
                st.error("ヘッダ行（得意先コードなど）が見つからない、または必須列（得意先コード、得意先名、純売上額）が不足しています。Excelの列構成をご確認ください。")
                st.stop()
            
            # --- 修正: 計算結果をセッションステートに保存 ---
            st.session_state.prev_clean = prev_clean
            st.session_state.curr_clean = curr_clean
            st.session_state.comp_df = compare_years(prev_clean, curr_clean)
            st.session_state.summary_df = summarize_by_category(st.session_state.comp_df)
            st.success("分析完了！")
            st.rerun() # 計算が完了したら、ページを再実行して結果を表示

        except Exception as e:
            st.error(f"分析中にエラーが発生しました。ファイルの内容を確認してください。エラー詳細: {e}")
    else:
        st.info("すべてのファイルがアップロードされていません。ファイル状況をご確認ください。")


# --- 修正: セッションステートにデータがある場合にのみ結果を表示 ---
if "comp_df" in st.session_state and "summary_df" in st.session_state:
    
    st.markdown("---") # 視覚的な区切り線
    
    st.markdown("### Step 1: 整理後データ")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("前年整理データ")
        st.dataframe(st.session_state.prev_clean, use_container_width=True)
    with col2:
        st.subheader("今年整理データ")
        st.dataframe(st.session_state.curr_clean, use_container_width=True)

    st.markdown("---")
    
    st.markdown("### Step 2: 前年 vs 今年 比較（千円単位）")
    st.dataframe(st.session_state.comp_df, use_container_width=True)
    
    st.markdown("---")

    st.markdown("### Step 3: 並び替えと集計")
    
    # オプション名を全角アンダースコアから半角アンダースコアに修正
    sort_options = (
        "大分類別_純売上額_今年順",
        "大分類別_差額ベスト順",
        "大分類別_差額ワースト順",
        "得意先別_純売上額_今年順",
        "得意先別_差額ベスト順",
        "得意先別_差額ワースト順",
    )
    option = st.selectbox(
        "並び替え基準を選んでください",
        sort_options,
        key="sort_option_select"
    )

    # 選択されたオプションに基づいて処理を分岐
    if option.startswith("大分類別"):
        summary_df = st.session_state.summary_df.copy() # オリジナルを保持するためコピー
        if "_純売上額_" in option:
            summary_sorted = summary_df.sort_values("純売上額_今年", ascending=False)
        elif "ベスト" in option:
            summary_sorted = summary_df.sort_values("差額", ascending=False)
        else: # ワースト
            summary_sorted = summary_df.sort_values("差額", ascending=True)
        
        st.subheader("大分類別：集計結果")
        if not summary_sorted.empty:
            st.dataframe(summary_sorted, use_container_width=True)
            st.bar_chart(summary_sorted.set_index("大分類")["純売上額_今年"])
        else:
            st.info("集計するデータがありません。")
    else: # 得意先別
        comp_df = st.session_state.comp_df.copy() # オリジナルを保持するためコピー
        if "_純売上額_" in option:
            df_sorted = comp_df.sort_values("純売上額_今年", ascending=False)
        elif "ベスト" in option:
            df_sorted = comp_df.sort_values("差額", ascending=False)
        else: # ワースト
            df_sorted = comp_df.sort_values("差額", ascending=True)
        
        st.subheader("得意先別：比較結果")
        if not df_sorted.empty:
            st.dataframe(df_sorted, use_container_width=True)
        else:
            st.info("集計するデータがありません。")
