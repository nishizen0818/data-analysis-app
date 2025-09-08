import streamlit as st
import pandas as pd
import re
import os
import json

# ファイル保存用ディレクトリと状態ファイルの設定
SAVE_DIR = "uploaded_files"
STATE_FILE = os.path.join(SAVE_DIR, "item_state.json")
os.makedirs(SAVE_DIR, exist_ok=True)

# ---------------------------- 状態ファイル管理 ----------------------------
def load_state():
    """状態ファイルを読み込み、存在しない場合は初期状態を返します。"""
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"class_file": None, "data_file": None}

def save_state(state):
    """状態ファイルを保存します。"""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

# ---------------------------- ヘルパー関数 ----------------------------

def read_uploaded_file(filepath):
    """
    保存されたExcelファイルを読み込み、シート名をキー、DataFrameを値とする辞書を返します。
    ファイルパスがNoneの場合は空の辞書を返します。
    """
    if filepath is not None and os.path.exists(filepath):
        return pd.read_excel(filepath, sheet_name=None, header=None)
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

# ---------------------------- Streamlit アプリ ----------------------------

st.set_page_config(page_title="商品分類別売上集計", layout="wide")
st.title("📊 アイテム別集計システム")

# 状態のロード
if "state" not in st.session_state:
    st.session_state.state = load_state()

# 左側にファイル状況、右側にアップロードUIを配置
left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("📂 現在のファイル状況")
    file_status_html = "<div style='background-color:#f9f5e9; padding:15px; border-radius:8px; border:1px solid #ddd;'>"
    
    status_map = {
        "class_file": "分類わけファイル",
        "data_file": "商品データファイル"
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
    st.header("① ファイルアップロード")
    class_file_uploader = st.file_uploader("🔼 分類わけファイル (.xlsx)", type=["xlsx", "xls"], key="class_file_uploader")
    data_file_uploader = st.file_uploader("🔼 商品データファイル (.xlsx)", type=["xlsx", "xls"], key="data_file_uploader")

    # アップロードされたファイルを保存
    save_file_and_update_state(class_file_uploader, "class_file")
    save_file_and_update_state(data_file_uploader, "data_file")

# 分析実行ボタン
if st.button("🚀 集計実行"):
    class_file_path = st.session_state.state.get("class_file", {}).get("path")
    data_file_path = st.session_state.state.get("data_file", {}).get("path")

    if class_file_path and data_file_path:
        try:
            # --- ② 分類ファイル読み込み ---
            df_class = pd.read_excel(class_file_path)
            df_class['優先フラグ'] = df_class['優先度'].fillna('').apply(lambda x: 1 if str(x).strip() == '〇' else 0)
            df_class['キーワード長'] = df_class['キーワード'].astype(str).apply(
                lambda x: sum(len(k.strip()) for k in str(x).split('・')) if pd.notna(x) else 0
            )
            df_class = df_class.sort_values(['優先フラグ', 'キーワード長'], ascending=[False, False])
            st.success("✅ 分類わけファイル読み込み完了")

            # --- ③ 商品データ読み込み ---
            df_data = pd.read_excel(data_file_path, header=0)
            st.success("✅ 商品データファイル読み込み完了")

            # --- ④ 商品名列検出と分類処理 ---
            product_cols = [col for col in df_data.columns if '商品' in str(col)]
            if product_cols:
                product_col = product_cols[0]
                df_data['商品名'] = df_data[product_col]
            else:
                st.error("❌ 『商品名』を含む列が見つかりません。")
                st.stop()

            def classify(name):
                if pd.isna(name):
                    return '未分類'
                for _, row in df_class.iterrows():
                    keywords = str(row['キーワード']).split('・')
                    if any(k.strip() in str(name) for k in keywords):
                        return row['分類']
                return '未分類'

            df_data['分類'] = df_data['商品名'].apply(classify)

            # --- 分類済みデータの表示 ---
            st.header("② 分類済みデータのプレビュー")
            preview_cols = ['商品名', '分類'] + [col for col in df_data.columns if '個数' in str(col) or '金額' in str(col)]
            preview_cols = [col for col in preview_cols if col in df_data.columns]

            if not df_data.empty and preview_cols:
                st.dataframe(df_data[preview_cols], use_container_width=True, key="classified_data_preview")
            else:
                st.info("分類後のプレビューデータがありません。")

            # --- ⑤ 年・個数・金額ペア抽出 ---
            records = []
            for col in df_data.columns:
                match = re.match(r'(\d{4})年\d+月_個数', col)
                if match:
                    year = int(match.group(1))
                    amt_col = col.replace('個数', '金額')
                    if amt_col in df_data.columns:
                        temp = df_data[['分類', col, amt_col]].copy()
                        temp.columns = ['分類', '個数', '金額']
                        temp['個数'] = pd.to_numeric(temp['個数'], errors='coerce').fillna(0)
                        temp['金額'] = pd.to_numeric(temp['金額'], errors='coerce').fillna(0)
                        temp['年'] = year
                        records.append(temp)

            if not records:
                st.error("❌ 年別の個数・金額列が見つかりませんでした。")
                st.stop()

            # --- ⑥ 集計と前年比 ---
            df_all = pd.concat(records)
            df_all = df_all.dropna(subset=['分類']).groupby(['分類', '年']).sum(numeric_only=True).reset_index()

            if df_all.empty:
                st.info("集計するデータがありません。")
                st.stop()

            df_all['前年金額'] = df_all.groupby('分類')['金額'].shift(1)
            df_all['金額_前年比'] = df_all.apply(
                lambda row: f"{(row['金額'] / row['前年金額'] * 100):.1f}%"
                if pd.notnull(row['前年金額']) and row['前年金額'] != 0 else
                (f"{100.0:.1f}%" if row['金額'] != 0 else "0.0%"),
                axis=1
            )
            df_all.drop(columns=['前年金額'], inplace=True)

            # --- ⑦ ピボット展開 ---
            def pivotify(df, column):
                p = df.pivot(index='分類', columns='年', values=column)
                p.columns = [f"{y}年_{column}" for y in p.columns]
                return p

            df_result = pd.concat([
                pivotify(df_all, '個数'),
                pivotify(df_all, '金額'),
                pivotify(df_all, '金額_前年比')
            ], axis=1).reset_index()

            # --- ⑧ 欠損値補完 ---
            for col in df_result.columns:
                if col.endswith('前年比'):
                    df_result[col] = df_result[col].replace('', '100.0%')
                else:
                    df_result[col] = df_result[col].fillna(0)

            # --- ⑨ 列順整列 ---
            all_years = sorted(df_all['年'].unique(), reverse=True)
            col_order = ['分類']
            for y in all_years:
                col_order += [f"{y}年_個数", f"{y}年_金額", f"{y}年_金額_前年比"]
            df_result = df_result[[col for col in col_order if col in df_result.columns]]

            # --- ⑩ 集計結果の表示 ---
            st.header("③ 集計結果プレビュー")
            if not df_result.empty:
                st.dataframe(df_result, use_container_width=True, key="final_summary_dataframe")
            else:
                st.info("集計結果が生成されませんでした。データを確認してください。")

        except Exception as e:
            st.error(f"⚠️ エラーが発生しました：\n\n{e}")
    else:
        st.info("📂 分類ファイルとデータファイルの両方をアップロードしてください。")
