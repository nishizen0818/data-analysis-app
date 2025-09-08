import streamlit as st
import pandas as pd
import re
from collections import Counter
from datetime import datetime
import openpyxl # openpyxlをインポート
import os
import json

# ファイル保存用ディレクトリと状態ファイルの設定
SAVE_DIR = "uploaded_files"
STATE_FILE = os.path.join(SAVE_DIR, "report_state.json")
os.makedirs(SAVE_DIR, exist_ok=True)

# ---------------------------- 状態ファイル管理 ----------------------------
def load_state():
    """状態ファイルを読み込み、存在しない場合は初期状態を返します。"""
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"uploaded_file": None}

def save_state(state):
    """状態ファイルを保存します。"""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

# ---------------------------- ヘルパー関数 ----------------------------
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

# 定数
KINIKI_AREAS = ["大阪", "奈良", "京都", "滋賀", "兵庫", "三重", "和歌山"]
VALID_CATEGORIES = ["駅", "高速", "空港", "一般店", "量販店", "商社"]

# ページ設定
st.set_page_config(layout="wide")
st.title("📊 アタックリスト分析")

# 状態のロード
if "state" not in st.session_state:
    st.session_state.state = load_state()

# 左側にファイル状況、右側にアップロードUIを配置
left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("📂 現在のファイル状況")
    file_status_html = "<div style='background-color:#f9f5e9; padding:15px; border-radius:8px; border:1px solid #ddd;'>"
    
    info = st.session_state.state.get("uploaded_file")
    if info and "name" in info:
        file_status_html += f"<p><strong>アタックリストファイル</strong>: ✅ {info['name']}</p>"
    else:
        file_status_html += f"<p><strong>アタックリストファイル</strong>: ❌ 未設定</p>"
    file_status_html += "</div>"
    st.markdown(file_status_html, unsafe_allow_html=True)

with right_col:
    # ファイルアップローダー
    uploaded_file = st.file_uploader("Excelファイル（.xlsx）をアップロード", type="xlsx", key="main_file_uploader")
    save_file_and_update_state(uploaded_file, "uploaded_file")

    # セッションステートに変数を初期化
    if 'df_filtered_display' not in st.session_state:
        st.session_state.df_filtered_display = None
        
    file_path = st.session_state.state.get("uploaded_file", {}).get("path")
    
    if file_path and os.path.exists(file_path):
        try:
            # ファイルがアップロードされたら、フィルター画面を表示
            st.markdown("---")
            st.markdown("### 🎛 訪問データの絞り込み")

            # openpyxlでワークブックを読み込み、非表示シートを特定
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            visible_sheet_names = []
            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                if ws.sheet_state == 'visible':
                    visible_sheet_names.append(sheet_name)

            # pandas.ExcelFileオブジェクトを作成
            xls = pd.ExcelFile(file_path)

            # 表示されているシート名のみを対象とする
            sheet_names = [s for s in xls.sheet_names if s in visible_sheet_names]

            # シートの分離
            log_sheet = "操作履歴"
            main_sheets = [s for s in sheet_names if s != log_sheet]

            # 主要データの読み込みと結合
            df_list = []
            for sheet in main_sheets:
                df_tmp = pd.read_excel(xls, sheet_name=sheet)
                df_tmp["シート名"] = sheet
                if "_" in sheet:
                    df_tmp["担当者"], df_tmp["種別"] = sheet.split("_")
                else:
                    df_tmp["担当者"] = "不明"
                    df_tmp["種別"] = "不明"
                df_list.append(df_tmp)

            df = pd.concat(df_list, ignore_index=True)
            df["記入日"] = pd.to_datetime(df["記入日"], errors="coerce")

            # 地域データの正規化
            df["地域"] = df["地域"].apply(lambda x: "未分類" if pd.isna(x) or str(x).strip() == "" or str(x).startswith("その他：") else x)
            df["地域"] = df["地域"].apply(lambda x: "その他" if x not in KINIKI_AREAS and x != "未分類" else x)

            # カテゴリの抽出 (採用・不採用理由から)
            df["カテゴリ"] = df["採用・不採用理由"].apply(
                lambda x: re.findall(r"【(.*?)】", str(x))[0].split("・") if re.findall(r"【(.*?)】", str(x)) else [])

            # 訪問データフィルターフォーム
            with st.form("main_filter_form"):
                persons_all = sorted(df["担当者"].dropna().unique())
                persons = [p for p in persons_all if p != "不明"]
                types_all = sorted(df["種別"].dropna().unique())
                types = [t for t in types_all if t != "不明"]
                areas_raw = df["地域"].dropna().unique().tolist()
                areas = sorted(list(set(areas_raw + ["未分類"])))
                cats = sorted([c for c in df["大分類"].dropna().unique() if c in VALID_CATEGORIES])

                selected_persons = st.multiselect("担当者", persons, default=persons)
                selected_types = st.multiselect("種別", types, default=types)
                selected_areas = st.multiselect("地域", areas, default=areas)
                selected_categories = st.multiselect("大分類", cats, default=cats)

                min_date = df["記入日"].min()
                max_date = df["記入日"].max()

                if pd.isna(min_date) or pd.isna(max_date):
                    st.warning("「記入日」データに有効な日付が見つかりませんでした。日付フィルターは利用できません。")
                    start_date = None
                    end_date = None
                else:
                    start_date, end_date = st.date_input("記入日", [min_date, max_date])

                submitted = st.form_submit_button("🚀 分析実行")

            if submitted:
                if start_date and end_date:
                    df_filtered_calc = df[
                        df["担当者"].isin(selected_persons) &
                        df["種別"].isin(selected_types) &
                        df["地域"].isin(selected_areas) &
                        df["大分類"].isin(selected_categories) &
                        df["記入日"].between(pd.to_datetime(start_date), pd.to_datetime(end_date), inclusive="both")
                    ]
                else:
                    df_filtered_calc = df[
                        df["担当者"].isin(selected_persons) &
                        df["種別"].isin(selected_types) &
                        df["地域"].isin(selected_areas) &
                        df["大分類"].isin(selected_categories)
                    ]
                st.session_state.df_filtered_display = df_filtered_calc

        except Exception as e:
            st.error(f"エラーが発生しました：{e}")
    else:
        st.info("Excelファイルがアップロードされていません。ファイル状況をご確認ください。")


# 訪問データ分析結果の表示 (セッションステートにデータがあれば表示)
if st.session_state.df_filtered_display is not None:
    st.markdown("---")
    df_filtered_to_display = st.session_state.df_filtered_display
    st.subheader("📈 訪問データ分析")

    if df_filtered_to_display.empty:
        st.info("選択されたフィルター条件に合致する訪問データがありません。")
    else:
        uuid_df = df_filtered_to_display.drop_duplicates("UUID")
        status_counts = uuid_df["ステータス"].value_counts()
        product_count = df_filtered_to_display["商品名"].notna().sum()
        result_counts = df_filtered_to_display["結果"].value_counts()

        st.markdown("#### ステータス（UUID単位）")
        for s in ["アポ", "訪問予定", "検討中", "完了"]:
            st.write(f"- {s}：{status_counts.get(s, 0)} 件")

        st.markdown("#### 商品ステータス（商品単位）")
        for s in ["採用", "不採用", "返答待ち"]:
            val = result_counts.get(s, 0)
            rate = val / product_count if product_count else 0
            st.write(f"- {s}：{val} 件（{rate:.1%}）")

        df_saiyo = df_filtered_to_display[df_filtered_to_display["結果"] == "採用"]
        df_fusaiyo = df_filtered_to_display[df_filtered_to_display["結果"] == "不採用"]
        cat_saiyo = Counter(sum(df_saiyo["カテゴリ"], []))
        cat_fusaiyo = Counter(sum(df_fusaiyo["カテゴリ"], []))

        df_saiyo_cat = pd.DataFrame(cat_saiyo.items(), columns=["カテゴリ", "件数"])
        df_fusaiyo_cat = pd.DataFrame(cat_fusaiyo.items(), columns=["カテゴリ", "件数"])

        if not df_saiyo_cat.empty:
            df_saiyo_cat["割合"] = (df_saiyo_cat["件数"] / df_saiyo_cat["件数"].sum() * 100).round(1).astype(str) + "%"
        if not df_fusaiyo_cat.empty:
            df_fusaiyo_cat["割合"] = (df_fusaiyo_cat["件数"] / df_fusaiyo_cat["件数"].sum() * 100).round(1).astype(str) + "%"

        st.markdown("#### 採用理由カテゴリ")
        if not df_saiyo_cat.empty:
            st.dataframe(df_saiyo_cat.sort_values("件数", ascending=False), use_container_width=True)
        else:
            st.write("該当するデータがありません。")

        st.markdown("#### 不採用理由カテゴリ")
        if not df_fusaiyo_cat.empty:
            st.dataframe(df_fusaiyo_cat.sort_values("件数", ascending=False), use_container_width=True)
        else:
            st.write("該当するデータがありません。")

        if st.checkbox("📂 訪問データのフィルター後データを見る", key="view_filtered_visit_data"):
            st.dataframe(df_filtered_to_display, use_container_width=True)

