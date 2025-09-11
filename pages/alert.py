import streamlit as st
import pandas as pd
import os
import json

SAVE_DIR = "uploaded_files"
STATE_FILE = os.path.join(SAVE_DIR, "state.json")
os.makedirs(SAVE_DIR, exist_ok=True)

# ------------------------
# 状態ファイル管理
# ------------------------
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"week1": None, "week2": None, "week3": None, "helper": None}

def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

# ------------------------
# データ読込関数
# ------------------------
def load_weekly_file(filepath):
    df = pd.read_excel(filepath, header=5)
    df = df[['得意先コード', '得意先名']].dropna()
    df = df.rename(columns={'得意先コード': '取引先コード', '得意先名': '取引先名'})
    df['取引先コード'] = df['取引先コード'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(4)
    return df.drop_duplicates(subset=['取引先コード'])

def load_helper_file(filepath):
    xls = pd.ExcelFile(filepath)
    delete_list = pd.read_excel(xls, sheet_name="削除依頼", header=None)[0].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(4).tolist()
    
    # 取引先リストシートを読み込み、ヘッダーを設定
    base_list = pd.read_excel(xls, sheet_name="取引先リスト", header=None, names=["取引先コード","取引先名", "大分類"])
    base_list['取引先コード'] = base_list['取引先コード'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(4)

    # 離脱リスト：A列=コード, B列=備考
    leave_df = pd.read_excel(xls, sheet_name="離脱リスト", header=None)
    leave_df[0] = leave_df[0].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(4)
    leave_map = dict(zip(leave_df[0], leave_df[1].fillna("")))

    # 大分類の辞書を作成
    category_map = dict(zip(base_list['取引先コード'], base_list['大分類'].fillna("")))

    return delete_list, base_list, leave_map, category_map

# ------------------------
# 分析関数
# ------------------------
def analyze(week1, week2, week3, helper):
    delete_list, base_list, leave_map, category_map = helper

    set1 = set(week1['取引先コード']) - set(delete_list)
    set2 = set(week2['取引先コード']) - set(delete_list)
    set3 = set(week3['取引先コード']) - set(delete_list)

    name_map = {}
    for df in [week1, week2, week3, base_list]:
        for _, row in df.iterrows():
            name_map[row['取引先コード']] = row['取引先名']

    # 2週間未取引
    two_weeks_none = [c for c in set1 if c not in set2 and c not in set3]
    df_two = pd.DataFrame([(c, name_map.get(c, ""), category_map.get(c, "")) for c in two_weeks_none], columns=["取引先コード", "取引先名", "大分類"])

    # 3週間未取引
    all_weeks = set1 | set2 | set3
    three_weeks_none = [c for c in base_list['取引先コード'] if c not in all_weeks]

    rows_normal = []
    rows_leave = []
    for c in three_weeks_none:
        name = name_map.get(c, "")
        category = category_map.get(c, "")
        if c in leave_map:
            rows_leave.append((c, name, category, leave_map[c]))
        else:
            rows_normal.append((c, name, category, ""))

    # 通常 → 離脱の順に結合
    df_three = pd.DataFrame(rows_normal + rows_leave, columns=["取引先コード", "取引先名", "大分類", "備考"])
    return df_two, df_three

# ------------------------
# Streamlit UI
# ------------------------
st.set_page_config(page_title="取引先分析ツール", layout="wide")
st.title("🚨離脱アラート🚨")

if "state" not in st.session_state:
    st.session_state.state = load_state()

left, right = st.columns([1, 2])

# 左側：ファイル状況表示（ベージュ背景）
with left:
    st.subheader("📂 現在のファイル状況")
    file_status_html = "<div style='background-color:#f9f5e9; padding:15px; border-radius:8px; border:1px solid #ddd;'>"
    for label in ["week1", "week2", "week3", "helper"]:
        info = st.session_state.state.get(label)
        if info and "name" in info:
            file_status_html += f"<p><strong>{label}</strong>: ✅ {info['name']}</p>"
        else:
            file_status_html += f"<p><strong>{label}</strong>: ❌ 未設定</p>"
    file_status_html += "</div>"
    st.markdown(file_status_html, unsafe_allow_html=True)

# 右側：ファイルアップロード
with right:
    st.header("📁 ファイルアップロード")
    week1_file = st.file_uploader("先々週の取引先リスト", type=["xlsx"], key="week1")
    week2_file = st.file_uploader("先週の取引先リスト", type=["xlsx"], key="week2")
    week3_file = st.file_uploader("今週の取引先リスト", type=["xlsx"], key="week3")
    helper_file = st.file_uploader("補助データファイル", type=["xlsx"], key="helper")

# ファイル保存関数
def save_file(uploaded_file, label):
    if uploaded_file:
        filepath = os.path.join(SAVE_DIR, f"{label}.xlsx")
        with open(filepath, "wb") as f:
            f.write(uploaded_file.read())
        st.session_state.state[label] = {"path": filepath, "name": uploaded_file.name}
        save_state(st.session_state.state)
        return filepath
    info = st.session_state.state.get(label)
    return info["path"] if info else None

# 保存処理
week1_path = save_file(week1_file, "week1")
week2_path = save_file(week2_file, "week2")
week3_path = save_file(week3_file, "week3")
helper_path = save_file(helper_file, "helper")

# 繰り上げ処理
if st.button("🔁 繰り上げ処理"):
    st.session_state.state["week1"] = st.session_state.state.get("week2")
    st.session_state.state["week2"] = st.session_state.state.get("week3")
    st.session_state.state["week3"] = None
    save_state(st.session_state.state)
    st.success("繰り上げ完了！ページを手動でリロードしてください（Ctrl+R または F5）")

# 分析実行
if st.button("🚀 分析実行"):
    try:
        w1 = load_weekly_file(st.session_state.state["week1"]["path"])
        w2 = load_weekly_file(st.session_state.state["week2"]["path"])
        w3 = load_weekly_file(st.session_state.state["week3"]["path"])
        helper = load_helper_file(st.session_state.state["helper"]["path"])

        df_two, df_three = analyze(w1, w2, w3, helper)

        st.subheader("📉 2週間取引なし")
        st.dataframe(df_two, use_container_width=True)

        st.subheader("📉 3週間以上取引なし")
        st.dataframe(df_three, use_container_width=True)

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")

