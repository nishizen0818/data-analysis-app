import streamlit as st

st.set_page_config(
    page_title="メインメニュー",
    page_icon="🏠",
    layout="centered"
)

st.title("分析ツール(202509)")
st.write("以下から見たい分析ページを選択してください。")

st.markdown("---")  # 区切り線

# 各ページへのリンクを設置
st.page_link("pages/attacklist.py", label="アタックリスト分析📊", icon="📊")
st.page_link("pages/sales.py", label="卸営業数値分析📈", icon="📈")
st.page_link("pages/item.py", label="アイテム別集計📦", icon="📦")
st.page_link("pages/alert.py", label="離脱アラート🚨", icon="🚨")

st.markdown("---")
st.info("💡 各リンクをクリックすると、それぞれの分析ページに移動します。")


