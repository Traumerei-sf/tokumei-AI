import streamlit as st
import os

st.set_page_config(page_title="プライバシーポリシー - 特命AI", page_icon="🛡️")

def show_privacy_policy():
    st.title("プライバシーポリシー")
    
    # ルートディレクトリの privacy_policy_v1.md を読み込んで表示
    filepath = os.path.join(os.path.dirname(__file__), "..", "privacy_policy_v1.md")
    
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        st.markdown(content)
    else:
        st.error("プライバシーポリシーのファイルが見つかりません。")
        
    st.markdown("---")
    if st.button("ログイン画面へ戻る"):
        st.switch_page("app.py")

show_privacy_policy()
