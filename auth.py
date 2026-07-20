import streamlit as st

def login():
    st.markdown("""
        <style>
            .login-container {
                max-width: 400px;
                margin: auto;
                padding: 2rem;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                background: white;
            }
            /* ログイン中の裏側のローディングログ（SpinnerやToast）を強制非表示にする */
            div[data-testid="stToastContainer"], 
            div[data-testid="stStatusWidget"],
            div[data-testid="stStatusContainer"],
            .stSpinner {
                display: none !important;
            }
        </style>
    """, unsafe_allow_html=True)

    from datetime import datetime
    import pandas as pd
    from process.u_googleSheets import read_sheet

    st.title("Tokumei AI - Login")

    with st.form("login_form"):
        user_id = st.text_input("User ID")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login", use_container_width=True)

        if submit:
            try:
                # サービスアカウント経由でアクセス管理シートを読み取る(閲覧制限されたシートのため)
                df = read_sheet("アクセス管理", ttl=0)

                # Check for matching credentials
                match = df[(df['id'] == user_id) & (df['password'] == password)]
                
                if not match.empty:
                    user_data = match.iloc[0]
                    
                    # Date validation
                    today = datetime.now().date()
                    
                    # Check for null dates and handle validation
                    start_date = pd.to_datetime(user_data['perStartDate'])
                    end_date = pd.to_datetime(user_data['perEndDate'])
                    
                    is_after_start = True
                    if pd.notnull(start_date):
                        is_after_start = today >= start_date.date()
                        
                    is_before_end = True
                    if pd.notnull(end_date):
                        is_before_end = today <= end_date.date()
                    
                    if is_after_start and is_before_end:
                        st.session_state["authenticated"] = True
                        st.session_state["user_id"] = user_id
                        st.session_state["user_info"] = user_data.to_dict()
                        st.success("Login successful!")
                        st.rerun()
                    else:
                        error_msg = "Access period validation failed."
                        if not is_after_start:
                            error_msg = f"Access period has not yet started (Starts: {start_date.date()})."
                        elif not is_before_end:
                            error_msg = f"Access period has expired (Ended: {end_date.date()})."
                        st.error(error_msg)
                else:
                    st.error("Invalid User ID or Password")
            except Exception as e:
                st.error(f"An error occurred during login: {str(e)}")
                
    st.markdown("<p style='font-size: 12px; text-align: center; margin-top: 10px;'>ログインすることで、以下のプライバシーポリシーに同意したものとみなします。</p>", unsafe_allow_html=True)
    st.page_link("pages/privacy.py", label="プライバシーポリシーを確認する", icon="🛡️")

def logout():
    st.session_state["authenticated"] = False
    st.rerun()

def check_auth():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    
    if not st.session_state["authenticated"]:
        login()
        return False
    return True
