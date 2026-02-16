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
        </style>
    """, unsafe_allow_html=True)

    from datetime import datetime
    import pandas as pd

    st.title("Tokumei AI - Login")
    
    with st.form("login_form"):
        user_id = st.text_input("User ID")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login", use_container_width=True)
        
        if submit:
            try:
                import urllib.parse
                
                # Spreadsheet ID
                spreadsheet_id = st.secrets["SPREADSHEET_ID"]
                worksheet_name = "アクセス管理"
                
                # Use direct CSV export URL for public sheets to avoid encoding issues with Japanese names
                encoded_worksheet = urllib.parse.quote(worksheet_name)
                csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_worksheet}"
                
                df = pd.read_csv(csv_url)
                
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
