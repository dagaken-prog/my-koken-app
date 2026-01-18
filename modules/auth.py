import streamlit as st

def check_password():
    """
    パスワード認証を行う関数。
    認証成功ならTrue、失敗・未認証ならFalseを返し、ログインフォームを表示する。
    """
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if st.session_state.password_correct:
        return True
    
    with st.container():
        with st.form("login_form"):
            st.markdown("## 🔒 ログイン")
            password = st.text_input("パスワードを入力してください", type="password")
            submitted = st.form_submit_button("ログイン")
            
            if submitted:
                if "APP_PASSWORD" in st.secrets:
                    if password == st.secrets["APP_PASSWORD"]:
                        st.session_state.password_correct = True
                        st.success("ログインしました")
                        st.rerun()
                    else:
                        st.error("パスワードが違います")
                else:
                    st.error("管理用パスワードが未設定です。secrets.tomlを確認してください。")
    return False
