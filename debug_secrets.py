import streamlit as st
import os

st.title("Debug Secrets")
st.write(f"Current Working Directory: `{os.getcwd()}`")

st.markdown("### Secrets Content")
try:
    st.write(dict(st.secrets))
except Exception as e:
    st.error(f"Error reading secrets: {e}")

if "APP_PASSWORD" in st.secrets:
    print(f"DEBUG_SUCCESS: APP_PASSWORD found: {st.secrets['APP_PASSWORD']}")
    st.success(f"APP_PASSWORD found: `{st.secrets['APP_PASSWORD']}`")
else:
    print("DEBUG_FAILURE: APP_PASSWORD NOT found in st.secrets")
    st.error("APP_PASSWORD NOT found in st.secrets")
