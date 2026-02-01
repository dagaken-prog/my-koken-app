import google.generativeai as genai
import streamlit as st
import os

def summarize_text(text):
    """
    入力されたテキストを要約し、活動記録に適した形式に整形する関数
    """
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        return "エラー: GEMINI_API_KEY が secrets.toml に設定されていません。"

    try:
        genai.configure(api_key=api_key)
        # 動作確認済みモデル: gemini-flash-latest
        model = genai.GenerativeModel('gemini-flash-latest')
        
        prompt = f"""
        以下のテキストは、成年後見業務における活動記録の下書き（音声入力やメモなど）です。
        この内容を、活動記録として適切な、簡潔かつ明確な日本語の文章に要約・整形してください。
        
        テキスト:
        {text}
        
        出力（要約のみ）:
        """
        
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"エラーが発生しました: {str(e)}"
