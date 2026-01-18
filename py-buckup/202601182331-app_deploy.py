import streamlit as st
import pandas as pd
from modules.auth import check_password
from modules.database import fetch_table, get_master_list
from modules.ui import (
    load_css, custom_title, render_sidebar, 
    render_activity_log, render_related_parties, render_assets_management,
    render_person_registration, render_reports, render_data_management, render_settings
)
from modules.utils import calculate_age
from modules.constants import MAP_PERSONS

st.set_page_config(page_title="成年後見業務支援システム", layout="wide")

def main():
    if not check_password(): return
    
    load_css()
    custom_title("成年後見業務支援システム")

    df_persons = fetch_table("persons", MAP_PERSONS)
    
    if '生年月日' in df_persons.columns and not df_persons.empty:
        df_persons['年齢'] = df_persons['生年月日'].apply(calculate_age)
        df_persons['年齢'] = pd.to_numeric(df_persons['年齢'], errors='coerce')

    menu = render_sidebar()

    # Session State Initialization
    for key in ['selected_person_id', 'delete_confirm_id', 'edit_asset_id', 'delete_asset_id', 
                'edit_related_id', 'delete_related_id', 'edit_activity_id', 'edit_person_id']:
        if key not in st.session_state: st.session_state[key] = None

    # マスタデータキャッシュの取得とフォールバック
    act_opts = get_master_list('activity') or ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"]
    rel_opts = get_master_list('relationship') or ["親族", "ケアマネ", "施設相談員", "病院SW", "主治医", "弁護士", "行政", "その他"]
    ast_opts = get_master_list('asset') or ["預貯金", "現金", "有価証券", "保険", "不動産", "負債", "その他"]
    guard_opts = get_master_list('guardian_type') or ["後見", "保佐", "補助", "任意", "未成年後見", "その他"]

    if menu == "利用者情報・活動記録":
        render_activity_log(df_persons, act_opts)
    elif menu == "関係者・連絡先":
        render_related_parties(df_persons, rel_opts)
    elif menu == "財産管理":
        render_assets_management(df_persons, ast_opts)
    elif menu == "利用者情報登録":
        render_person_registration(df_persons, guard_opts)
    elif menu == "帳票作成":
        render_reports(df_persons)
    elif menu == "データ管理・移行":
        render_data_management()
    elif menu == "初期設定":
        render_settings()

if __name__ == "__main__":
    main()