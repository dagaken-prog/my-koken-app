import streamlit as st
import pandas as pd
import datetime
import io
import openpyxl
import re
from .constants import (
    MAP_PERSONS, MAP_ACTIVITIES, MAP_ASSETS, MAP_RELATED, MAP_SYSTEM, MAP_MASTER
)
from .utils import calculate_age, to_safe_id
from .database import (
    fetch_table, insert_data, update_data, delete_data, process_import, check_usage_count
)
from .ai import summarize_text

# --- CSSロード ---
def load_css():
    st.markdown("""
        <meta http-equiv="Content-Language" content="ja">
        <script>
        document.documentElement.lang = 'ja';
        try {
            Object.defineProperty(navigator, 'language', {
                get: function() { return 'ja-JP'; }
            });
            Object.defineProperty(navigator, 'languages', {
                get: function() { return ['ja-JP', 'ja']; }
            });
        } catch (e) {
            console.log(e);
        }
        }
        </script>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <style>
        html, body, [class*="css"] { font-family: "Noto Sans JP", sans-serif; color: #333; }
        
        .block-container { 
            padding-top: 6rem !important; 
            padding-bottom: 3rem !important; 
            padding-left: 1rem !important; 
            padding-right: 1rem !important; 
        }
        
        div[data-testid="stVerticalBlock"] { gap: 0.5rem !important; }
        div[data-testid="stElementContainer"] { margin-bottom: 0.3rem !important; }
        
        div[data-testid="stBorder"] { 
            margin: 5px 0 !important; 
            padding: 10px !important; 
            border: 1px solid #ddd !important; 
            border-radius: 8px; 
            background-color: #fff;
        }
        
        .streamlit-expanderHeader {
            font-size: 14px !important;
            font-weight: bold !important;
            background-color: #f9f9f9;
            border: 1px solid #ddd;
            border-radius: 8px;
            margin-bottom: 5px;
            white-space: normal !important;
            height: auto !important;
        }

        [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { padding: 6px !important; font-size: 14px !important; }
        
        p { margin-bottom: 0.5rem !important; line-height: 1.6 !important; }
        h2 { padding: 10px 0 !important; margin-bottom: 20px !important; line-height: 1.5 !important; }
        
        .custom-title { font-size: 20px; font-weight: bold; color: #006633; border-left: 6px solid #006633; padding: 5px 0 5px 10px; margin: 5px 0 10px 0; background-color: #f8f9fa; }
        .custom-header { font-size: 16px; font-weight: bold; color: #006633; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin: 20px 0 10px 0; }
        .custom-header-text { font-size: 16px; font-weight: bold; color: #006633; margin: 0; padding-top: 5px; white-space: nowrap; }
        .custom-header-line { border-bottom: 1px solid #ccc; margin: 0 0 5px 0; }
        
        .stTextInput input, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea, .stNumberInput input { border: 1px solid #666 !important; background-color: #fff !important; border-radius: 6px !important; padding: 8px 8px !important; font-size: 14px !important; }
        .stSelectbox div[data-baseweb="select"] > div { height: auto !important; min-height: 40px !important; }
        .stTextInput label, .stSelectbox label, .stDateInput label, .stTextArea label, .stNumberInput label, .stCheckbox label { margin-bottom: 2px !important; font-size: 13px !important; font-weight: bold; }
        
        div[data-testid="stPopover"] button { padding: 0 8px !important; height: auto !important; border: 1px solid #ccc !important; }
        section[data-testid="stSidebar"] button { width: 100%; border: 1px solid #ccc; border-radius: 8px; margin-bottom: 8px; padding: 12px; font-size: 16px !important; font-weight: bold; text-align: left; background-color: white; color: #333; }
        section[data-testid="stSidebar"] button:hover { border-color: #006633; color: #006633; background-color: #f0fff0; }
        
        [data-testid="stFileUploaderDropzone"] div div span, [data-testid="stFileUploaderDropzone"] div div small { display: none; }
        [data-testid="stFileUploaderDropzone"] div div::after { content: "ファイルをドラッグ＆ドロップまたは選択"; font-size: 12px; font-weight: bold; color: #333; display: block; margin: 5px 0; }
        [data-testid="stFileUploaderDropzone"] div div::before { content: "CSV/Excelファイル (200MBまで)"; font-size: 12px; color: #666; display: block; margin-bottom: 5px; }
        </style>
    """, unsafe_allow_html=True)

def custom_title(text):
    st.markdown(f'<div class="custom-title">{text}</div>', unsafe_allow_html=True)

def custom_header(text, help_text=None):
    # help_text引数は互換性のために残しますが、機能は無効化します（ボタンは表示しません）
    st.markdown(f'<div class="custom-header">{text}</div>', unsafe_allow_html=True)

def fill_excel_template(template_file, data_dict):
    wb = openpyxl.load_workbook(template_file)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    text = cell.value
                    matches = re.findall(r'\{\{(.*?)\}\}', text)
                    if matches:
                        new_text = text
                        for key in matches:
                            if key in data_dict:
                                val = str(data_dict[key]) if data_dict[key] is not None else ""
                                new_text = new_text.replace(f'{{{{{key}}}}}', val)
                        cell.value = new_text
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def render_sidebar():
    # サイドバー自動クローズロジック
    if st.session_state.get('close_sidebar_flag'):
        js = """
        <script>
            var sidebar = window.parent.document.querySelector('section[data-testid="stSidebar"]');
            if (sidebar && sidebar.getAttribute('aria-expanded') === 'true') {
                var collapse = window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"]');
                if (collapse) {
                    collapse.click();
                } else {
                    // モバイル等の場合、Xボタンを探す
                     var closeBtn = sidebar.querySelector('button[kind="header"]');
                     if (closeBtn) { closeBtn.click(); }
                }
            }
        </script>
        """
        st.markdown(js, unsafe_allow_html=True)
        st.session_state['close_sidebar_flag'] = False

    if 'current_menu' not in st.session_state:
        st.session_state.current_menu = "利用者情報・活動記録"

    with st.sidebar:
        st.markdown("### メニュー")
        menu_items = [
            ("利用者情報・活動記録", "利用者情報・活動記録"),
            ("関係者・連絡先", "関係者・連絡先"),
            ("財産管理", "財産管理"),
            ("利用者情報登録", "利用者情報登録"),
            ("帳票作成", "帳票作成"),
            ("データ管理・移行", "データ管理・移行"),
            ("初期設定", "初期設定")
        ]
        for label, key_val in menu_items:
            display_label = f"👉 {label}" if st.session_state.current_menu == key_val else label
            if st.button(display_label, key=f"menu_btn_{key_val}", use_container_width=True):
                st.session_state.current_menu = key_val
                st.session_state['close_sidebar_flag'] = True
                st.rerun()
    return st.session_state.current_menu

def render_activity_log(df_persons, act_opts):
    df_activities = fetch_table("activities", MAP_ACTIVITIES)
    df_related = fetch_table("related_parties", MAP_RELATED)
    
    custom_header("受任中利用者一覧", help_text="一覧から対象者をクリックすると詳細が表示されます。")
    
    if not df_persons.empty and '現在の状態' in df_persons.columns:
        mask = df_persons['現在の状態'].fillna('').astype(str).isin(['受任中', '', 'nan'])
        df_active = df_persons[mask].copy()
        if df_active.empty: df_active = df_persons.copy()
    else:
        df_active = pd.DataFrame(columns=MAP_PERSONS.keys())

    display_cols = ['ケース番号', '氏名', '生年月日', '年齢', '類型']
    df_display = df_active[display_cols] if not df_active.empty else pd.DataFrame(columns=display_cols)
    
    selection = st.dataframe(
        df_display,
        column_config={
            "ケース番号": st.column_config.TextColumn("No."),
            "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
            "類型": st.column_config.TextColumn("後見類型"),
        },
        use_container_width=True, on_select="rerun", selection_mode="single-row", hide_index=True
    )

    if selection.selection.rows:
        idx = selection.selection.rows[0]
        selected_row = df_active.iloc[idx]
        current_pid = selected_row['person_id']
        st.session_state.selected_person_id = current_pid
        
        st.markdown("---")
        age_val = selected_row.get('年齢')
        age_str = f" ({int(age_val)}歳)" if pd.notnull(age_val) else ""
        custom_header(f"{selected_row.get('氏名')}{age_str} さんの詳細・活動記録")

        kp_html = ""
        if not df_related.empty:
            df_related['safe_pid'] = df_related['person_id'].apply(to_safe_id)
            current_pid_safe = to_safe_id(current_pid)
            
            kp_df = df_related[
                (df_related['safe_pid'] == current_pid_safe) & 
                (df_related['キーパーソン'] == True)
            ]
            if not kp_df.empty:
                kp_html = "<div style='margin-top:8px; padding-top:8px; border-top:1px dashed #ccc; width:100%; grid-column: 1 / -1;'>"
                kp_html += "<div><b>★ キーパーソン:</b></div>"
                for _, kp in kp_df.iterrows():
                    tel = kp.get('電話番号')
                    tel_html = f'<a href="tel:{tel}" style="text-decoration:none; color:#0066cc;">📞 {tel}</a>' if tel else ''
                    kp_html += f"<div style='margin-left:10px;'>【{kp.get('関係種別')}】 {kp.get('氏名')} {tel_html}</div>"
                kp_html += "</div>"

        with st.expander("▼ 基本情報", expanded=True):
            grid_html = f"""
            <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; font-size: 14px;">
                <div><span style="font-weight:bold; color:#555;">No.:</span> {selected_row.get('ケース番号')}</div>
                <div><span style="font-weight:bold; color:#555;">事件番号:</span> {selected_row.get('基本事件番号')}</div>
                <div><span style="font-weight:bold; color:#555;">類型:</span> {selected_row.get('類型')}</div>
                <div><span style="font-weight:bold; color:#555;">氏名:</span> {selected_row.get('氏名')}</div>
                <div><span style="font-weight:bold; color:#555;">ｼﾒｲ:</span> {selected_row.get('ｼﾒｲ')}</div>
                <div><span style="font-weight:bold; color:#555;">生年月日:</span> {selected_row.get('生年月日')}</div>
                <div style="grid-column: 1 / -1;"><span style="font-weight:bold; color:#555;">住所:</span> {selected_row.get('住所') or '-'}</div>
                <div style="grid-column: 1 / -1;"><span style="font-weight:bold; color:#555;">居所:</span> {selected_row.get('居所') or '-'}</div>
                <div><span style="font-weight:bold; color:#555;">障害類型:</span> {selected_row.get('障害類型')}</div>
                <div><span style="font-weight:bold; color:#555;">申立人:</span> {selected_row.get('申立人')}</div>
                <div><span style="font-weight:bold; color:#555;">審判日:</span> {selected_row.get('審判確定日')}</div>
                <div><span style="font-weight:bold; color:#555;">家裁:</span> {selected_row.get('管轄家裁')}</div>
                <div><span style="font-weight:bold; color:#555;">報告月:</span> {selected_row.get('家裁報告月')}</div>
                <div><span style="font-weight:bold; color:#555;">状態:</span> {selected_row.get('現在の状態')}</div>
                {kp_html}
            </div>
            """
            st.markdown(grid_html, unsafe_allow_html=True)
        
        st.markdown("### 📝 活動記録")
        with st.expander("➕ 新しい活動記録を追加する", expanded=False):
            # Session Stateの初期化 (入力保持用)
            if 'new_act_summary' not in st.session_state: st.session_state.new_act_summary = ""
            if 'new_act_place' not in st.session_state: st.session_state.new_act_place = ""
            if 'new_act_time' not in st.session_state: st.session_state.new_act_time = 0
            if 'new_act_cost' not in st.session_state: st.session_state.new_act_cost = 0
            

            st.caption("📝 下書きや音声入力したテキストを「内容」に入力し、「🤖 AI要約」ボタンを押すと、活動記録に適した形式に自動整形します。")
            if st.button("🤖 AI要約実行 (下書きを整形)"):
                if st.session_state.new_act_summary:
                    with st.spinner("AIが要約を行っています..."):
                        summarized = summarize_text(st.session_state.new_act_summary)
                        st.session_state.new_act_summary = summarized
                        st.rerun()
                else:
                    st.warning("まずは「内容」にテキストを入力してください。")

            # 入力フォーム (st.formは使用せず、セッション状態で管理)
            input_summary = st.text_area("内容", key="new_act_summary", height=120)
            
            col1, col2 = st.columns(2)
            in_date = col1.date_input("活動日", datetime.date.today(), key="new_act_date", format="YYYY/MM/DD")
            in_type = col2.selectbox("活動", act_opts, key="new_act_type")
            
            col3, col4, col5, col6 = st.columns(4)
            in_time = col3.number_input("時間(分)", min_value=0, step=10, key="new_act_time")
            in_place = col4.text_input("場所", placeholder="自宅、病院など", key="new_act_place")
            in_cost = col5.number_input("費用(円)", min_value=0, step=100, key="new_act_cost")
            in_imp = col6.checkbox("★重要", key="new_act_imp")
            
            def on_register_click():
                if not st.session_state.new_act_summary:
                    st.toast("内容は必須です", icon="⚠️")
                    return

                new_data = {
                    'person_id': current_pid, 
                    '記録日': str(st.session_state.new_act_date), 
                    '活動': st.session_state.new_act_type,
                    '場所': st.session_state.new_act_place, 
                    '所要時間': st.session_state.new_act_time, 
                    '交通費・立替金': st.session_state.new_act_cost,
                    '重要': st.session_state.new_act_imp, 
                    '要点': st.session_state.new_act_summary
                }
                
                # insert_data内でst.toastが表示される
                if insert_data("activities", new_data, MAP_ACTIVITIES):
                    # 入力内容のクリア（コールバック内なら安全にクリア可能）
                    st.session_state.new_act_summary = ""
                    st.session_state.new_act_place = ""
                    st.session_state.new_act_time = 0
                    st.session_state.new_act_cost = 0

            st.button("登録", type="primary", on_click=on_register_click)

        custom_header("過去の活動履歴", help_text="履歴の「詳細・操作」を開くと編集・削除ができます。")
        if not df_activities.empty:
            df_activities['safe_pid'] = df_activities['person_id'].apply(to_safe_id)
            current_pid_safe = to_safe_id(current_pid)
            
            my_acts = df_activities[df_activities['safe_pid'] == current_pid_safe].copy()
            
            if not my_acts.empty:
                if '作成日時' in my_acts.columns:
                    my_acts['作成日時'] = pd.to_datetime(my_acts['作成日時'], errors='coerce')
                    my_acts = my_acts.sort_values(by=['記録日', '作成日時'], ascending=[False, False])
                else:
                    my_acts = my_acts.sort_values('記録日', ascending=False)
                
                if st.session_state.edit_activity_id:
                    edit_row = my_acts[my_acts['activity_id'] == st.session_state.edit_activity_id].iloc[0]
                    with st.container(border=True):
                        st.markdown(f"#### ✏️ 修正")
                        with st.form("edit_act_form"):
                            ed_note = st.text_area("内容", value=edit_row['要点'], height=120)
                            
                            c_d, c_t = st.columns(2)
                            ed_date = c_d.date_input("活動日", pd.to_datetime(edit_row['記録日']), format="YYYY/MM/DD")
                            try:
                                idx = act_opts.index(edit_row['活動'])
                            except:
                                idx = 0
                            ed_type = c_t.selectbox("活動", act_opts, index=idx)
                            
                            c3, c4, c5, c6 = st.columns(4)
                            val_time = edit_row.get('所要時間')
                            if pd.isna(val_time) or val_time == "": val_time = 0
                            ed_time = c3.number_input("時間", value=int(val_time), min_value=0, step=10)

                            ed_place = c4.text_input("場所", value=str(edit_row.get('場所') or ""))

                            val_cost = edit_row.get('交通費・立替金')
                            if pd.isna(val_cost) or val_cost == "": val_cost = 0
                            ed_cost = c5.number_input("費用", value=int(val_cost), min_value=0, step=100)
                            
                            ed_imp = c6.checkbox("重要", value=bool(edit_row['重要']))
                            
                            c_sv, c_cl = st.columns(2)
                            if c_sv.form_submit_button("保存"):
                                upd_data = {'記録日': str(ed_date), '活動': ed_type, '場所': ed_place, '所要時間': ed_time, '交通費・立替金': ed_cost, '重要': ed_imp, '要点': ed_note}
                                if update_data("activities", "activity_id", st.session_state.edit_activity_id, upd_data, MAP_ACTIVITIES):
                                    st.session_state.edit_activity_id = None
                                    st.rerun()
                            if c_cl.form_submit_button("キャンセル"):
                                st.session_state.edit_activity_id = None
                                st.rerun()

                for _, row in my_acts.iterrows():
                    star = "★" if row['重要'] else ""
                    with st.container(border=True):
                        summary = row.get('要点', '') or ''
                        label_text = f"{star} {row['記録日']} | {summary}"
                        
                        with st.expander(label_text, expanded=False):
                            st.markdown(f"**活動種別:** {row['活動']}")
                            st.markdown(f"""
                            - **場所:** {row.get('場所') or '-'}
                            - **時間:** {row.get('所要時間') or '0'} 分
                            - **費用:** {row.get('交通費・立替金') or '0'} 円
                            """)
                            st.markdown("---")
                            c_ed, c_dl = st.columns(2)
                            if c_ed.button("編集", key=f"ed_act_{row['activity_id']}"):
                                st.session_state.edit_activity_id = row['activity_id']
                                st.rerun()
                            if c_dl.button("削除", key=f"dl_act_{row['activity_id']}"):
                                st.session_state.delete_confirm_id = row['activity_id']
                                st.rerun()
                            
                            if st.session_state.delete_confirm_id == row['activity_id']:
                                st.warning("本当に削除しますか？")
                                if st.button("はい、削除", key=f"yes_act_{row['activity_id']}"):
                                    if delete_data("activities", "activity_id", row['activity_id'], MAP_ACTIVITIES):
                                        st.session_state.delete_confirm_id = None
                                        st.rerun()
            else:
                if my_acts.empty:
                    st.write("まだ記録がありません。")

def render_related_parties(df_persons, rel_opts):
    custom_header("関係者・連絡先")
    person_opts = {f"{r['氏名']}": r['person_id'] for _, r in df_persons.iterrows()}
    
    # 選択状態の維持ロジック
    default_idx = 0
    if st.session_state.selected_person_id:
        # 名前（キー）を探す
        for i, (name, pid) in enumerate(person_opts.items()):
            if pid == st.session_state.selected_person_id:
                default_idx = i
                break
    
    target_name = st.selectbox("対象者", list(person_opts.keys()), index=default_idx)
    
    if target_name:
        pid = person_opts[target_name]
        # 選択されたらセッション状態も更新
        if pid != st.session_state.selected_person_id:
            st.session_state.selected_person_id = pid
        pid = person_opts[target_name]
        
        # 編集フォーム
        if st.session_state.edit_related_id:
            df_rel_all = fetch_table("related_parties", MAP_RELATED)
            df_rel_all['related_id'] = df_rel_all['related_id'].apply(to_safe_id)
            target_rid_safe = to_safe_id(st.session_state.edit_related_id)
            
            edit_rows = df_rel_all[df_rel_all['related_id'] == target_rid_safe]
            if not edit_rows.empty:
                edit_row = edit_rows.iloc[0]
                st.markdown(f"#### ✏️ 編集: {edit_row['氏名']}")
                with st.form("edit_rel_form"):
                    c1, c2 = st.columns(2)
                    try: idx = rel_opts.index(edit_row['関係種別'])
                    except: idx = 0
                    er_type = c1.selectbox("種別", rel_opts, index=idx)
                    er_name = c2.text_input("氏名", value=edit_row['氏名'])
                    er_org = st.text_input("所属", value=edit_row['所属・名称'])
                    c3, c4 = st.columns(2)
                    er_tel = c3.text_input("電話", value=edit_row['電話番号'])
                    er_mail = c4.text_input("Email", value=edit_row['e-mail'])
                    er_zip = c3.text_input("〒", value=edit_row['〒'])
                    er_addr = c4.text_input("住所", value=edit_row['住所'])
                    curr_kp = True if str(edit_row.get('キーパーソン', '')).upper() == 'TRUE' else False
                    er_kp = st.checkbox("★キーパーソン", value=curr_kp)
                    er_memo = st.text_area("メモ", value=edit_row['連携メモ'])
                    
                    c_sv, c_cl = st.columns(2)
                    if c_sv.form_submit_button("保存"):
                        k_str = "TRUE" if er_kp else ""
                        upd_dict = {
                            '関係種別': er_type, '氏名': er_name, '所属・名称': er_org, 
                            '電話番号': er_tel, 'e-mail': er_mail, '〒': er_zip, '住所': er_addr, 
                            'キーパーソン': k_str, '連携メモ': er_memo
                        }
                        if update_data("related_parties", "related_id", st.session_state.edit_related_id, upd_dict, MAP_RELATED):
                            st.session_state.edit_related_id = None
                            st.rerun()
                    if c_cl.form_submit_button("キャンセル"):
                        st.session_state.edit_related_id = None
                        st.rerun()
                st.markdown("---")

        with st.expander("➕ 新しい関係者を追加", expanded=False):
            with st.form("new_rel"):
                c1, c2 = st.columns(2)
                r_type = c1.selectbox("種別", rel_opts)
                r_name = c2.text_input("氏名")
                r_org = st.text_input("所属")
                c3, c4 = st.columns(2)
                r_tel = c3.text_input("電話")
                r_mail = c4.text_input("Email")
                r_zip = c3.text_input("〒")
                r_addr = c4.text_input("住所")
                r_kp = st.checkbox("★キーパーソン")
                r_memo = st.text_area("メモ")
                if st.form_submit_button("登録"):
                    new_data = {'person_id': pid, '関係種別': r_type, '氏名': r_name, '所属・名称': r_org, '電話番号': r_tel, 'e-mail': r_mail, '〒': r_zip, '住所': r_addr, 'キーパーソン': r_kp, '連携メモ': r_memo}
                    if insert_data("related_parties", new_data, MAP_RELATED):
                        st.rerun()
        
        st.markdown("---")
        df_rel = fetch_table("related_parties", MAP_RELATED)
        if not df_rel.empty:
            df_rel['safe_pid'] = df_rel['person_id'].apply(to_safe_id)
            current_pid_safe = to_safe_id(pid)
            my_rel = df_rel[df_rel['safe_pid'] == current_pid_safe]
            
            for _, row in my_rel.iterrows():
                kp_mark = "★" if str(row.get('キーパーソン', '')).upper() == 'TRUE' else ""
                label_text = f"{kp_mark}【{row['関係種別']}】 {row['氏名']} ({row['所属・名称']})"
                
                with st.expander(label_text, expanded=False):
                    tel_link = f"[{row['電話番号']}](tel:{row['電話番号']})" if row['電話番号'] else "なし"
                    email_link = f"[{row['e-mail']}](mailto:{row['e-mail']})" if row['e-mail'] else "なし"
                    
                    st.markdown(f"**電話:** {tel_link}　　**Email:** {email_link}")
                    st.markdown(f"**住所:** 〒{row.get('〒','')} {row.get('住所','')}")
                    if row['連携メモ']: st.info(f"📝 {row['連携メモ']}")
                    
                    c_ed, c_dl = st.columns(2)
                    if c_ed.button("編集", key=f"rel_edit_{row['related_id']}"):
                        st.session_state.edit_related_id = row['related_id']
                        st.rerun()
                    if c_dl.button("削除", key=f"del_rel_{row['related_id']}"):
                        if delete_data("related_parties", "related_id", row['related_id'], MAP_RELATED):
                            st.rerun()
        else:
            st.info("登録された関係者はいません。")

def render_assets_management(df_persons, ast_opts):
    custom_header("財産管理")
    person_opts = {f"{r['氏名']}": r['person_id'] for _, r in df_persons.iterrows()}
    
    # 選択状態の維持ロジック
    default_idx = 0
    if st.session_state.selected_person_id:
        for i, (name, pid) in enumerate(person_opts.items()):
            if pid == st.session_state.selected_person_id:
                default_idx = i
                break

    target_name = st.selectbox("対象者", list(person_opts.keys()), index=default_idx)
    
    if target_name:
        pid = person_opts[target_name]
        # 選択されたらセッション状態も更新
        if pid != st.session_state.selected_person_id:
            st.session_state.selected_person_id = pid
        pid = person_opts[target_name]
        with st.expander("➕ 財産追加", expanded=False):
            with st.form("new_asset"):
                c1, c2 = st.columns(2)
                a_type = c1.selectbox("種別", ast_opts)
                a_name = c2.text_input("名称")
                c3, c4 = st.columns(2)
                a_det = c3.text_input("詳細")
                a_num = c4.text_input("口座番号等")
                a_val = c1.text_input("評価額")
                a_loc = c2.text_input("保管場所")
                a_rem = st.text_area("備考")
                if st.form_submit_button("登録"):
                    nd = {'person_id': pid, '財産種別': a_type, '名称・機関名': a_name, '支店・詳細': a_det, '口座番号・記号': a_num, '評価額・残高': a_val, '保管場所': a_loc, '備考': a_rem}
                    if insert_data("assets", nd, MAP_ASSETS):
                        st.rerun()
        
        st.markdown("---")
        df_assets = fetch_table("assets", MAP_ASSETS)
        if not df_assets.empty:
            df_assets['safe_pid'] = df_assets['person_id'].apply(to_safe_id)
            current_pid_safe = to_safe_id(pid)
            my_assets = df_assets[df_assets['safe_pid'] == current_pid_safe]
            
            for _, row in my_assets.iterrows():
                label_text = f"【{row['財産種別']}】 {row['名称・機関名']} ({row['評価額・残高']})"
                with st.expander(label_text, expanded=False):
                    st.markdown(f"""
                    - **詳細:** {row['支店・詳細']}
                    - **番号:** {row['口座番号・記号']}
                    - **場所:** {row['保管場所']}
                    - **備考:** {row['備考']}
                    """)
                    c_ed, c_dl = st.columns(2)
                    if c_ed.button("編集", key=f"ast_edit_{row['asset_id']}"):
                        st.session_state.edit_asset_id = row['asset_id']
                        st.rerun()
                    if c_dl.button("削除", key=f"del_ast_{row['asset_id']}"):
                        if delete_data("assets", "asset_id", row['asset_id'], MAP_ASSETS):
                            st.rerun()

                # 編集フォーム（財産）
                if st.session_state.edit_asset_id == row['asset_id']:
                    st.markdown(f"#### ✏️ 編集: {row['名称・機関名']}")
                    with st.form("edit_asset_form"):
                        c1, c2 = st.columns(2)
                        try: idx = ast_opts.index(row['財産種別'])
                        except: idx = 0
                        ea_type = c1.selectbox("種別", ast_opts, index=idx)
                        ea_name = c2.text_input("名称", value=row['名称・機関名'])
                        c3, c4 = st.columns(2)
                        ea_det = c3.text_input("詳細", value=row['支店・詳細'])
                        ea_num = c4.text_input("口座番号等", value=row['口座番号・記号'])
                        ea_val = c1.text_input("評価額", value=row['評価額・残高'])
                        ea_loc = c2.text_input("保管場所", value=row['保管場所'])
                        ea_rem = st.text_area("備考", value=row['備考'])
                        
                        c_sv, c_cl = st.columns(2)
                        if c_sv.form_submit_button("保存"):
                            nd = {'財産種別': ea_type, '名称・機関名': ea_name, '支店・詳細': ea_det, '口座番号・記号': ea_num, '評価額・残高': ea_val, '保管場所': ea_loc, '備考': ea_rem}
                            if update_data("assets", "asset_id", st.session_state.edit_asset_id, nd, MAP_ASSETS):
                                st.session_state.edit_asset_id = None
                                st.rerun()
                        if c_cl.form_submit_button("キャンセル"):
                            st.session_state.edit_asset_id = None
                            st.rerun()
        else:
            st.info("登録された財産はありません。")

def render_person_registration(df_persons, guard_opts):
    custom_header("利用者情報登録")
    
    # 新規登録フォーム
    with st.expander("➕ 新規登録", expanded=False):
        with st.form("new_person"):
            col1, col2 = st.columns(2)
            p_case = col1.text_input("ケース番号")
            p_basic = col2.text_input("基本事件番号")
            
            col_nm, col_kn = st.columns(2)
            p_name = col_nm.text_input("氏名 (必須)")
            p_kana = col_kn.text_input("カナ")
            
            p_addr = st.text_input("住所")
            p_res = st.text_input("居所 (施設名など)")
            
            col_dob, col_typ = st.columns(2)
            p_dob = col_dob.date_input("生年月日", value=None, min_value=datetime.date(1900, 1, 1), format="YYYY/MM/DD")
            p_type = col_typ.selectbox("類型", guard_opts)
            
            if st.form_submit_button("登録"):
                if not p_name:
                    st.error("氏名は必須です")
                else:
                    nd = {
                        'ケース番号': p_case, '基本事件番号': p_basic, '氏名': p_name, 'ｼﾒｲ': p_kana,
                        '住所': p_addr, '居所': p_res, '生年月日': str(p_dob) if p_dob else None, 
                        '類型': p_type, '現在の状態': "受任中"
                    }
                    if insert_data("persons", nd, MAP_PERSONS):
                        st.rerun()
    
    if not df_persons.empty:
        st.markdown("### 登録済み一覧")
        display_cols = ['ケース番号', '氏名', '生年月日', '年齢', '現在の状態']
        df_display = df_persons[display_cols].copy()
        
        selection = st.dataframe(
            df_display,
            column_config={
                "ケース番号": st.column_config.TextColumn("No."),
                "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
            },
            use_container_width=True, on_select="rerun", selection_mode="single-row", hide_index=True
        )
        
        # 選択されたら編集フォーム表示
        if selection.selection.rows:
            idx = selection.selection.rows[0]
            edit_row = df_persons.iloc[idx]
            target_pid = edit_row['person_id']
            
            st.markdown("---")
            st.markdown(f"#### ✏️ {edit_row['氏名']} さんの情報を編集")
            
            with st.form(f"edit_person_full"):
                col1, col2 = st.columns(2)
                ep_case = col1.text_input("ケース番号", value=edit_row.get('ケース番号') or "")
                ep_basic = col2.text_input("基本事件番号", value=edit_row.get('基本事件番号') or "")
                
                col_nm, col_kn = st.columns(2)
                ep_name = col_nm.text_input("氏名", value=edit_row.get('氏名') or "")
                ep_kana = col_kn.text_input("カナ", value=edit_row.get('ｼﾒｲ') or "")
                
                col_dob, col_typ = st.columns(2)
                
                ep_dob_val = pd.to_datetime(edit_row['生年月日']).date() if pd.notnull(edit_row.get('生年月日')) and edit_row['生年月日'] else None
                ep_dob = col_dob.date_input("生年月日", value=ep_dob_val, min_value=datetime.date(1900, 1, 1), format="YYYY/MM/DD")
                
                try: g_idx = guard_opts.index(edit_row.get('類型'))
                except: g_idx = 0
                ep_type = col_typ.selectbox("類型", guard_opts, index=g_idx)
                
                ep_addr = st.text_input("住所", value=edit_row.get('住所') or "")
                ep_res = st.text_input("居所", value=edit_row.get('居所') or "")
                
                c_dis, c_pet = st.columns(2)
                ep_disability = c_dis.text_input("障害類型", value=edit_row.get('障害類型') or "")
                ep_petitioner = c_pet.text_input("申立人", value=edit_row.get('申立人') or "")
                
                c_jud, c_crt = st.columns(2)
                ep_judg_val = pd.to_datetime(edit_row['審判確定日']).date() if pd.notnull(edit_row.get('審判確定日')) and edit_row['審判確定日'] else None
                ep_judg = c_jud.date_input("審判確定日", value=ep_judg_val, min_value=datetime.date(1900, 1, 1), format="YYYY/MM/DD")
                ep_court = c_crt.text_input("管轄家裁", value=edit_row.get('管轄家裁') or "")
                
                c_rep, c_st = st.columns(2)
                ep_report = c_rep.text_input("家裁報告月", value=edit_row.get('家裁報告月') or "")
                
                try: s_idx = ["受任中", "終了"].index(edit_row.get('現在の状態'))
                except: s_idx = 0
                ep_stat = c_st.selectbox("状態", ["受任中", "終了"], index=s_idx)

                if st.form_submit_button("更新"):
                    upd_data = {
                        'ケース番号': ep_case, '基本事件番号': ep_basic, '氏名': ep_name, 'ｼﾒｲ': ep_kana,
                        '住所': ep_addr, '居所': ep_res,
                        '生年月日': str(ep_dob) if ep_dob else None, '類型': ep_type, '障害類型': ep_disability,
                        '申立人': ep_petitioner, '審判確定日': str(ep_judg) if ep_judg else None,
                        '管轄家裁': ep_court, '家裁報告月': ep_report, '現在の状態': ep_stat
                    }
                    if update_data("persons", "person_id", target_pid, upd_data, MAP_PERSONS):
                        st.rerun()

def render_reports(df_persons):
    custom_header("帳票作成")
    uploaded = st.file_uploader("Excelテンプレート")
    if not df_persons.empty:
        # 選択状態の維持
        current_name_idx = 0
        names = df_persons['氏名'].tolist()
        if st.session_state.selected_person_id:
            # IDから名前を取得してインデックスを探す
            row = df_persons[df_persons['person_id'] == st.session_state.selected_person_id]
            if not row.empty:
                current_name = row.iloc[0]['氏名']
                try:
                    current_name_idx = names.index(current_name)
                except ValueError:
                    current_name_idx = 0

        target = st.selectbox("対象者", names, index=current_name_idx)
        # ここで選択を変えた場合もセッションに反映するかは任意だが、統一感を出すなら反映する
        selected_row = df_persons[df_persons['氏名'] == target]
        if not selected_row.empty:
            pid = selected_row.iloc[0]['person_id']
            if pid != st.session_state.selected_person_id:
                st.session_state.selected_person_id = pid
        if st.button("作成") and uploaded:
            p_data = df_persons[df_persons['氏名'] == target].iloc[0].to_dict()
            excel = fill_excel_template(uploaded, p_data)
            st.download_button("ダウンロード", excel, f"{target}.xlsx")

def render_data_management():
    custom_header("データ管理")
    st.info("Supabaseへのデータ移行用です。")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["利用者", "活動", "財産", "関係者", "システム"])
    
    with tab1:
        csv_exp = fetch_table("persons", MAP_PERSONS).to_csv(index=False).encode('cp932')
        st.download_button("CSVエクスポート", csv_exp, "Persons.csv", "text/csv")
        up = st.file_uploader("インポート (Persons)")
        if up and st.button("実行", key="imp_p"):
            process_import(up, "persons", MAP_PERSONS, "person_id")

    with tab2:
        csv_exp = fetch_table("activities", MAP_ACTIVITIES).to_csv(index=False).encode('cp932')
        st.download_button("CSVエクスポート", csv_exp, "Activities.csv", "text/csv")
        up = st.file_uploader("インポート (Activities)")
        if up and st.button("実行", key="imp_a"):
            process_import(up, "activities", MAP_ACTIVITIES, "activity_id")
    
    with tab3:
        csv_exp = fetch_table("assets", MAP_ASSETS).to_csv(index=False).encode('cp932')
        st.download_button("CSVエクスポート", csv_exp, "Assets.csv", "text/csv")
        up = st.file_uploader("インポート (Assets)")
        if up and st.button("実行", key="imp_ast"):
            process_import(up, "assets", MAP_ASSETS, "asset_id")
    
    with tab4:
        csv_exp = fetch_table("related_parties", MAP_RELATED).to_csv(index=False).encode('cp932')
        st.download_button("CSVエクスポート", csv_exp, "RelatedParties.csv", "text/csv")
        up = st.file_uploader("インポート (Related)")
        if up and st.button("実行", key="imp_rel"):
            process_import(up, "related_parties", MAP_RELATED, "related_id")

    with tab5:
        csv_exp = fetch_table("app_system_user", MAP_SYSTEM).to_csv(index=False).encode('cp932')
        st.download_button("CSVエクスポート", csv_exp, "SystemUser.csv", "text/csv")
        up = st.file_uploader("インポート (SystemUser)")
        if up and st.button("実行", key="imp_sys"):
            process_import(up, "app_system_user", MAP_SYSTEM, "id")

def render_settings():
    custom_header("初期設定")
    
    st.markdown("#### マスタ管理 (選択肢の編集)")
    tabs_m = st.tabs(["活動種別", "財産種別", "関係種別", "後見類型"])
    
    master_cats = {
        "活動種別": "activity",
        "財産種別": "asset",
        "関係種別": "relationship",
        "後見類型": "guardian_type"
    }
    
    df_master = fetch_table("master_options", MAP_MASTER)
    
    for i, (label, cat_key) in enumerate(master_cats.items()):
        with tabs_m[i]:
            if not df_master.empty:
                current_opts = df_master[df_master['カテゴリ'] == cat_key].sort_values('順序')
                for _, row in current_opts.iterrows():
                    c1, c2 = st.columns([8, 2])
                    c1.write(f"{row['名称']} (順序:{row['順序']})")
                    if c2.button("削除", key=f"del_mst_{row['id']}"):
                        usage = check_usage_count(cat_key, row['名称'])
                        if usage > 0:
                            st.error(f"「{row['名称']}」は現在 {usage} 件のデータで使用されているため削除できません。")
                        else:
                            if delete_data("master_options", "id", row['id'], MAP_MASTER):
                                st.rerun()

            with st.form(f"add_mst_{cat_key}"):
                c_name = st.text_input("名称")
                c_order = st.number_input("順序", min_value=0, value=100)
                if st.form_submit_button("追加"):
                    if c_name:
                        if insert_data("master_options", {'カテゴリ': cat_key, '名称': c_name, '順序': c_order}, MAP_MASTER):
                            st.rerun()
    
    st.markdown("---")
    st.markdown("#### システム利用者情報")
    df_sys = fetch_table("app_system_user", MAP_SYSTEM)
    curr = df_sys.iloc[0].to_dict() if not df_sys.empty else {}
    
    with st.form("sys_user"):
        c1, c2 = st.columns(2)
        s_name = c1.text_input("氏名", value=curr.get('氏名', ''))
        s_kana = c2.text_input("カナ", value=curr.get('シメイ', ''))
        s_zip = c1.text_input("〒", value=curr.get('〒', ''))
        s_addr = c2.text_input("住所", value=curr.get('住所', ''))
        s_tel = st.text_input("電話", value=curr.get('連絡先電話番号', ''))
        s_mail = st.text_input("email", value=curr.get('e-mail', ''))
        if st.form_submit_button("保存"):
            nd = {'氏名': s_name, 'シメイ': s_kana, '〒': s_zip, '住所': s_addr, '連絡先電話番号': s_tel, 'e-mail': s_mail}
            if not df_sys.empty:
                if update_data("app_system_user", "id", curr['id'], nd, MAP_SYSTEM):
                    st.rerun()
            else:
                if insert_data("app_system_user", nd, MAP_SYSTEM):
                    st.rerun()
