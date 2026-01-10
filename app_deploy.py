import streamlit as st
import pandas as pd
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import re

# --- 設定・定数 ---
SPREADSHEET_NAME = '成年後見システム台帳'
KEY_FILE = 'credentials.json'

# --- 項目定義 ---
COL_DEF_PERSONS = [
    'person_id',
    'ケース番号',
    '基本事件番号',
    '氏名',
    'ｼﾒｲ',
    '生年月日',
    '類型',
    '障害類型',
    '申立人',
    '審判確定日',
    '管轄家裁',
    '家裁報告月',
    '現在の状態'
]

# 次回予定日を削除
COL_DEF_ACTIVITIES = ['activity_id', 'person_id', '記録日', '活動', '要点', '作成日時']

st.set_page_config(page_title="成年後見業務支援システム", layout="wide")

# --- CSS (デザイン調整・スマホ最適化) ---
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: "Noto Sans JP", sans-serif;
        color: #333333;
    }
    
    /* --- 全体的な余白の削減（スマホ最適化） --- */
    
    /* メインエリアの上部余白を削減 */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    
    /* 要素間の隙間（ギャップ）を詰める */
    div[data-testid="stVerticalBlock"] {
        gap: 0.5rem !important;
    }
    
    /* 各要素のコンテナ余白を削減 */
    div[data-testid="stElementContainer"] {
        margin-bottom: 0.2rem !important;
    }

    /* テーブルの行間を狭く */
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
        padding-top: 2px !important;
        padding-bottom: 2px !important;
        font-size: 13px !important;
    }
    
    /* 基本情報の表示行間を狭くする */
    div[data-testid="stExpander"] .stMarkdown p {
        margin-bottom: 0px !important;
        line-height: 1.4 !important;
    }
    
    /* タイトルスタイル */
    .custom-title {
        font-size: 20px !important;
        font-weight: bold !important;
        color: #006633 !important;
        border-left: 6px solid #006633;
        padding-left: 10px;
        margin-top: 5px;
        margin-bottom: 10px;
        background-color: #f8f9fa;
        padding: 5px;
    }
    
    /* 見出しスタイル */
    .custom-header {
        font-size: 16px !important;
        font-weight: bold !important;
        color: #006633 !important;
        border-bottom: 1px solid #ccc;
        padding-bottom: 2px;
        margin-top: 15px;
        margin-bottom: 5px;
    }

    /* 見出しテキスト（ボタン横並び用） */
    .custom-header-text {
        font-size: 16px !important;
        font-weight: bold !important;
        color: #006633 !important;
        margin: 0 !important;
        padding-top: 5px; /* ボタンの高さに合わせる */
        white-space: nowrap;
    }
    /* 分離した下線 */
    .custom-header-line {
        border-bottom: 1px solid #ccc;
        margin-top: 0px;
        margin-bottom: 5px;
    }
    
    /* 入力フォームのデザイン調整（角丸・余白削減） */
    .stTextInput input, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea {
        border: 1px solid #666 !important;
        background-color: #ffffff !important;
        border-radius: 6px !important;
        padding: 4px 8px !important;
        min-height: 0px !important;
    }
    /* ラベルの余白も詰める */
    .stTextInput label, .stSelectbox label, .stDateInput label, .stTextArea label {
        margin-bottom: 0px !important;
        font-size: 13px !important;
    }
    
    /* ヘルプボタン（ポップオーバー）の微調整 */
    div[data-testid="stPopover"] button {
        padding: 0px 8px !important;
        height: auto !important;
        min-height: 0px !important;
        border: 1px solid #ccc !important;
    }

    /* --- ファイルアップローダー --- */
    [data-testid="stFileUploaderDropzone"] div div span, [data-testid="stFileUploaderDropzone"] div div small {
        display: none;
    }
    [data-testid="stFileUploaderDropzone"] div div::after {
        content: "ファイルをドラッグ＆ドロップまたは選択";
        font-size: 12px;
        font-weight: bold;
        color: #333;
        display: block;
        margin: 5px 0;
    }
    [data-testid="stFileUploaderDropzone"] div div::before {
        content: "CSVファイル (200MBまで)";
        font-size: 12px;
        color: #666;
        display: block;
        margin-bottom: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 認証機能 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if st.session_state.password_correct:
        return True
    st.markdown("## 🔒 ログイン")
    password = st.text_input("パスワードを入力してください", type="password")
    if st.button("ログイン"):
        correct_password = "admin" 
        try:
            if "APP_PASSWORD" in st.secrets:
                correct_password = st.secrets["APP_PASSWORD"]
        except:
            pass
        if password == correct_password:
            st.session_state.password_correct = True
            st.success("ログインしました")
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

# --- Google接続関数 ---
def get_spreadsheet_connection():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    except:
        pass
    if creds is None:
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
        except Exception as e:
            return f"鍵ファイルが見つかりません。({str(e)})"
    try:
        client = gspread.authorize(creds)
        sheet = client.open(SPREADSHEET_NAME)
        return sheet
    except Exception as e:
        return str(e)

# --- ユーティリティ関数 ---
def normalize_date_str(date_val):
    if date_val is None: return ""
    text = str(date_val).strip()
    if not text or text.lower() == "nan": return ""
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    eras = {'明治': 1868, '大正': 1912, '昭和': 1926, '平成': 1989, '令和': 2019,
            'M': 1868, 'T': 1912, 'S': 1926, 'H': 1989, 'R': 2019}
    match = re.match(r'([明治大正昭和平成令和MTSHR])\s*(\d+)\D+(\d+)\D+(\d+)', text, re.IGNORECASE)
    if match:
        era_str, year_str, month_str, day_str = match.groups()
        era_str = era_str.upper()
        base_year = 1900
        for k, v in eras.items():
            if k == era_str:
                base_year = v
                break
        year = int(year_str)
        west_year = base_year + year - 1 if year > 0 else base_year
        return f"{west_year}-{int(month_str):02d}-{int(day_str):02d}"
    try:
        dt = pd.to_datetime(text, errors='coerce')
        if pd.isna(dt): return text
        return dt.strftime('%Y-%m-%d')
    except:
        return text

def calculate_age(born):
    if not born: return None
    born_str = str(born).strip()
    if not born_str or born_str.lower() == 'nan': return None
    try:
        born_date = pd.to_datetime(born_str, errors='coerce')
        if pd.isna(born_date): return None
        born_date = born_date.date()
        today = datetime.date.today()
        return today.year - born_date.year - ((today.month, today.day) < (born.month, born.day))
    except:
        return None

def load_data_from_sheet(sheet):
    try:
        ws_persons = sheet.worksheet("Persons")
    except:
        ws_persons = sheet.add_worksheet(title="Persons", rows="100", cols="20")
        ws_persons.append_row(COL_DEF_PERSONS)
    try:
        ws_activities = sheet.worksheet("Activities")
    except:
        ws_activities = sheet.add_worksheet(title="Activities", rows="1000", cols="20")
        ws_activities.append_row(COL_DEF_ACTIVITIES)
    
    df_persons = pd.DataFrame(ws_persons.get_all_records())
    df_activities = pd.DataFrame(ws_activities.get_all_records())

    for col in COL_DEF_PERSONS:
        if col not in df_persons.columns: df_persons[col] = ""
    for col in COL_DEF_ACTIVITIES:
        if col not in df_activities.columns: df_activities[col] = ""

    for col in ['生年月日', '審判確定日']:
        if col in df_persons.columns:
            df_persons[col] = df_persons[col].apply(normalize_date_str)
    for col in ['記録日']:
        if col in df_activities.columns:
            df_activities[col] = df_activities[col].apply(normalize_date_str)

    return df_persons, df_activities

def add_data_to_sheet(sheet_name, new_row_list):
    sheet = get_spreadsheet_connection()
    worksheet = sheet.worksheet(sheet_name)
    worksheet.append_row(new_row_list)

def update_sheet_data(sheet_name, id_column, target_id, update_dict):
    sheet = get_spreadsheet_connection()
    if isinstance(sheet, str):
        st.error(f"接続エラー: {sheet}")
        return False
    worksheet = sheet.worksheet(sheet_name)
    header_cells = worksheet.row_values(1)
    try:
        pid_col_index = header_cells.index(id_column) + 1
    except ValueError:
        st.error(f"システムエラー: {id_column} 列が見つかりません。")
        return False
    all_ids = worksheet.col_values(pid_col_index)
    target_row_num = -1
    str_search_id = str(target_id)
    for i, val in enumerate(all_ids):
        if str(val) == str_search_id:
            target_row_num = i + 1
            break
    if target_row_num == -1:
        st.error(f"更新対象のID ({target_id}) が見つかりませんでした。")
        return False
    try:
        cells_to_update = []
        for col_name, value in update_dict.items():
            if col_name in header_cells:
                col_num = header_cells.index(col_name) + 1
                cells_to_update.append(gspread.Cell(target_row_num, col_num, str(value)))
        if cells_to_update:
            worksheet.update_cells(cells_to_update)
            st.toast("情報を更新しました", icon="✅")
            return True
        return False
    except Exception as e:
        st.error(f"更新エラー: {str(e)}")
        return False

def import_csv_to_sheet_safe(sheet_name, df_upload, target_columns, id_column, date_columns=[]):
    sheet = get_spreadsheet_connection()
    worksheet = sheet.worksheet(sheet_name)
    existing_records = worksheet.get_all_records()
    df_existing = pd.DataFrame(existing_records)
    existing_ids = set()
    if not df_existing.empty and id_column in df_existing.columns:
        existing_ids = set(df_existing[id_column].astype(str))
    export_data = []
    skipped_count = 0
    for index, row in df_upload.iterrows():
        if id_column in row and str(row[id_column]) in existing_ids:
            skipped_count += 1
            continue
        new_row = []
        for col in target_columns:
            val = ""
            if col in row:
                raw_val = row[col]
                if not pd.isna(raw_val):
                    if col in date_columns:
                        val = normalize_date_str(raw_val)
                    else:
                        val = str(raw_val)
            new_row.append(val)
        export_data.append(new_row)
    if export_data:
        worksheet.append_rows(export_data)
        return len(export_data), skipped_count
    return 0, skipped_count

def custom_title(text):
    st.markdown(f'<div class="custom-title">{text}</div>', unsafe_allow_html=True)

# --- カスタムヘッダー関数（スマホ配置修正版） ---
def custom_header(text, help_text=None):
    if help_text:
        # スマホで崩れないよう、比率を調整し余白カラムを削除
        col1, col2 = st.columns([9, 1], gap="small")
        with col1:
            st.markdown(f'<div class="custom-header-text">{text}</div>', unsafe_allow_html=True)
        with col2:
            with st.popover("?", use_container_width=True):
                st.info(help_text)
        st.markdown('<div class="custom-header-line"></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="custom-header">{text}</div>', unsafe_allow_html=True)

# --- メイン処理 ---
def main():
    if not check_password(): return
    custom_title("成年後見業務支援システム")

    sheet_connection = get_spreadsheet_connection()
    if isinstance(sheet_connection, str):
        st.error(f"接続エラー: {sheet_connection}")
        return

    df_persons, df_activities = load_data_from_sheet(sheet_connection)

    if '生年月日' in df_persons.columns:
        if not df_persons.empty:
            df_persons['年齢'] = df_persons['生年月日'].apply(calculate_age)
            df_persons['年齢'] = pd.to_numeric(df_persons['年齢'], errors='coerce')
        else:
            df_persons['年齢'] = None

    # メニュー名変更
    menu = st.sidebar.radio("メニュー", ["利用者基本情報・活動記録", "基本情報登録", "データ管理・移行"])

    # ステート管理
    if 'selected_person_id' not in st.session_state:
        st.session_state.selected_person_id = None

    # =========================================================
    # 1. 利用者基本情報・活動記録
    # =========================================================
    if menu == "利用者基本情報・活動記録":
        custom_header("受任中利用者一覧", help_text="一覧から対象者をクリックすると詳細が表示されます。")
        
        if not df_persons.empty and '現在の状態' in df_persons.columns:
            mask = df_persons['現在の状態'].fillna('').astype(str).isin(['受任中', '', 'nan'])
            df_active = df_persons[mask].copy()
        else:
            df_active = df_persons.copy()

        display_columns = ['ケース番号', '氏名', '生年月日', '年齢', '類型']
        available_cols = [c for c in display_columns if c in df_active.columns]
        df_display = df_active[available_cols] if not df_active.empty and len(available_cols) > 0 else pd.DataFrame(columns=display_columns)

        if '年齢' in df_display.columns:
            df_display['年齢'] = pd.to_numeric(df_display['年齢'], errors='coerce')

        selection = st.dataframe(
            df_display, 
            column_config={
                "ケース番号": st.column_config.TextColumn("No."),
                "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
                "類型": st.column_config.TextColumn("後見類型"),
            },
            use_container_width=True,
            on_select="rerun", 
            selection_mode="single-row", 
            hide_index=True
        )
        
        if selection.selection.rows:
            idx = selection.selection.rows[0]
            selected_row = df_active.iloc[idx]
            current_person_id = selected_row['person_id']
            st.session_state.selected_person_id = current_person_id
            
            st.markdown("---")
            age_val = selected_row.get('年齢')
            age_str = f" ({int(age_val)}歳)" if (age_val is not None and not pd.isna(age_val)) else ""
            custom_header(f"{selected_row.get('氏名', '名称不明')}{age_str} さんの詳細・活動記録")

            with st.expander("▼ 基本情報", expanded=False):
                c1, c2, c3 = st.columns(3)
                c1.markdown(f"**No.:** {selected_row.get('ケース番号', '')}")
                c2.markdown(f"**事件番号:** {selected_row.get('基本事件番号', '')}")
                c3.markdown(f"**類型:** {selected_row.get('類型', '')}")
                c4, c5, c6 = st.columns(3)
                c4.markdown(f"**氏名:** {selected_row.get('氏名', '')}")
                c5.markdown(f"**ｼﾒｲ:** {selected_row.get('ｼﾒｲ', '')}")
                c6.markdown(f"**生年月日:** {selected_row.get('生年月日', '')}")
                c7, c8, c9 = st.columns(3)
                c7.markdown(f"**障害類型:** {selected_row.get('障害類型', '')}")
                c8.markdown(f"**申立人:** {selected_row.get('申立人', '')}")
                c9.markdown(f"**審判日:** {selected_row.get('審判確定日', '')}")
                c10, c11, c12 = st.columns(3)
                c10.markdown(f"**家裁:** {selected_row.get('管轄家裁', '')}")
                c11.markdown(f"**報告月:** {selected_row.get('家裁報告月', '')}")
                c12.markdown(f"**状態:** {selected_row.get('現在の状態', '')}")

            st.markdown("### 📝 活動記録の入力")
            with st.container(border=True):
                with st.form("new_activity_form"):
                    col_a, col_b = st.columns(2)
                    input_date = col_a.date_input("活動日", value=datetime.date.today(), min_value=datetime.date(2000, 1, 1))
                    activity_opts = ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"]
                    input_activity = col_b.selectbox("活動", activity_opts)
                    input_summary = st.text_area("要点・内容", height=80) # 高さを少し減らして省スペース化
                    
                    if st.form_submit_button("登録"):
                        new_id = 1
                        if len(df_activities) > 0:
                            try: new_id = pd.to_numeric(df_activities['activity_id']).max() + 1
                            except: pass
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        new_row = [int(new_id), int(current_person_id), str(input_date), input_activity, input_summary, now_str]
                        add_data_to_sheet("Activities", new_row)
                        st.rerun()

            custom_header("過去の活動履歴", help_text="カードの右下にある「編集」ボタンで内容を修正できます。")
            if 'edit_activity_id' not in st.session_state:
                st.session_state.edit_activity_id = None

            try:
                df_activities['person_id'] = pd.to_numeric(df_activities['person_id'], errors='coerce')
                my_activities = df_activities[df_activities['person_id'] == int(current_person_id)].copy()
                
                if not my_activities.empty:
                    my_activities = my_activities.sort_values('記録日', ascending=False)
                    
                    # 編集フォーム
                    if st.session_state.edit_activity_id:
                        edit_row = my_activities[my_activities['activity_id'] == st.session_state.edit_activity_id].iloc[0]
                        with st.container(border=True):
                            st.markdown(f"#### ✏️ 活動履歴の修正 (ID: {edit_row['activity_id']})")
                            with st.form("edit_activity_form"):
                                ea_date_val = pd.to_datetime(edit_row['記録日']).date() if edit_row['記録日'] else None
                                ea_date = st.date_input("活動日", value=ea_date_val, min_value=datetime.date(2000, 1, 1))
                                curr_act = edit_row['活動'] if edit_row['活動'] in ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"] else "その他"
                                ea_act = st.selectbox("活動", ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"], index=["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"].index(curr_act))
                                ea_summary = st.text_area("要点・内容", value=edit_row['要点'], height=100)
                                c_save, c_cancel = st.columns(2)
                                with c_save:
                                    if st.form_submit_button("保存"):
                                        upd_dict = {'記録日': str(ea_date), '活動': ea_act, '要点': ea_summary}
                                        if update_sheet_data("Activities", "activity_id", st.session_state.edit_activity_id, upd_dict):
                                            st.session_state.edit_activity_id = None
                                            st.rerun()
                                with c_cancel:
                                    if st.form_submit_button("キャンセル"):
                                        st.session_state.edit_activity_id = None
                                        st.rerun()

                    # 一覧表示（カード形式 + 右下ボタン）
                    for idx, row in my_activities.iterrows():
                        with st.container(border=True):
                            # 上段
                            c_date, c_act = st.columns([1, 2])
                            with c_date: st.write(f"📅 **{row['記録日']}**")
                            with c_act: st.write(f"📝 **{row['活動']}**")
                            
                            # 中段（要点）
                            st.write(row['要点'])
                            
                            # 下段（右寄せ編集ボタン）
                            # 左に空白カラム(8割)、右にボタンカラム(2割)
                            c_void, c_btn = st.columns([8, 2]) 
                            with c_btn:
                                if st.button("編集", key=f"btn_edit_{row['activity_id']}", use_container_width=True):
                                    st.session_state.edit_activity_id = row['activity_id']
                                    st.rerun()
                else:
                    st.write("まだ記録がありません。")
            except Exception as e:
                st.write(f"読込エラー: {e}")

    # =========================================================
    # 2. 基本情報登録
    # =========================================================
    elif menu == "基本情報登録":
        custom_header("基本情報登録", help_text="新規登録の場合はフォームに入力してください。\n修正の場合は、下の一覧から対象者をクリックしてください。")
        
        if 'edit_person_id' not in st.session_state:
            st.session_state.edit_person_id = None
        
        st.markdown("### 全利用者一覧")
        
        reg_list_cols = ['ケース番号', '氏名', '生年月日', '年齢', '現在の状態']
        available_reg_cols = [c for c in reg_list_cols if c in df_persons.columns]
        df_display_reg = df_persons[available_reg_cols] if not df_persons.empty and len(available_reg_cols) > 0 else pd.DataFrame(columns=reg_list_cols)
        
        if not df_display_reg.empty and '年齢' in df_display_reg.columns:
            df_display_reg['年齢'] = pd.to_numeric(df_display_reg['年齢'], errors='coerce')

        selection_reg = st.dataframe(
            df_display_reg,
            column_config={
                "ケース番号": st.column_config.TextColumn("No."),
                "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
            },
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            hide_index=True,
            height=200
        )
        
        selected_data = {}
        is_edit_mode = False
        
        if selection_reg.selection.rows:
            idx = selection_reg.selection.rows[0]
            full_row = df_persons.iloc[idx]
            st.session_state.edit_person_id = full_row['person_id']
            selected_data = full_row.to_dict()
            is_edit_mode = True
            st.markdown(f"### ✏️ 編集モード: {selected_data.get('氏名', '')}")
            if st.button("選択を解除"):
                st.session_state.edit_person_id = None
                st.rerun()
        else:
            st.markdown("### ✨ 新規登録モード")

        with st.form("person_info_form"):
            col1, col2 = st.columns(2)
            val_case_no = selected_data.get('ケース番号', '')
            val_basic_no = selected_data.get('基本事件番号', '')
            val_name = selected_data.get('氏名', '')
            val_kana = selected_data.get('ｼﾒｲ', '')
            type_options = ["後見", "保佐", "補助", "任意", "未成年後見", "その他"]
            val_type_raw = selected_data.get('類型', '後見')
            val_type_index = type_options.index(val_type_raw) if val_type_raw in type_options else 0
            val_disability = selected_data.get('障害類型', '')
            val_petitioner = selected_data.get('申立人', '')
            val_court = selected_data.get('管轄家裁', '')
            val_report_month = selected_data.get('家裁報告月', '')
            status_options = ["受任中", "終了"]
            val_status_raw = selected_data.get('現在の状態', '受任中')
            val_status_index = status_options.index(val_status_raw) if val_status_raw in status_options else 0
            val_dob = pd.to_datetime(selected_data.get('生年月日')).date() if selected_data.get('生年月日') else None
            val_ref_date = pd.to_datetime(selected_data.get('審判確定日')).date() if selected_data.get('審判確定日') else None

            in_case_no = col1.text_input("ケース番号", value=val_case_no)
            in_basic_no = col2.text_input("基本事件番号", value=val_basic_no)
            in_name = col1.text_input("氏名 (必須)", value=val_name)
            in_kana = col2.text_input("ｼﾒｲ (カナ)", value=val_kana)
            in_dob = col1.date_input("生年月日", value=val_dob, min_value=datetime.date(1900, 1, 1))
            in_type = col2.selectbox("類型", type_options, index=val_type_index)
            in_disability = col1.text_input("障害類型", value=val_disability)
            in_petitioner = col2.text_input("申立人", value=val_petitioner)
            in_ref_date = col1.date_input("審判確定日", value=val_ref_date, min_value=datetime.date(2000, 1, 1))
            in_court = col2.text_input("管轄家裁", value=val_court)
            in_report_month = col1.text_input("家裁報告月", value=val_report_month)
            in_status = col2.selectbox("現在の状態", status_options, index=val_status_index)

            if st.form_submit_button("情報を更新する" if is_edit_mode else "新規登録する"):
                if not in_name:
                    st.error("氏名は必須です。")
                else:
                    update_data = {
                        'ケース番号': in_case_no, '基本事件番号': in_basic_no,
                        '氏名': in_name, 'ｼﾒｲ': in_kana,
                        '生年月日': str(in_dob) if in_dob else "",
                        '類型': in_type, '障害類型': in_disability,
                        '申立人': in_petitioner,
                        '審判確定日': str(in_ref_date) if in_ref_date else "",
                        '管轄家裁': in_court, '家裁報告月': in_report_month,
                        '現在の状態': in_status
                    }
                    if is_edit_mode:
                        if update_sheet_data("Persons", "person_id", st.session_state.edit_person_id, update_data):
                            st.session_state.edit_person_id = None
                            st.rerun()
                    else:
                        new_pid = 1
                        if len(df_persons) > 0:
                            try: new_pid = pd.to_numeric(df_persons['person_id']).max() + 1
                            except: pass
                        new_row = [int(new_pid), in_case_no, in_basic_no, in_name, in_kana,
                                   str(in_dob) if in_dob else "", in_type, in_disability, in_petitioner,
                                   str(in_ref_date) if in_ref_date else "", in_court, in_report_month, in_status]
                        add_data_to_sheet("Persons", new_row)
                        st.success(f"{in_name} さんを新規登録しました。")
                        st.rerun()

    # =========================================================
    # 3. データ管理・移行
    # =========================================================
    elif menu == "データ管理・移行":
        custom_header("データ一括インポート・エクスポート", help_text="指定のCSV様式を使って、データの一括登録やバックアップができます。")
        st.markdown("データのバックアップ（エクスポート）や、CSVファイルによる一括取り込みができます。")

        tab1, tab2 = st.tabs(["1. 利用者データ (Persons)", "2. 活動記録データ (Activities)"])

        with tab1:
            st.subheader("利用者データの管理")
            st.markdown("#### 📤 データエクスポート")
            csv_exp_p = df_persons.to_csv(index=False).encode('cp932')
            st.download_button("現在のデータをダウンロード (Persons_Export.csv)", csv_exp_p, "Persons_Export.csv", "text/csv")
            st.markdown("---")
            st.markdown("#### 📥 データインポート")
            df_template_p = pd.DataFrame(columns=COL_DEF_PERSONS)
            csv_template_p = df_template_p.to_csv(index=False).encode('cp932')
            st.download_button("空の様式をダウンロード (Persons_Template.csv)", csv_template_p, "Persons_Template.csv", "text/csv")
            uploaded_file_p = st.file_uploader("CSVアップロード", type=["csv"], key="upload_p")
            if uploaded_file_p:
                try:
                    try: df_upload_p = pd.read_csv(uploaded_file_p)
                    except: 
                        uploaded_file_p.seek(0)
                        df_upload_p = pd.read_csv(uploaded_file_p, encoding='cp932')
                    st.write(df_upload_p.head())
                    if st.button("取り込み (Persons)", key="btn_imp_p"):
                        date_columns = ['生年月日', '審判確定日']
                        count, skipped = import_csv_to_sheet_safe("Persons", df_upload_p, COL_DEF_PERSONS, "person_id", date_columns)
                        st.success(f"{count} 件追加しました。（重複スキップ: {skipped} 件）")
                except Exception as e: st.error(f"エラー: {e}")

        with tab2:
            st.subheader("活動記録データの管理")
            st.markdown("#### 📤 データエクスポート")
            csv_exp_a = df_activities.to_csv(index=False).encode('cp932')
            st.download_button("現在のデータをダウンロード (Activities_Export.csv)", csv_exp_a, "Activities_Export.csv", "text/csv")
            st.markdown("---")
            st.markdown("#### 📥 データインポート")
            df_template_a = pd.DataFrame(columns=COL_DEF_ACTIVITIES)
            csv_template_a = df_template_a.to_csv(index=False).encode('cp932')
            st.download_button("空の様式をダウンロード (Activities_Template.csv)", csv_template_a, "Activities_Template.csv", "text/csv")
            uploaded_file_a = st.file_uploader("CSVアップロード", type=["csv"], key="upload_a")
            if uploaded_file_a:
                try:
                    try: df_upload_a = pd.read_csv(uploaded_file_a)
                    except: 
                        uploaded_file_a.seek(0)
                        df_upload_a = pd.read_csv(uploaded_file_a, encoding='cp932')
                    st.write(df_upload_a.head())
                    if st.button("取り込み (Activities)", key="btn_imp_a"):
                        date_columns = ['記録日']
                        count, skipped = import_csv_to_sheet_safe("Activities", df_upload_a, COL_DEF_ACTIVITIES, "activity_id", date_columns)
                        st.success(f"{count} 件追加しました。（重複スキップ: {skipped} 件）")
                except Exception as e: st.error(f"エラー: {e}")

if __name__ == "__main__":
    main()