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

# 新しい基本情報の項目定義
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

COL_DEF_ACTIVITIES = ['activity_id', 'person_id', '記録日', '手段', '要点', '次回予定日', '作成日時']

st.set_page_config(page_title="成年後見業務支援システム", layout="wide")

# --- デザイン調整用CSS ---
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: "Noto Sans JP", sans-serif;
        color: #333333;
    }
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
        padding-top: 4px !important;
        padding-bottom: 4px !important;
        font-size: 14px !important;
    }
    .custom-title {
        font-size: 22px !important;
        font-weight: bold !important;
        color: #006633 !important;
        border-left: 6px solid #006633;
        padding-left: 12px;
        margin-top: 10px;
        margin-bottom: 20px;
        background-color: #f8f9fa;
        padding: 5px;
    }
    .custom-header {
        font-size: 18px !important;
        font-weight: bold !important;
        color: #006633 !important;
        margin-top: 25px;
        margin-bottom: 10px;
        border-bottom: 1px solid #ccc;
        padding-bottom: 5px;
    }
    .stTextInput > div > div > input {
        border: 1px solid #666 !important;
        background-color: #ffffff !important;
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

# --- 関数定義 ---
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

# --- 日付正規化・和暦対応ロジック ---
def normalize_date_str(date_val):
    if date_val is None:
        return ""
    text = str(date_val).strip()
    if not text or text.lower() == "nan":
        return ""
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
        if year == 1:
            west_year = base_year
        else:
            west_year = base_year + year - 1
        return f"{west_year}-{int(month_str):02d}-{int(day_str):02d}"
    try:
        dt = pd.to_datetime(text, errors='coerce')
        if pd.isna(dt):
            return text
        return dt.strftime('%Y-%m-%d')
    except:
        return text

def calculate_age(born):
    if not born:
        return ""
    try:
        born_date = pd.to_datetime(born, errors='coerce')
        if pd.isna(born_date):
            return ""
        born_date = born_date.date()
        today = datetime.date.today()
        age = today.year - born_date.year - ((today.month, today.day) < (born_date.month, born_date.day))
        return age
    except:
        return ""

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
    
    data_persons = ws_persons.get_all_records()
    df_persons = pd.DataFrame(data_persons)
    
    data_activities = ws_activities.get_all_records()
    df_activities = pd.DataFrame(data_activities)

    for col in COL_DEF_PERSONS:
        if col not in df_persons.columns:
            df_persons[col] = ""

    for col in COL_DEF_ACTIVITIES:
        if col not in df_activities.columns:
            df_activities[col] = ""

    date_cols_p = ['生年月日', '審判確定日']
    for col in date_cols_p:
        if col in df_persons.columns:
            df_persons[col] = df_persons[col].apply(normalize_date_str)

    date_cols_a = ['記録日', '次回予定日']
    for col in date_cols_a:
        if col in df_activities.columns:
            df_activities[col] = df_activities[col].apply(normalize_date_str)

    return df_persons, df_activities

def add_data_to_sheet(sheet_name, new_row_list):
    sheet = get_spreadsheet_connection()
    worksheet = sheet.worksheet(sheet_name)
    worksheet.append_row(new_row_list)

def update_person_data(person_id, update_dict):
    """
    利用者情報を更新する関数（修正版）
    行ズレを防ぐため、スプレッドシート上のIDを直接検索して更新します。
    """
    sheet = get_spreadsheet_connection()
    if isinstance(sheet, str):
        st.error(f"接続エラー: {sheet}")
        return False
        
    worksheet = sheet.worksheet("Persons")
    
    # ヘッダー行を取得して列の位置を把握
    header_cells = worksheet.row_values(1)
    
    # person_id 列がどこにあるか探す
    try:
        pid_col_index = header_cells.index('person_id') + 1
    except ValueError:
        st.error("システムエラー: person_id 列が見つかりません。")
        return False
    
    # person_id 列の値を全て取得して、更新対象の行番号を探す
    # (APIコール節約のため col_values を使用)
    all_ids = worksheet.col_values(pid_col_index)
    
    target_row_num = -1
    str_search_id = str(person_id)
    
    # リストの中からIDが一致する行を探す
    for i, val in enumerate(all_ids):
        if str(val) == str_search_id:
            target_row_num = i + 1 # リストは0始まり、行番号は1始まり
            break
            
    if target_row_num == -1:
        st.error(f"更新対象のID ({person_id}) が見つかりませんでした。")
        return False
        
    # 更新実行
    try:
        # 安全のため、複数のセルをまとめて更新する
        cells_to_update = []
        for col_name, value in update_dict.items():
            if col_name in header_cells:
                col_num = header_cells.index(col_name) + 1
                # 値をセット（全て文字列として扱う）
                cells_to_update.append(gspread.Cell(target_row_num, col_num, str(value)))
        
        if cells_to_update:
            worksheet.update_cells(cells_to_update)
            st.toast("情報を更新しました", icon="✅")
            return True
        else:
            return False
            
    except Exception as e:
        st.error(f"更新エラー: {str(e)}")
        return False

def import_csv_to_sheet(sheet_name, df_upload, target_columns, date_columns=[]):
    sheet = get_spreadsheet_connection()
    worksheet = sheet.worksheet(sheet_name)
    export_data = []
    for index, row in df_upload.iterrows():
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
        return len(export_data)
    return 0

def custom_title(text):
    st.markdown(f'<div class="custom-title">{text}</div>', unsafe_allow_html=True)

def custom_header(text):
    st.markdown(f'<div class="custom-header">{text}</div>', unsafe_allow_html=True)

# --- メイン処理 ---
def main():
    if not check_password():
        return
    
    custom_title("成年後見業務支援システム")

    sheet_connection = get_spreadsheet_connection()
    if isinstance(sheet_connection, str):
        st.error(f"接続エラー: {sheet_connection}")
        return

    df_persons, df_activities = load_data_from_sheet(sheet_connection)

    # 年齢計算
    if '生年月日' in df_persons.columns:
        if not df_persons.empty:
            df_persons['年齢'] = df_persons['生年月日'].apply(calculate_age)
        else:
            df_persons['年齢'] = None

    menu = st.sidebar.radio("メニュー", ["利用者一覧・活動記録", "基本情報登録", "データ管理・移行"])

    # --- 画面1: 利用者一覧・活動記録 (受任中のみ表示) ---
    if menu == "利用者一覧・活動記録":
        custom_header("受任中利用者一覧")
        
        # フィルタリング: 現在の状態が「受任中」または「空欄」
        if not df_persons.empty and '現在の状態' in df_persons.columns:
            mask = df_persons['現在の状態'].fillna('').astype(str).isin(['受任中', '', 'nan'])
            df_active = df_persons[mask].copy()
        else:
            df_active = df_persons.copy()

        # 表示カラム: ケース番号, 氏名, 生年月日, 年齢, 後見類型
        display_columns = ['ケース番号', '氏名', '生年月日', '年齢', '類型']
        available_cols = [c for c in display_columns if c in df_active.columns]
        
        if not df_active.empty and len(available_cols) > 0:
            df_display = df_active[available_cols]
        else:
            df_display = pd.DataFrame(columns=display_columns)

        selection = st.dataframe(
            df_display, 
            column_config={
                "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
                "類型": st.column_config.TextColumn("後見類型"),
            },
            use_container_width=False,
            on_select="rerun", 
            selection_mode="single-row", 
            hide_index=True
        )
        
        if selection.selection.rows:
            selected_row_index = selection.selection.rows[0]
            selected_row = df_active.iloc[selected_row_index]
            current_person_id = selected_row['person_id']
            
            st.markdown("---")
            age_val = selected_row.get('年齢')
            age_str = f" ({int(age_val)}歳)" if (age_val is not None and not pd.isna(age_val) and age_val != "") else ""
            
            custom_header(f"{selected_row.get('氏名', '名称不明')}{age_str} さんの詳細・活動記録")

            with st.expander("▼ 基本情報を全て表示", expanded=True):
                c1, c2, c3 = st.columns(3)
                c1.markdown(f"**ケース番号:** {selected_row.get('ケース番号', '')}")
                c2.markdown(f"**基本事件番号:** {selected_row.get('基本事件番号', '')}")
                c3.markdown(f"**類型:** {selected_row.get('類型', '')}")
                
                c4, c5, c6 = st.columns(3)
                c4.markdown(f"**氏名:** {selected_row.get('氏名', '')}")
                c5.markdown(f"**ｼﾒｲ:** {selected_row.get('ｼﾒｲ', '')}")
                c6.markdown(f"**生年月日:** {selected_row.get('生年月日', '')}")
                
                c7, c8, c9 = st.columns(3)
                c7.markdown(f"**障害類型:** {selected_row.get('障害類型', '')}")
                c8.markdown(f"**申立人:** {selected_row.get('申立人', '')}")
                c9.markdown(f"**審判確定日:** {selected_row.get('審判確定日', '')}")

                c10, c11, c12 = st.columns(3)
                c10.markdown(f"**管轄家裁:** {selected_row.get('管轄家裁', '')}")
                c11.markdown(f"**家裁報告月:** {selected_row.get('家裁報告月', '')}")
                c12.markdown(f"**現在の状態:** {selected_row.get('現在の状態', '')}")

            st.markdown("### 📝 活動記録の入力")
            with st.container(border=True):
                with st.form("new_activity_form"):
                    col_a, col_b = st.columns(2)
                    input_date = col_a.date_input("記録日", datetime.date.today())
                    input_method = col_b.selectbox("手段", ["訪問", "電話", "メール", "面談", "その他"])
                    
                    input_summary = st.text_area("要点・内容", height=100)
                    input_next_date = st.date_input("次回予定日", datetime.date.today() + datetime.timedelta(days=30))
                    
                    if st.form_submit_button("登録してクラウドへ送信"):
                        new_id = 1
                        if len(df_activities) > 0:
                            try: new_id = pd.to_numeric(df_activities['activity_id']).max() + 1
                            except: pass
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        new_row = [int(new_id), int(current_person_id), str(input_date), input_method, input_summary, str(input_next_date), now_str]
                        add_data_to_sheet("Activities", new_row)
                        st.rerun()

            custom_header("過去の活動履歴")
            try:
                df_activities['person_id'] = pd.to_numeric(df_activities['person_id'], errors='coerce')
                my_activities = df_activities[df_activities['person_id'] == int(current_person_id)].copy()
                if not my_activities.empty:
                    my_activities = my_activities.sort_values('記録日', ascending=False)
                    df_disp_act = my_activities[['記録日', '手段', '要点', '次回予定日']]
                    st.dataframe(df_disp_act, use_container_width=True, hide_index=True)
                else:
                    st.write("まだ記録がありません。")
            except:
                st.write("まだ記録がありません（または読込エラー）。")

    # --- 画面2: 基本情報登録（新規・編集） ---
    elif menu == "基本情報登録":
        custom_header("基本情報登録")
        
        if 'edit_person_id' not in st.session_state:
            st.session_state.edit_person_id = None
        
        st.markdown("### 全利用者一覧")
        
        reg_list_cols = ['ケース番号', '氏名', '生年月日', '年齢', '現在の状態']
        available_reg_cols = [c for c in reg_list_cols if c in df_persons.columns]
        
        if not df_persons.empty and len(available_reg_cols) > 0:
            df_display_reg = df_persons[available_reg_cols]
        else:
            df_display_reg = pd.DataFrame(columns=reg_list_cols)
        
        selection_reg = st.dataframe(
            df_display_reg,
            column_config={
                "年齢": st.column_config.NumberColumn("年齢", format="%d歳"),
            },
            use_container_width=False,
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
            st.markdown(f"### ✏️ 編集モード: {selected_data.get('氏名', '')} さんの情報を修正中")
            if st.button("選択を解除（新規登録へ戻る）"):
                st.session_state.edit_person_id = None
                st.rerun()
        else:
            st.markdown("### ✨ 新規登録モード")
            st.session_state.edit_person_id = None

        with st.form("person_info_form"):
            col1, col2 = st.columns(2)
            val_case_no = selected_data.get('ケース番号', '')
            val_basic_no = selected_data.get('基本事件番号', '')
            val_name = selected_data.get('氏名', '')
            val_kana = selected_data.get('ｼﾒｲ', '')
            
            type_options = ["後見", "保佐", "補助", "任意", "未成年後見", "その他"]
            val_type_raw = selected_data.get('類型', '後見')
            val_type_index = 0
            if val_type_raw in type_options:
                val_type_index = type_options.index(val_type_raw)
            
            val_disability = selected_data.get('障害類型', '')
            val_petitioner = selected_data.get('申立人', '')
            val_court = selected_data.get('管轄家裁', '')
            val_report_month = selected_data.get('家裁報告月', '')
            
            status_options = ["受任中", "終了"]
            val_status_raw = selected_data.get('現在の状態', '受任中')
            val_status_index = 0
            if val_status_raw in status_options:
                val_status_index = status_options.index(val_status_raw)
            
            val_dob = None
            if selected_data.get('生年月日'):
                try: val_dob = pd.to_datetime(selected_data.get('生年月日')).date()
                except: pass
            
            val_ref_date = None
            if selected_data.get('審判確定日'):
                try: val_ref_date = pd.to_datetime(selected_data.get('審判確定日')).date()
                except: pass

            in_case_no = col1.text_input("ケース番号", value=val_case_no)
            in_basic_no = col2.text_input("基本事件番号", value=val_basic_no)
            
            in_name = col1.text_input("氏名 (必須)", value=val_name)
            in_kana = col2.text_input("ｼﾒｲ (カナ)", value=val_kana)
            
            in_dob = col1.date_input("生年月日", value=val_dob if val_dob else None)
            in_type = col2.selectbox("類型", type_options, index=val_type_index)
            
            in_disability = col1.text_input("障害類型", value=val_disability)
            in_petitioner = col2.text_input("申立人", value=val_petitioner)
            
            in_ref_date = col1.date_input("審判確定日", value=val_ref_date if val_ref_date else None)
            in_court = col2.text_input("管轄家裁", value=val_court)
            
            in_report_month = col1.text_input("家裁報告月", value=val_report_month)
            in_status = col2.selectbox("現在の状態", status_options, index=val_status_index)

            submit_btn_text = "情報を更新する" if is_edit_mode else "新規登録する"
            submitted = st.form_submit_button(submit_btn_text)
            
            if submitted:
                if not in_name:
                    st.error("氏名は必須です。")
                else:
                    update_data = {
                        'ケース番号': in_case_no,
                        '基本事件番号': in_basic_no,
                        '氏名': in_name,
                        'ｼﾒｲ': in_kana,
                        '生年月日': str(in_dob) if in_dob else "",
                        '類型': in_type,
                        '障害類型': in_disability,
                        '申立人': in_petitioner,
                        '審判確定日': str(in_ref_date) if in_ref_date else "",
                        '管轄家裁': in_court,
                        '家裁報告月': in_report_month,
                        '現在の状態': in_status
                    }

                    if is_edit_mode:
                        target_id = st.session_state.edit_person_id
                        if update_person_data(target_id, update_data):
                            st.session_state.edit_person_id = None
                            st.rerun()
                    else:
                        new_pid = 1
                        if len(df_persons) > 0:
                            try: new_pid = pd.to_numeric(df_persons['person_id']).max() + 1
                            except: pass
                        
                        new_row = [
                            int(new_pid),
                            in_case_no, in_basic_no, in_name, in_kana,
                            str(in_dob) if in_dob else "",
                            in_type, in_disability, in_petitioner,
                            str(in_ref_date) if in_ref_date else "",
                            in_court, in_report_month, in_status
                        ]
                        add_data_to_sheet("Persons", new_row)
                        st.success(f"{in_name} さんを新規登録しました。")
                        st.rerun()

    # --- 画面3: データ移行 ---
    elif menu == "データ管理・移行":
        custom_header("データ一括インポート")
        st.markdown("指定のCSV様式に貼り付けてアップロードしてください。")

        tab1, tab2 = st.tabs(["1. 利用者データ (Persons)", "2. 活動記録データ (Activities)"])

        with tab1:
            st.subheader("利用者データの移行")
            df_template_p = pd.DataFrame(columns=COL_DEF_PERSONS)
            csv_template_p = df_template_p.to_csv(index=False).encode('cp932')
            st.download_button("📥 様式DL (Persons_Template.csv)", csv_template_p, "Persons_Template.csv", "text/csv")
            
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
                        count = import_csv_to_sheet("Persons", df_upload_p, COL_DEF_PERSONS, date_columns)
                        st.success(f"{count} 件取り込み完了")
                except Exception as e: st.error(f"エラー: {e}")

        with tab2:
            st.subheader("活動記録データの移行")
            df_template_a = pd.DataFrame(columns=COL_DEF_ACTIVITIES)
            csv_template_a = df_template_a.to_csv(index=False).encode('cp932')
            st.download_button("📥 様式DL (Activities_Template.csv)", csv_template_a, "Activities_Template.csv", "text/csv")
            
            uploaded_file_a = st.file_uploader("CSVアップロード", type=["csv"], key="upload_a")
            if uploaded_file_a:
                try:
                    try: df_upload_a = pd.read_csv(uploaded_file_a)
                    except: 
                        uploaded_file_a.seek(0)
                        df_upload_a = pd.read_csv(uploaded_file_a, encoding='cp932')
                    
                    st.write(df_upload_a.head())
                    if st.button("取り込み (Activities)", key="btn_imp_a"):
                        date_columns = ['記録日', '次回予定日']
                        count = import_csv_to_sheet("Activities", df_upload_a, COL_DEF_ACTIVITIES, date_columns)
                        st.success(f"{count} 件取り込み完了")
                except Exception as e: st.error(f"エラー: {e}")

if __name__ == "__main__":
    main()