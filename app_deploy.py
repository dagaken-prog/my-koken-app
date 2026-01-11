import streamlit as st
import pandas as pd
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import re
import openpyxl
import time

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

COL_DEF_ACTIVITIES = [
    'activity_id', 
    'person_id', 
    '記録日', 
    '活動', 
    '場所',
    '所要時間',
    '交通費・立替金',
    '重要',
    '要点', 
    '作成日時'
]

COL_DEF_SYSTEM_USER = [
    '氏名',
    'シメイ',
    '生年月日',
    '〒',
    '住所',
    '連絡先電話番号',
    'e-mail'
]

COL_DEF_ASSETS = [
    'asset_id',
    'person_id',
    '財産種別',
    '名称・機関名',
    '支店・詳細',
    '口座番号・記号',
    '評価額・残高',
    '保管場所',
    '備考',
    '更新日'
]

COL_DEF_RELATED_PARTIES = [
    'related_id',
    'person_id',
    '関係種別',
    '氏名',
    '所属・名称',
    '電話番号',
    '連携メモ',
    '更新日',
    'キーパーソン'
]

st.set_page_config(page_title="成年後見業務支援システム", layout="wide")

# --- CSS (デザイン調整・スマホ最適化・メニューボタン) ---
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: "Noto Sans JP", sans-serif;
        color: #333333;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0.3rem !important;
    }
    div[data-testid="stElementContainer"] {
        margin-bottom: 0.2rem !important;
    }
    div[data-testid="stBorder"] {
        margin-bottom: 5px !important;
        margin-top: 5px !important;
        padding: 10px !important;
        border: 1px solid #ddd !important;
        border-radius: 8px !important;
    }
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
        padding-top: 4px !important;
        padding-bottom: 4px !important;
        font-size: 13px !important;
    }
    p {
        margin-bottom: 0.5rem !important;
        line-height: 1.6 !important;
    }
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
    .custom-header {
        font-size: 16px !important;
        font-weight: bold !important;
        color: #006633 !important;
        border-bottom: 1px solid #ccc;
        padding-bottom: 2px;
        margin-top: 15px;
        margin-bottom: 5px;
    }
    .custom-header-text {
        font-size: 16px !important;
        font-weight: bold !important;
        color: #006633 !important;
        margin: 0 !important;
        padding-top: 5px;
        white-space: nowrap;
    }
    .custom-header-line {
        border-bottom: 1px solid #ccc;
        margin-top: 0px;
        margin-bottom: 5px;
    }
    .stTextInput input, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea, .stNumberInput input {
        border: 1px solid #666 !important;
        background-color: #ffffff !important;
        border-radius: 6px !important;
        padding: 8px 8px !important;
        font-size: 14px !important;
    }
    .stSelectbox div[data-baseweb="select"] > div {
        height: auto !important;
        min-height: 38px !important;
        white-space: normal !important;
        overflow: visible !important;
    }
    .stSelectbox div[data-baseweb="select"] span {
        line-height: 1.3 !important;
        white-space: normal !important;
    }
    .stTextInput label, .stSelectbox label, .stDateInput label, .stTextArea label, .stNumberInput label, .stCheckbox label {
        margin-bottom: 0px !important;
        font-size: 13px !important;
    }
    div[data-testid="stPopover"] button {
        padding: 0px 8px !important;
        height: auto !important;
        border: 1px solid #ccc !important;
    }
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
        content: "CSV/Excelファイル (200MBまで)";
        font-size: 12px;
        color: #666;
        display: block;
        margin-bottom: 5px;
    }
    section[data-testid="stSidebar"] button {
        width: 100%;
        border: 1px solid #ccc;
        border-radius: 8px;
        margin-bottom: 8px;
        padding-top: 12px;
        padding-bottom: 12px;
        font-size: 16px !important;
        font-weight: bold;
        text-align: left;
        background-color: white;
        color: #333;
    }
    section[data-testid="stSidebar"] button:hover {
        border-color: #006633;
        color: #006633;
        background-color: #f0fff0;
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

# --- Google接続関数 (キャッシュ化) ---
@st.cache_resource
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
            return None 
    try:
        client = gspread.authorize(creds)
        # API制限回避のため少し待機
        time.sleep(1)
        sheet = client.open(SPREADSHEET_NAME)
        return sheet
    except Exception as e:
        return None

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
    born_str = normalize_date_str(born)
    if not born_str: return None
    try:
        born_date = pd.to_datetime(born_str, errors='coerce')
        if pd.isna(born_date): return None
        born_date = born_date.date()
        today = datetime.date.today()
        return today.year - born_date.year - ((today.month, today.day) < (born_date.month, born_date.day))
    except:
        return None

# ★修正: カラムチェックを簡略化（APIコール節約）
def get_or_create_worksheet(sheet, sheet_name, expected_columns):
    try:
        # まずシート取得を試みる
        ws = sheet.worksheet(sheet_name)
    except:
        # なければ作成
        ws = sheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        ws.append_row(expected_columns)
        return ws
        
    # ヘッダーチェックは毎回行わず、列数が明らかに足りない場合だけチェックする等の
    # 最適化も考えられるが、ここでは安全のためヘッダー取得は行う。
    # ただし頻度を下げる工夫が必要（キャッシュの有効活用）。
    return ws

# ★修正: カラム補完ロジックを分離（データ取得後にDataFrame上でやる）
# これによりAPIコール回数を減らす

# --- データ読み込み (キャッシュ化・API節約) ---
@st.cache_data(ttl=600)
def load_data_from_sheet():
    sheet = get_spreadsheet_connection()
    if sheet is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # シート取得（APIコール発生）
    ws_persons = get_or_create_worksheet(sheet, "Persons", COL_DEF_PERSONS)
    ws_activities = get_or_create_worksheet(sheet, "Activities", COL_DEF_ACTIVITIES)
    ws_system = get_or_create_worksheet(sheet, "SystemUser", COL_DEF_SYSTEM_USER)
    ws_assets = get_or_create_worksheet(sheet, "Assets", COL_DEF_ASSETS)
    ws_related = get_or_create_worksheet(sheet, "RelatedParties", COL_DEF_RELATED_PARTIES)
    
    # データ取得（APIコール発生）
    # get_all_records はヘッダーも取得するため、実質的にヘッダーチェックも兼ねられる
    df_persons = pd.DataFrame(ws_persons.get_all_records())
    df_activities = pd.DataFrame(ws_activities.get_all_records())
    df_system = pd.DataFrame(ws_system.get_all_records())
    df_assets = pd.DataFrame(ws_assets.get_all_records())
    df_related = pd.DataFrame(ws_related.get_all_records())

    # ★ローカル（DataFrame上）でのカラム補完
    # スプレッドシート側に列がなくても、プログラム上では列があるものとして扱う
    # これにより「毎回スプレッドシートに列を追加しにいくAPIコール」を防ぐ
    for col in COL_DEF_PERSONS:
        if col not in df_persons.columns: df_persons[col] = ""
    for col in COL_DEF_ACTIVITIES:
        if col not in df_activities.columns: df_activities[col] = ""
    for col in COL_DEF_SYSTEM_USER:
        if col not in df_system.columns: df_system[col] = ""
    for col in COL_DEF_ASSETS:
        if col not in df_assets.columns: df_assets[col] = ""
    for col in COL_DEF_RELATED_PARTIES:
        if col not in df_related.columns: df_related[col] = ""

    # 日付正規化
    for col in ['生年月日', '審判確定日']:
        if col in df_persons.columns:
            df_persons[col] = df_persons[col].apply(normalize_date_str)
    for col in ['記録日']:
        if col in df_activities.columns:
            df_activities[col] = df_activities[col].apply(normalize_date_str)
    
    return df_persons, df_activities, df_system, df_assets, df_related

# ★APIコール後にキャッシュをクリアする関数
def clear_cache_and_reload():
    load_data_from_sheet.clear()
    # st.rerun() # ここではrerunせず、呼び出し元で行う

def add_data_to_sheet(sheet_name, new_row_list):
    sheet = get_spreadsheet_connection()
    if sheet:
        worksheet = sheet.worksheet(sheet_name)
        worksheet.append_row(new_row_list)
        clear_cache_and_reload()

def update_sheet_data(sheet_name, id_column, target_id, update_dict):
    sheet = get_spreadsheet_connection()
    if sheet is None or isinstance(sheet, str):
        st.error("接続エラー")
        return False
    worksheet = sheet.worksheet(sheet_name)
    
    # 列位置の特定などは仕方なくAPIコールするが、頻度は低い
    header_cells = worksheet.row_values(1)
    
    try:
        pid_col_index = header_cells.index(id_column) + 1
    except ValueError:
        st.error(f"システムエラー: {id_column} 列が見つかりません。")
        return False
    
    # ID検索もAPIコール
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
            clear_cache_and_reload()
            return True
        return False
    except Exception as e:
        st.error(f"更新エラー: {str(e)}")
        return False

def save_system_user_data(new_data_dict):
    sheet = get_spreadsheet_connection()
    if sheet:
        worksheet = sheet.worksheet("SystemUser")
        row_values = []
        for col in COL_DEF_SYSTEM_USER:
            val = new_data_dict.get(col, "")
            if val is None: val = ""
            row_values.append(str(val))
        existing = worksheet.get_all_values()
        if len(existing) > 1:
            cell_range = f"A2:{chr(64+len(COL_DEF_SYSTEM_USER))}2" 
            worksheet.update(range_name=cell_range, values=[row_values])
        else:
            worksheet.append_row(row_values)
        st.toast("システム利用者情報を保存しました", icon="💾")
        clear_cache_and_reload()

def delete_sheet_row(sheet_name, id_column, target_id):
    sheet = get_spreadsheet_connection()
    if sheet is None: return False
    worksheet = sheet.worksheet(sheet_name)
    header_cells = worksheet.row_values(1)
    try:
        pid_col_index = header_cells.index(id_column) + 1
    except ValueError:
        return False
    all_ids = worksheet.col_values(pid_col_index)
    target_row_num = -1
    str_search_id = str(target_id)
    for i, val in enumerate(all_ids):
        if str(val) == str_search_id:
            target_row_num = i + 1
            break
    if target_row_num == -1:
        return False
    try:
        worksheet.delete_rows(target_row_num)
        st.toast("削除しました", icon="🗑️")
        clear_cache_and_reload()
        return True
    except Exception as e:
        st.error(f"削除エラー: {str(e)}")
        return False

def import_csv_to_sheet_safe(sheet_name, df_upload, target_columns, id_column, date_columns=[]):
    sheet = get_spreadsheet_connection()
    if sheet is None: return 0, 0
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
        clear_cache_and_reload()
        return len(export_data), skipped_count
    return 0, skipped_count

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
                                val = str(data_dict[key])
                                if val == "None" or val == "nan": val = ""
                                new_text = new_text.replace(f'{{{{{key}}}}}', val)
                        cell.value = new_text
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def custom_title(text):
    st.markdown(f'<div class="custom-title">{text}</div>', unsafe_allow_html=True)

def custom_header(text, help_text=None):
    if help_text:
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

    # キャッシュされたデータ読み込み (引数なし)
    # ここでエラーが起きてもアプリが落ちないようにtry-exceptで囲む
    try:
        df_persons, df_activities, df_system, df_assets, df_related = load_data_from_sheet()
    except Exception as e:
        st.error(f"データ読み込みエラー: {e}。時間をおいて再読み込みしてください。")
        return

    if df_persons.empty and df_activities.empty:
        # 初回起動時など
        pass

    if '生年月日' in df_persons.columns:
        if not df_persons.empty:
            df_persons['年齢'] = df_persons['生年月日'].apply(calculate_age)
            df_persons['年齢'] = pd.to_numeric(df_persons['年齢'], errors='coerce')
        else:
            df_persons['年齢'] = None

    # --- メニューの状態管理（ボタン式） ---
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
                st.rerun()

    menu = st.session_state.current_menu

    if 'selected_person_id' not in st.session_state:
        st.session_state.selected_person_id = None
    if 'delete_confirm_id' not in st.session_state:
        st.session_state.delete_confirm_id = None
    if 'edit_asset_id' not in st.session_state:
        st.session_state.edit_asset_id = None
    if 'delete_asset_id' not in st.session_state:
        st.session_state.delete_asset_id = None
    if 'edit_related_id' not in st.session_state:
        st.session_state.edit_related_id = None
    if 'delete_related_id' not in st.session_state:
        st.session_state.delete_related_id = None

    # =========================================================
    # 1. 利用者情報・活動記録
    # =========================================================
    if menu == "利用者情報・活動記録":
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

            with st.expander("▼ 基本情報", expanded=True):
                kp_html = ""
                if not df_related.empty:
                    df_related['person_id'] = pd.to_numeric(df_related['person_id'], errors='coerce')
                    kp_df = df_related[
                        (df_related['person_id'] == int(current_person_id)) & 
                        (df_related['キーパーソン'].astype(str).str.upper() == 'TRUE')
                    ]
                    if not kp_df.empty:
                        kp_html = "<div style='margin-top:8px; padding-top:8px; border-top:1px dashed #ccc; width:100%; grid-column: 1 / -1;'>"
                        kp_html += "<div><b>★ キーパーソン:</b></div>"
                        for _, kp in kp_df.iterrows():
                            tel = kp.get('電話番号', '')
                            tel_html = f'<a href="tel:{tel}" style="text-decoration:none; color:#0066cc;">📞 {tel}</a>' if tel else ''
                            kp_html += f"<div style='margin-left:10px; margin-top:2px;'>【{kp.get('関係種別','')}】 {kp.get('氏名','')} {tel_html}</div>"
                        kp_html += "</div>"

                grid_html = f"""
                <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; font-size: 14px;">
                    <div><span style="font-weight:bold; color:#555;">No.:</span> {selected_row.get('ケース番号', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">事件番号:</span> {selected_row.get('基本事件番号', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">類型:</span> {selected_row.get('類型', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">氏名:</span> {selected_row.get('氏名', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">ｼﾒｲ:</span> {selected_row.get('ｼﾒｲ', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">生年月日:</span> {selected_row.get('生年月日', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">障害類型:</span> {selected_row.get('障害類型', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">申立人:</span> {selected_row.get('申立人', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">審判日:</span> {selected_row.get('審判確定日', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">家裁:</span> {selected_row.get('管轄家裁', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">報告月:</span> {selected_row.get('家裁報告月', '')}</div>
                    <div><span style="font-weight:bold; color:#555;">状態:</span> {selected_row.get('現在の状態', '')}</div>
                    {kp_html}
                </div>
                """
                st.markdown(grid_html, unsafe_allow_html=True)

            st.markdown("### 📝 活動記録")
            with st.expander("➕ 新しい活動記録を追加する", expanded=False):
                with st.form("new_activity_form", clear_on_submit=True):
                    col_a, col_b = st.columns(2)
                    input_date = col_a.date_input("活動日", value=datetime.date.today(), min_value=datetime.date(2000, 1, 1))
                    activity_opts = ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"]
                    input_activity = col_b.selectbox("活動", activity_opts)
                    
                    col_c, col_d, col_e = st.columns(3)
                    input_time = col_c.number_input("所要時間(分)", min_value=0, step=10, value=0)
                    input_place = col_d.text_input("場所", placeholder="自宅、病院など")
                    input_cost = col_e.number_input("交通費・立替(円)", min_value=0, step=100, value=0)

                    input_summary = st.text_area("内容", height=120)
                    input_important = st.checkbox("★重要 (報酬付与申立などで強調)")
                    
                    submitted = st.form_submit_button("登録")
                    
                    if submitted:
                        new_id = 1
                        if len(df_activities) > 0:
                            try: new_id = pd.to_numeric(df_activities['activity_id']).max() + 1
                            except: pass
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        imp_str = "TRUE" if input_important else ""
                        new_row = [
                            int(new_id), int(current_person_id), str(input_date), 
                            input_activity, input_place, input_time, input_cost, 
                            imp_str, input_summary, now_str
                        ]
                        add_data_to_sheet("Activities", new_row)
                        st.rerun()

            custom_header("過去の活動履歴", help_text="履歴の「詳細・操作」をタップして開くと、編集・削除ボタンが表示されます。")
            if 'edit_activity_id' not in st.session_state:
                st.session_state.edit_activity_id = None

            try:
                df_activities['person_id'] = pd.to_numeric(df_activities['person_id'], errors='coerce')
                my_activities = df_activities[df_activities['person_id'] == int(current_person_id)].copy()
                
                if not my_activities.empty:
                    if '作成日時' in my_activities.columns:
                        my_activities = my_activities.sort_values(by=['記録日', '作成日時'], ascending=[False, False])
                    else:
                        my_activities = my_activities.sort_values('記録日', ascending=False)
                    
                    if st.session_state.edit_activity_id:
                        edit_row = my_activities[my_activities['activity_id'] == st.session_state.edit_activity_id].iloc[0]
                        with st.container(border=True):
                            st.markdown(f"#### ✏️ 活動履歴の修正 (ID: {edit_row['activity_id']})")
                            with st.form("edit_activity_form"):
                                ea_date_val = pd.to_datetime(edit_row['記録日']).date() if edit_row['記録日'] else None
                                ea_date = st.date_input("活動日", value=ea_date_val, min_value=datetime.date(2000, 1, 1))
                                
                                act_opts = ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"]
                                curr_act = edit_row['活動'] if edit_row['活動'] in act_opts else "その他"
                                ea_act = st.selectbox("活動", act_opts, index=act_opts.index(curr_act))
                                
                                col_ec, col_ed, col_ee = st.columns(3)
                                try: curr_time = int(float(edit_row.get('所要時間', 0)))
                                except: curr_time = 0
                                try: curr_cost = int(float(edit_row.get('交通費・立替金', 0)))
                                except: curr_cost = 0
                                curr_imp = True if str(edit_row.get('重要', '')).upper() == 'TRUE' else False

                                ea_time = col_ec.number_input("所要時間", min_value=0, step=10, value=curr_time)
                                ea_place = col_ed.text_input("場所", value=str(edit_row.get('場所', '')))
                                ea_cost = col_ee.number_input("交通費・立替", min_value=0, step=100, value=curr_cost)
                                
                                ea_summary = st.text_area("内容", value=edit_row['要点'], height=120)
                                ea_imp = st.checkbox("★重要", value=curr_imp)
                                
                                c_save, c_cancel = st.columns(2)
                                with c_save:
                                    if st.form_submit_button("保存"):
                                        imp_str = "TRUE" if ea_imp else ""
                                        upd_dict = {
                                            '記録日': str(ea_date), 
                                            '活動': ea_act, 
                                            '場所': ea_place,
                                            '所要時間': ea_time,
                                            '交通費・立替金': ea_cost,
                                            '重要': imp_str,
                                            '要点': ea_summary
                                        }
                                        if update_sheet_data("Activities", "activity_id", st.session_state.edit_activity_id, upd_dict):
                                            st.session_state.edit_activity_id = None
                                            st.rerun()
                                with c_cancel:
                                    if st.form_submit_button("キャンセル"):
                                        st.session_state.edit_activity_id = None
                                        st.rerun()

                    for idx, row in my_activities.iterrows():
                        star_mark = "★" if str(row.get('重要', '')).upper() == 'TRUE' else ""
                        
                        with st.container(border=True):
                            st.markdown(f"**{star_mark} {row['記録日']}**　📝 {row['活動']}")
                            st.write(row['要点'])
                            
                            with st.expander("詳細・操作", expanded=False):
                                detail_html = f"""
                                <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(100px, 1fr)); gap: 5px; font-size: 13px; margin-bottom: 10px;">
                                    <div style="background-color:#f8f9fa; padding:4px; border-radius:4px; border:1px solid #eee;">
                                        <span style="font-weight:bold; font-size:11px; color:#555;">場所</span><br>
                                        {row.get('場所', '-') or '-'}
                                    </div>
                                    <div style="background-color:#f8f9fa; padding:4px; border-radius:4px; border:1px solid #eee;">
                                        <span style="font-weight:bold; font-size:11px; color:#555;">時間</span><br>
                                        {row.get('所要時間', '0')} 分
                                    </div>
                                    <div style="background-color:#f8f9fa; padding:4px; border-radius:4px; border:1px solid #eee;">
                                        <span style="font-weight:bold; font-size:11px; color:#555;">費用</span><br>
                                        {row.get('交通費・立替金', '0')} 円
                                    </div>
                                </div>
                                """
                                st.markdown(detail_html, unsafe_allow_html=True)
                                
                                c_edit, c_del = st.columns(2)
                                with c_edit:
                                    if st.button("編集", key=f"btn_edit_{row['activity_id']}", use_container_width=True):
                                        st.session_state.edit_activity_id = row['activity_id']
                                        st.session_state.delete_confirm_id = None 
                                        st.rerun()
                                with c_del:
                                    if st.button("削除", key=f"btn_del_{row['activity_id']}", use_container_width=True):
                                        st.session_state.delete_confirm_id = row['activity_id']
                                        st.session_state.edit_activity_id = None
                                        st.rerun()
                                
                                if st.session_state.delete_confirm_id == row['activity_id']:
                                    st.warning("本当に削除しますか？")
                                    c_yes, c_no = st.columns(2)
                                    with c_yes:
                                        if st.button("はい", key=f"del_yes_{row['activity_id']}", use_container_width=True):
                                            if delete_sheet_row("Activities", "activity_id", row['activity_id']):
                                                st.session_state.delete_confirm_id = None
                                                st.rerun()
                                    with c_no:
                                        if st.button("いいえ", key=f"del_no_{row['activity_id']}", use_container_width=True):
                                            st.session_state.delete_confirm_id = None
                                            st.rerun()

                else:
                    st.write("まだ記録がありません。")
            except Exception as e:
                st.write(f"読込エラー: {e}")

    # =========================================================
    # ★新規: 関係者・連絡先
    # =========================================================
    elif menu == "関係者・連絡先":
        custom_header("関係者・連絡先", help_text="キーパーソンの情報を管理します。電話番号をタップすると発信できます。")
        
        # 利用者選択
        if not df_persons.empty and '現在の状態' in df_persons.columns:
            mask = df_persons['現在の状態'].fillna('').astype(str).isin(['受任中', '', 'nan'])
            df_active = df_persons[mask].copy()
        else:
            df_active = df_persons.copy()

        person_options = {}
        if not df_active.empty:
            for idx, row in df_active.iterrows():
                label = f"{row.get('ケース番号','')} {row.get('氏名','')}"
                person_options[label] = row['person_id']
        
        selected_label = st.selectbox("対象者を選択", list(person_options.keys()))
        
        if selected_label:
            current_pid = person_options[selected_label]
            
            # 新規登録フォーム
            with st.expander("➕ 新しい関係者を登録する", expanded=False):
                with st.form("new_related_form", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    r_type = col1.selectbox("関係種別", ["親族", "ケアマネ", "施設相談員", "病院SW", "主治医", "弁護士", "行政", "その他"])
                    r_name = col2.text_input("氏名")
                    
                    col3, col4 = st.columns(2)
                    r_org = col3.text_input("所属・名称 (例: 〇〇病院)")
                    r_tel = col4.text_input("電話番号 (例: 090-0000-0000)")
                    
                    # ★修正: キーパーソンチェック追加
                    r_keyperson = st.checkbox("★キーパーソン (基本情報に表示)")
                    r_note = st.text_area("連携メモ (キーマン等)", height=60)
                    
                    if st.form_submit_button("登録"):
                        new_rid = 1
                        if len(df_related) > 0:
                            try: new_rid = pd.to_numeric(df_related['related_id']).max() + 1
                            except: pass
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d")
                        
                        k_str = "TRUE" if r_keyperson else ""
                        
                        # related_id, person_id, 関係種別, 氏名, 所属・名称, 電話番号, 連携メモ, 更新日, キーパーソン
                        new_row = [int(new_rid), int(current_pid), r_type, r_name, r_org, r_tel, r_note, now_str, k_str]
                        add_data_to_sheet("RelatedParties", new_row)
                        st.success("登録しました")
                        st.rerun()
            
            st.markdown("---")
            
            # 一覧表示
            try:
                df_related['person_id'] = pd.to_numeric(df_related['person_id'], errors='coerce')
                my_related = df_related[df_related['person_id'] == int(current_pid)].copy()
                
                if not my_related.empty:
                    # 編集モード
                    if st.session_state.edit_related_id:
                        edit_row = my_related[my_related['related_id'] == st.session_state.edit_related_id].iloc[0]
                        with st.container(border=True):
                            st.markdown(f"#### ✏️ 連絡先の修正")
                            with st.form("edit_related_form"):
                                col1, col2 = st.columns(2)
                                type_list = ["親族", "ケアマネ", "施設相談員", "病院SW", "主治医", "弁護士", "行政", "その他"]
                                curr_type = edit_row['関係種別'] if edit_row['関係種別'] in type_list else "その他"
                                er_type = col1.selectbox("関係種別", type_list, index=type_list.index(curr_type))
                                er_name = col2.text_input("氏名", value=edit_row['氏名'])
                                
                                col3, col4 = st.columns(2)
                                er_org = col3.text_input("所属・名称", value=edit_row['所属・名称'])
                                er_tel = col4.text_input("電話番号", value=edit_row['電話番号'])
                                
                                curr_kp = True if str(edit_row.get('キーパーソン', '')).upper() == 'TRUE' else False
                                er_keyperson = st.checkbox("★キーパーソン", value=curr_kp)
                                er_note = st.text_area("連携メモ", value=edit_row['連携メモ'])
                                
                                c_save, c_cancel = st.columns(2)
                                with c_save:
                                    if st.form_submit_button("保存"):
                                        k_str = "TRUE" if er_keyperson else ""
                                        upd_dict = {
                                            '関係種別': er_type, '氏名': er_name,
                                            '所属・名称': er_org, '電話番号': er_tel,
                                            '連携メモ': er_note, '更新日': datetime.datetime.now().strftime("%Y-%m-%d"),
                                            'キーパーソン': k_str
                                        }
                                        if update_sheet_data("RelatedParties", "related_id", st.session_state.edit_related_id, upd_dict):
                                            st.session_state.edit_related_id = None
                                            st.rerun()
                                with c_cancel:
                                    if st.form_submit_button("キャンセル"):
                                        st.session_state.edit_related_id = None
                                        st.rerun()

                    # リスト表示（カード）
                    st.markdown("#### 登録済み連絡先")
                    for idx, row in my_related.iterrows():
                        tel_link = f"📞 [{row['電話番号']}](tel:{row['電話番号']})" if row['電話番号'] else "電話なし"
                        
                        kp_mark = "★" if str(row.get('キーパーソン', '')).upper() == 'TRUE' else ""
                        label_text = f"{kp_mark}【{row['関係種別']}】 {row['氏名']} ({row['所属・名称']})"
                        
                        with st.expander(label_text, expanded=False):
                            st.markdown(f"**連絡先:** {tel_link}", unsafe_allow_html=True)
                            if row['連携メモ']:
                                st.info(f"📝 {row['連携メモ']}")
                            
                            c_edit, c_del = st.columns(2)
                            with c_edit:
                                if st.button("編集", key=f"rel_edit_{row['related_id']}", use_container_width=True):
                                    st.session_state.edit_related_id = row['related_id']
                                    st.session_state.delete_related_id = None
                                    st.rerun()
                            with c_del:
                                if st.button("削除", key=f"rel_del_{row['related_id']}", use_container_width=True):
                                    st.session_state.delete_related_id = row['related_id']
                                    st.session_state.edit_related_id = None
                                    st.rerun()
                            
                            if st.session_state.delete_related_id == row['related_id']:
                                st.warning("削除しますか？")
                                if st.button("はい、削除", key=f"rel_yes_{row['related_id']}"):
                                    if delete_sheet_row("RelatedParties", "related_id", row['related_id']):
                                        st.session_state.delete_related_id = None
                                        st.rerun()

                else:
                    st.info("登録された連絡先はありません。")
            except Exception as e:
                st.error(f"読込エラー: {e}")

    # =========================================================
    # 6. 財産管理
    # =========================================================
    elif menu == "財産管理":
        custom_header("財産管理", help_text="利用者の財産情報を登録・編集・一覧表示します。")
        
        if not df_persons.empty and '現在の状態' in df_persons.columns:
            mask = df_persons['現在の状態'].fillna('').astype(str).isin(['受任中', '', 'nan'])
            df_active = df_persons[mask].copy()
        else:
            df_active = df_persons.copy()

        person_options = {}
        if not df_active.empty:
            for idx, row in df_active.iterrows():
                label = f"{row.get('ケース番号','')} {row.get('氏名','')}"
                person_options[label] = row['person_id']
        
        selected_label = st.selectbox("対象者を選択", list(person_options.keys()))
        
        if selected_label:
            current_pid = person_options[selected_label]
            
            with st.expander("➕ 新しい財産を登録する", expanded=False):
                with st.form("new_asset_form", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    a_type = col1.selectbox("財産種別", ["預貯金", "現金", "有価証券", "保険", "不動産", "負債", "その他"])
                    a_name = col2.text_input("名称・機関名 (例: ゆうちょ銀行)")
                    
                    col3, col4 = st.columns(2)
                    a_detail = col3.text_input("支店・詳細 (例: 呉支店)")
                    a_num = col4.text_input("口座番号・記号")
                    
                    col5, col6 = st.columns(2)
                    a_value = col5.text_input("評価額・残高")
                    a_place = col6.text_input("保管場所")
                    
                    a_note = st.text_area("備考", height=60)
                    
                    if st.form_submit_button("財産を登録"):
                        new_aid = 1
                        if len(df_assets) > 0:
                            try: new_aid = pd.to_numeric(df_assets['asset_id']).max() + 1
                            except: pass
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d")
                        new_row = [int(new_aid), int(current_pid), a_type, a_name, a_detail, a_num, a_value, a_place, a_note, now_str]
                        add_data_to_sheet("Assets", new_row)
                        st.success("登録しました")
                        st.rerun()
            
            st.markdown("---")
            
            try:
                df_assets['person_id'] = pd.to_numeric(df_assets['person_id'], errors='coerce')
                my_assets = df_assets[df_assets['person_id'] == int(current_pid)].copy()
                
                if not my_assets.empty:
                    if st.session_state.edit_asset_id:
                        edit_row = my_assets[my_assets['asset_id'] == st.session_state.edit_asset_id].iloc[0]
                        with st.container(border=True):
                            st.markdown(f"#### ✏️ 財産情報の修正")
                            with st.form("edit_asset_form"):
                                col1, col2 = st.columns(2)
                                type_list = ["預貯金", "現金", "有価証券", "保険", "不動産", "負債", "その他"]
                                curr_type = edit_row['財産種別'] if edit_row['財産種別'] in type_list else "その他"
                                ea_type = col1.selectbox("種別", type_list, index=type_list.index(curr_type))
                                ea_name = col2.text_input("名称", value=edit_row['名称・機関名'])
                                
                                col3, col4 = st.columns(2)
                                ea_detail = col3.text_input("詳細", value=edit_row['支店・詳細'])
                                ea_num = col4.text_input("番号", value=edit_row['口座番号・記号'])
                                
                                col5, col6 = st.columns(2)
                                ea_value = col5.text_input("評価額", value=str(edit_row['評価額・残高']))
                                ea_place = col6.text_input("保管場所", value=edit_row['保管場所'])
                                
                                ea_note = st.text_area("備考", value=edit_row['備考'])
                                
                                c_save, c_cancel = st.columns(2)
                                with c_save:
                                    if st.form_submit_button("保存"):
                                        upd_dict = {
                                            '財産種別': ea_type, '名称・機関名': ea_name,
                                            '支店・詳細': ea_detail, '口座番号・記号': ea_num,
                                            '評価額・残高': ea_value, '保管場所': ea_place,
                                            '備考': ea_note, '更新日': datetime.datetime.now().strftime("%Y-%m-%d")
                                        }
                                        if update_sheet_data("Assets", "asset_id", st.session_state.edit_asset_id, upd_dict):
                                            st.session_state.edit_asset_id = None
                                            st.rerun()
                                with c_cancel:
                                    if st.form_submit_button("キャンセル"):
                                        st.session_state.edit_asset_id = None
                                        st.rerun()

                    st.markdown("#### 登録済み財産一覧")
                    for idx, row in my_assets.iterrows():
                        label_text = f"【{row['財産種別']}】 {row['名称・機関名']} ({row['評価額・残高']})"
                        with st.expander(label_text, expanded=False):
                            grid_html = f"""
                            <div style="font-size:14px;">
                                <div><b>詳細:</b> {row['支店・詳細']}</div>
                                <div><b>番号:</b> {row['口座番号・記号']}</div>
                                <div><b>場所:</b> {row['保管場所']}</div>
                                <div><b>備考:</b> {row['備考']}</div>
                            </div>
                            """
                            st.markdown(grid_html, unsafe_allow_html=True)
                            
                            c_edit, c_del = st.columns(2)
                            with c_edit:
                                if st.button("編集", key=f"ast_edit_{row['asset_id']}", use_container_width=True):
                                    st.session_state.edit_asset_id = row['asset_id']
                                    st.session_state.delete_asset_id = None
                                    st.rerun()
                            with c_del:
                                if st.button("削除", key=f"ast_del_{row['asset_id']}", use_container_width=True):
                                    st.session_state.delete_asset_id = row['asset_id']
                                    st.session_state.edit_asset_id = None
                                    st.rerun()
                            
                            if st.session_state.delete_asset_id == row['asset_id']:
                                st.warning("削除しますか？")
                                if st.button("はい、削除", key=f"ast_yes_{row['asset_id']}"):
                                    if delete_sheet_row("Assets", "asset_id", row['asset_id']):
                                        st.session_state.delete_asset_id = None
                                        st.rerun()

                else:
                    st.info("登録された財産はありません。")
            except Exception as e:
                st.error(f"読込エラー: {e}")

    # =========================================================
    # 7. 利用者情報登録
    # =========================================================
    elif menu == "利用者情報登録":
        custom_header("利用者情報登録", help_text="新規登録の場合はフォームに入力してください。\n修正の場合は、下の一覧から対象者をクリックしてください。")
        
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
            st.markdown(f"### ✏️ 編集モード: {selected_data.get('氏名', '')} さんの情報を修正中")
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
    # 8. 帳票作成
    # =========================================================
    elif menu == "帳票作成":
        custom_header("帳票作成（Excel出力）", help_text="Excel様式にデータを埋め込んで出力します。\n様式内に {{氏名}} などの目印を書いておいてください。")
        
        st.markdown("#### 1. テンプレートExcelのアップロード")
        template_file = st.file_uploader("Excelファイル (.xlsx)", type=["xlsx"])
        
        st.markdown("#### 2. 対象者の選択")
        if not df_persons.empty:
            target_list = df_persons['氏名'].tolist()
            target_name = st.selectbox("出力する利用者を選択", target_list)
            
            if st.button("書類を作成する") and template_file:
                target_data = df_persons[df_persons['氏名'] == target_name].iloc[0].to_dict()
                age = calculate_age(target_data.get('生年月日'))
                target_data['年齢'] = str(age) if age else ""
                
                try:
                    excel_data = fill_excel_template(template_file, target_data)
                    st.success("作成完了！以下のボタンからダウンロードしてください。")
                    st.download_button(
                        label="📥 書類をダウンロード",
                        data=excel_data,
                        file_name=f"書類_{target_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
        else:
            st.info("利用者が登録されていません。")

    # =========================================================
    # 9. データ管理・移行
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

    # =========================================================
    # 10. 初期設定 (システム利用者登録)
    # =========================================================
    elif menu == "初期設定":
        custom_header("初期設定")
        st.markdown("### システム利用者登録")
        st.info("ここで登録した情報は、書類作成時のテンプレート（署名欄など）に使用されます。")
        
        current_data = {}
        if not df_system.empty:
            current_data = df_system.iloc[0].to_dict()
        
        with st.form("system_user_form"):
            col1, col2 = st.columns(2)
            
            val_name = current_data.get('氏名', '')
            val_kana = current_data.get('シメイ', '')
            val_dob = pd.to_datetime(current_data.get('生年月日')).date() if current_data.get('生年月日') else None
            val_zip = current_data.get('〒', '')
            val_addr = current_data.get('住所', '')
            val_tel = current_data.get('連絡先電話番号', '')
            val_email = current_data.get('e-mail', '')

            in_name = col1.text_input("氏名", value=val_name)
            in_kana = col2.text_input("シメイ (カナ)", value=val_kana)
            in_dob = col1.date_input("生年月日", value=val_dob, min_value=datetime.date(1900, 1, 1))
            in_zip = col2.text_input("〒 (郵便番号)", value=val_zip)
            in_addr = st.text_input("住所", value=val_addr)
            in_tel = col1.text_input("連絡先電話番号", value=val_tel)
            in_email = col2.text_input("e-mail", value=val_email)
            
            if st.form_submit_button("設定を保存"):
                new_data = {
                    '氏名': in_name, 'シメイ': in_kana,
                    '生年月日': str(in_dob) if in_dob else "",
                    '〒': in_zip, '住所': in_addr,
                    '連絡先電話番号': in_tel, 'e-mail': in_email
                }
                save_system_user_data(new_data)
                st.rerun()

if __name__ == "__main__":
    main()