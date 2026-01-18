import streamlit as st
import pandas as pd
import datetime
from supabase import create_client, Client
import io
import openpyxl
import time
import re

st.set_page_config(page_title="成年後見業務支援システム", layout="wide")

# --- Supabase接続設定 ---
try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("【設定エラー】Secretsが見つかりません。.streamlit/secrets.toml を確認してください。")
    st.stop()

# --- データベースとアプリの項目名マッピング ---
MAP_PERSONS = {
    'person_id': 'person_id', 'ケース番号': 'case_number', '基本事件番号': 'basic_case_number',
    '氏名': 'name', 'ｼﾒｲ': 'kana', '生年月日': 'dob', '類型': 'guardianship_type',
    '障害類型': 'disability_type', '申立人': 'petitioner', '審判確定日': 'judgment_date',
    '管轄家裁': 'court', '家裁報告月': 'report_month', '現在の状態': 'status'
}

MAP_ACTIVITIES = {
    'activity_id': 'activity_id', 'person_id': 'person_id', '記録日': 'activity_date',
    '活動': 'activity_type', '場所': 'location', '所要時間': 'duration',
    '交通費・立替金': 'expense', '重要': 'is_important', '要点': 'note', '作成日時': 'created_at'
}

MAP_ASSETS = {
    'asset_id': 'asset_id', 'person_id': 'person_id', '財産種別': 'asset_type',
    '名称・機関名': 'name', '支店・詳細': 'detail', '口座番号・記号': 'account_number',
    '評価額・残高': 'value', '保管場所': 'storage_location', '備考': 'note', '更新日': 'updated_at'
}

MAP_RELATED = {
    'related_id': 'related_id', 'person_id': 'person_id', '関係種別': 'relationship',
    '氏名': 'name', '所属・名称': 'organization', '電話番号': 'phone', '〒': 'postal_code',
    '住所': 'address', 'e-mail': 'email', '連携メモ': 'note', '更新日': 'updated_at',
    'キーパーソン': 'is_keyperson'
}

MAP_SYSTEM = {
    'id': 'id', '氏名': 'name', 'シメイ': 'kana', '生年月日': 'dob',
    '〒': 'postal_code', '住所': 'address', '連絡先電話番号': 'phone', 'e-mail': 'email'
}

# 逆引き用辞書
R_MAP_PERSONS = {v: k for k, v in MAP_PERSONS.items()}
R_MAP_ACTIVITIES = {v: k for k, v in MAP_ACTIVITIES.items()}
R_MAP_ASSETS = {v: k for k, v in MAP_ASSETS.items()}
R_MAP_RELATED = {v: k for k, v in MAP_RELATED.items()}
R_MAP_SYSTEM = {v: k for k, v in MAP_SYSTEM.items()}

# --- CSS設定 (スマホ最適化・ヘッダー非表示) ---
st.markdown("""
    <style>
    html, body, [class*="css"] { font-family: "Noto Sans JP", sans-serif; color: #333; }
    
    /* ★追加: Streamlit標準のヘッダーバーを非表示にする */
    header[data-testid="stHeader"] {
        display: none;
    }
    
    /* 余白設定 */
    .block-container { 
        padding-top: 1rem !important; /* ヘッダーを消したので上部余白を確保 */
        padding-bottom: 3rem !important; 
        padding-left: 1rem !important; 
        padding-right: 1rem !important; 
    }
    div[data-testid="stVerticalBlock"] { gap: 0.5rem !important; }
    div[data-testid="stElementContainer"] { margin-bottom: 0.3rem !important; }
    
    /* カードデザイン */
    div[data-testid="stBorder"] { 
        margin: 5px 0 !important; 
        padding: 10px !important; 
        border: 1px solid #ddd !important; 
        border-radius: 8px; 
        background-color: #fff;
    }

    /* テーブルスタイル */
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { padding: 6px !important; font-size: 14px !important; }
    
    /* テキストスタイル */
    p { margin-bottom: 0.5rem !important; line-height: 1.6 !important; }
    h2 { padding: 10px 0 !important; margin-bottom: 20px !important; line-height: 1.5 !important; }
    
    /* タイトル・ヘッダー */
    .custom-title { font-size: 20px; font-weight: bold; color: #006633; border-left: 6px solid #006633; padding: 5px 0 5px 10px; margin: 5px 0 10px 0; background-color: #f8f9fa; }
    .custom-header { font-size: 16px; font-weight: bold; color: #006633; border-bottom: 1px solid #ccc; padding-bottom: 2px; margin: 20px 0 10px 0; }
    .custom-header-text { font-size: 16px; font-weight: bold; color: #006633; margin: 0; padding-top: 5px; white-space: nowrap; }
    .custom-header-line { border-bottom: 1px solid #ccc; margin: 0 0 5px 0; }
    
    /* フォーム部品 */
    .stTextInput input, .stDateInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea, .stNumberInput input { border: 1px solid #666 !important; background-color: #fff !important; border-radius: 6px !important; padding: 8px 8px !important; font-size: 14px !important; }
    .stSelectbox div[data-baseweb="select"] > div { height: auto !important; min-height: 40px !important; }
    .stTextInput label, .stSelectbox label, .stDateInput label, .stTextArea label, .stNumberInput label, .stCheckbox label { margin-bottom: 2px !important; font-size: 13px !important; font-weight: bold; }
    
    /* ボタン類 */
    div[data-testid="stPopover"] button { padding: 0 8px !important; height: auto !important; border: 1px solid #ccc !important; }
    section[data-testid="stSidebar"] button { width: 100%; border: 1px solid #ccc; border-radius: 8px; margin-bottom: 8px; padding: 12px; font-size: 16px !important; font-weight: bold; text-align: left; background-color: white; color: #333; }
    section[data-testid="stSidebar"] button:hover { border-color: #006633; color: #006633; background-color: #f0fff0; }
    
    /* アップローダー */
    [data-testid="stFileUploaderDropzone"] div div span, [data-testid="stFileUploaderDropzone"] div div small { display: none; }
    [data-testid="stFileUploaderDropzone"] div div::after { content: "ファイルをドラッグ＆ドロップまたは選択"; font-size: 12px; font-weight: bold; color: #333; display: block; margin: 5px 0; }
    [data-testid="stFileUploaderDropzone"] div div::before { content: "CSV/Excelファイル (200MBまで)"; font-size: 12px; color: #666; display: block; margin-bottom: 5px; }
    </style>
""", unsafe_allow_html=True)

# --- 認証機能 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if st.session_state.password_correct:
        return True
    
    with st.container():
        st.markdown("## 🔒 ログイン")
        password = st.text_input("パスワードを入力してください", type="password")
        if st.button("ログイン"):
            correct_password = "admin"
            if "APP_PASSWORD" in st.secrets:
                correct_password = st.secrets["APP_PASSWORD"]
            if password == correct_password:
                st.session_state.password_correct = True
                st.success("ログインしました")
                st.rerun()
            else:
                st.error("パスワードが違います")
    return False

# --- Supabase操作関数 ---
@st.cache_resource
def init_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

# データをキャッシュして高速化
@st.cache_data(ttl=600)
def fetch_table(table_name, mapping_dict):
    client = init_supabase()
    try:
        response = client.table(table_name).select("*").execute()
        data = response.data
    except Exception as e:
        return pd.DataFrame(columns=mapping_dict.keys())
    
    if not data:
        return pd.DataFrame(columns=mapping_dict.keys())
    
    df = pd.DataFrame(data)
    reverse_map = {v: k for k, v in mapping_dict.items()}
    df = df.rename(columns=reverse_map)
    
    for col in mapping_dict.keys():
        if col not in df.columns:
            df[col] = None
    
    # ★重要: IDカラムを文字列に統一して型不一致を防ぐ
    id_cols = ['person_id', 'activity_id', 'asset_id', 'related_id', 'id']
    for col in id_cols:
        if col in df.columns:
            # 1.0 -> 1 -> "1" のように変換
            df[col] = df[col].apply(lambda x: str(int(float(x))) if x is not None and str(x).replace('.', '', 1).isdigit() else str(x) if x is not None else "")
            
    return df

# ★マスタ取得関数（リストで返す）
def get_master_list(category):
    # マスタテーブルがない場合のエラー回避
    try:
        MAP_MASTER = {'id': 'id', 'カテゴリ': 'category', '名称': 'name', '順序': 'sort_order'}
        df_master = fetch_table("master_options", MAP_MASTER)
        if df_master.empty: return []
        filtered = df_master[df_master['カテゴリ'] == category].copy()
        if filtered.empty: return []
        if '順序' in filtered.columns:
            filtered['順序'] = pd.to_numeric(filtered['順序'], errors='coerce')
            filtered = filtered.sort_values('順序')
        return filtered['名称'].tolist()
    except:
        return []

def insert_data(table_name, data_dict, mapping_dict):
    client = init_supabase()
    db_data = {}
    for jp_key, val in data_dict.items():
        if jp_key in mapping_dict:
            if val == "": val = None
            db_data[mapping_dict[jp_key]] = val
    try:
        client.table(table_name).insert(db_data).execute()
        st.toast("登録しました", icon="✅")
        st.cache_data.clear() # ★登録後に必ずキャッシュクリア
    except Exception as e:
        st.error(f"登録エラー: {e}")

def update_data(table_name, id_col_jp, target_id, data_dict, mapping_dict):
    client = init_supabase()
    db_data = {}
    for jp_key, val in data_dict.items():
        if jp_key in mapping_dict:
            if val == "": val = None
            db_data[mapping_dict[jp_key]] = val
    id_col_en = mapping_dict[id_col_jp]
    try:
        client.table(table_name).update(db_data).eq(id_col_en, target_id).execute()
        st.toast("更新しました", icon="✅")
        st.cache_data.clear() # ★更新後に必ずキャッシュクリア
    except Exception as e:
        st.error(f"更新エラー: {e}")

def delete_data(table_name, id_col_jp, target_id, mapping_dict):
    client = init_supabase()
    id_col_en = mapping_dict[id_col_jp]
    try:
        client.table(table_name).delete().eq(id_col_en, target_id).execute()
        st.toast("削除しました", icon="🗑️")
        st.cache_data.clear() # ★削除後に必ずキャッシュクリア
    except Exception as e:
        st.error(f"削除エラー: {e}")

# --- インポート処理 ---
def process_import(file_obj, table_name, mapping_dict, id_column=None):
    try:
        try:
            df = pd.read_csv(file_obj)
        except UnicodeDecodeError:
            file_obj.seek(0)
            df = pd.read_csv(file_obj, encoding='cp932')
            
        count = 0
        client = init_supabase()
        records = []
        for _, row in df.iterrows():
            db_data = {}
            for jp_k, val in row.items():
                if jp_k in mapping_dict:
                    if pd.isna(val): val = None
                    db_data[mapping_dict[jp_k]] = val
            if id_column and id_column in row:
                db_data[mapping_dict[id_column]] = row[id_column]
            records.append(db_data)

        for rec in records:
            client.table(table_name).upsert(rec).execute()
            count += 1
            
        st.success(f"{count}件 インポート完了")
        st.cache_data.clear()
    except Exception as e:
        st.error(f"インポートエラー: {e}")

# --- ユーティリティ ---
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

# ★IDを安全に文字列化する関数 (照合用)
def to_safe_id(val):
    try:
        # 一度floatにしてからintにし、文字列化 (1.0 -> 1 -> "1")
        return str(int(float(val)))
    except:
        return str(val)

# --- メイン処理 ---
def main():
    if not check_password(): return
    custom_title("成年後見業務支援システム")

    df_persons = fetch_table("persons", MAP_PERSONS)
    
    if '生年月日' in df_persons.columns and not df_persons.empty:
        df_persons['年齢'] = df_persons['生年月日'].apply(calculate_age)
        df_persons['年齢'] = pd.to_numeric(df_persons['年齢'], errors='coerce')

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

    for key in ['selected_person_id', 'delete_confirm_id', 'edit_asset_id', 'delete_asset_id', 'edit_related_id', 'delete_related_id', 'edit_activity_id']:
        if key not in st.session_state: st.session_state[key] = None

    # マスタデータの取得 (動的リスト)
    act_opts = get_master_list('activity') or ["面会", "打ち合わせ", "電話", "メール", "行政手続き", "財産管理", "その他"]
    rel_opts = get_master_list('relationship') or ["親族", "ケアマネ", "施設相談員", "病院SW", "主治医", "弁護士", "行政", "その他"]
    ast_opts = get_master_list('asset') or ["預貯金", "現金", "有価証券", "保険", "不動産", "負債", "その他"]
    guard_opts = get_master_list('guardian_type') or ["後見", "保佐", "補助", "任意", "未成年後見", "その他"]

    # === 1. 利用者情報・活動記録 ===
    if menu == "利用者情報・活動記録":
        df_activities = fetch_table("activities", MAP_ACTIVITIES)
        df_related = fetch_table("related_parties", MAP_RELATED)
        
        custom_header("受任中利用者一覧", help_text="一覧から対象者をクリックすると詳細が表示されます。")
        
        # フィルタ
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

            # キーパーソン
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

            # 基本情報表示
            with st.expander("▼ 基本情報", expanded=True):
                grid_html = f"""
                <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; font-size: 14px;">
                    <div><span style="font-weight:bold; color:#555;">No.:</span> {selected_row.get('ケース番号')}</div>
                    <div><span style="font-weight:bold; color:#555;">事件番号:</span> {selected_row.get('基本事件番号')}</div>
                    <div><span style="font-weight:bold; color:#555;">類型:</span> {selected_row.get('類型')}</div>
                    <div><span style="font-weight:bold; color:#555;">氏名:</span> {selected_row.get('氏名')}</div>
                    <div><span style="font-weight:bold; color:#555;">ｼﾒｲ:</span> {selected_row.get('ｼﾒｲ')}</div>
                    <div><span style="font-weight:bold; color:#555;">生年月日:</span> {selected_row.get('生年月日')}</div>
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
            
            # 活動記録
            st.markdown("### 📝 活動記録")
            with st.expander("➕ 新しい活動記録を追加する", expanded=False):
                with st.form("new_act_form", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    in_date = col1.date_input("活動日", datetime.date.today())
                    in_type = col2.selectbox("活動", act_opts)
                    c3, c4, c5 = st.columns(3)
                    in_time = c3.number_input("所要時間(分)", min_value=0, step=10)
                    in_place = c4.text_input("場所", placeholder="自宅、病院など")
                    in_cost = c5.number_input("費用(円)", min_value=0, step=100)
                    in_note = st.text_area("内容", height=120)
                    in_imp = st.checkbox("★重要")
                    
                    if st.form_submit_button("登録"):
                        new_data = {
                            'person_id': current_pid, '記録日': str(in_date), '活動': in_type,
                            '場所': in_place, '所要時間': in_time, '交通費・立替金': in_cost,
                            '重要': in_imp, '要点': in_note
                        }
                        insert_data("activities", new_data, MAP_ACTIVITIES)
                        st.rerun()

            custom_header("過去の活動履歴", help_text="履歴の「詳細・操作」を開くと編集・削除ができます。")
            if not df_activities.empty:
                # ★修正: ID照合ロジック
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
                                ed_date = st.date_input("活動日", pd.to_datetime(edit_row['記録日']))
                                try:
                                    idx = act_opts.index(edit_row['活動'])
                                except:
                                    idx = 0
                                ed_type = st.selectbox("活動", act_opts, index=idx)
                                c3, c4, c5 = st.columns(3)
                                ed_time = c3.number_input("時間", value=int(edit_row['所要時間'] or 0))
                                ed_place = c4.text_input("場所", value=edit_row['場所'] or "")
                                ed_cost = c5.number_input("費用", value=int(edit_row['交通費・立替金'] or 0))
                                ed_note = st.text_area("内容", value=edit_row['要点'], height=120)
                                ed_imp = st.checkbox("重要", value=bool(edit_row['重要']))
                                
                                c_sv, c_cl = st.columns(2)
                                if c_sv.form_submit_button("保存"):
                                    upd_data = {'記録日': str(ed_date), '活動': ed_type, '場所': ed_place, '所要時間': ed_time, '交通費・立替金': ed_cost, '重要': ed_imp, '要点': ed_note}
                                    update_data("activities", "activity_id", st.session_state.edit_activity_id, upd_data, MAP_ACTIVITIES)
                                    st.session_state.edit_activity_id = None
                                    st.rerun()
                                if c_cl.form_submit_button("キャンセル"):
                                    st.session_state.edit_activity_id = None
                                    st.rerun()

                for _, row in my_acts.iterrows():
                    star = "★" if row['重要'] else ""
                    with st.container(border=True):
                        st.markdown(f"**{star} {row['記録日']}**　📝 {row['活動']}")
                        # 内容を常時表示
                        st.write(row['要点'])
                        
                        with st.expander("詳細・操作", expanded=False):
                            # ★修正: シンプルなマークダウンリストに変更
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
                                    delete_data("activities", "activity_id", row['activity_id'], MAP_ACTIVITIES)
                                    st.session_state.delete_confirm_id = None
                                    st.rerun()
                else:
                    if my_acts.empty:
                        st.write("まだ記録がありません。")

    # --- 2. 関係者・連絡先 ---
    elif menu == "関係者・連絡先":
        custom_header("関係者・連絡先")
        person_opts = {f"{r['氏名']}": r['person_id'] for _, r in df_persons.iterrows()}
        target_name = st.selectbox("対象者", list(person_opts.keys()))
        
        if target_name:
            pid = person_opts[target_name]
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
                        insert_data("related_parties", new_data, MAP_RELATED)
                        st.rerun()
            
            st.markdown("---")
            df_rel = fetch_table("related_parties", MAP_RELATED)
            if not df_rel.empty:
                df_rel['safe_pid'] = df_rel['person_id'].apply(to_safe_id)
                current_pid_safe = to_safe_id(pid)
                my_rel = df_rel[df_rel['safe_pid'] == current_pid_safe]
                
                for _, row in my_rel.iterrows():
                    kp_mark = "★" if row['キーパーソン'] else ""
                    with st.container(border=True):
                        st.markdown(f"**{kp_mark}【{row['関係種別']}】 {row['氏名']}** ({row['所属・名称']})")
                        if row['電話番号']: st.markdown(f"📞 [{row['電話番号']}](tel:{row['電話番号']})")
                        if row['e-mail']: st.markdown(f"✉️ {row['e-mail']}")
                        if row['連携メモ']: st.info(row['連携メモ'])
                        
                        if st.button("削除", key=f"del_rel_{row['related_id']}"):
                            delete_data("related_parties", "related_id", row['related_id'], MAP_RELATED)
                            st.rerun()
            else:
                st.info("登録された関係者はいません。")

    # --- 3. 財産管理 ---
    elif menu == "財産管理":
        custom_header("財産管理")
        person_opts = {f"{r['氏名']}": r['person_id'] for _, r in df_persons.iterrows()}
        target_name = st.selectbox("対象者", list(person_opts.keys()))
        
        if target_name:
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
                        insert_data("assets", nd, MAP_ASSETS)
                        st.rerun()
            
            st.markdown("---")
            df_assets = fetch_table("assets", MAP_ASSETS)
            if not df_assets.empty:
                df_assets['safe_pid'] = df_assets['person_id'].apply(to_safe_id)
                current_pid_safe = to_safe_id(pid)
                my_assets = df_assets[df_assets['safe_pid'] == current_pid_safe]
                
                for _, row in my_assets.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**【{row['財産種別']}】 {row['名称・機関名']}**")
                        st.write(f"額: {row['評価額・残高']} / 場所: {row['保管場所']}")
                        if st.button("削除", key=f"del_ast_{row['asset_id']}"):
                            delete_data("assets", "asset_id", row['asset_id'], MAP_ASSETS)
                            st.rerun()
            else:
                st.info("登録された財産はありません。")

    # --- 4. 利用者情報登録 ---
    elif menu == "利用者情報登録":
        custom_header("利用者情報登録")
        
        with st.expander("➕ 新規登録", expanded=True):
            with st.form("new_person"):
                c1, c2 = st.columns(2)
                p_case = c1.text_input("ケース番号")
                p_name = c1.text_input("氏名")
                p_kana = c2.text_input("カナ")
                p_type = c2.selectbox("類型", guard_opts)
                p_stat = st.selectbox("状態", ["受任中", "終了"])
                if st.form_submit_button("登録"):
                    nd = {'ケース番号': p_case, '氏名': p_name, 'ｼﾒｲ': p_kana, '類型': p_type, '現在の状態': p_stat}
                    insert_data("persons", nd, MAP_PERSONS)
                    st.rerun()
        
        if not df_persons.empty:
            st.markdown("### 登録済み一覧")
            for _, row in df_persons.iterrows():
                with st.expander(f"{row['氏名']} ({row['類型']})"):
                    with st.form(f"edit_p_{row['person_id']}"):
                        try:
                            idx = ["受任中", "終了"].index(row['現在の状態'])
                        except:
                            idx = 0
                        e_stat = st.selectbox("状態", ["受任中", "終了"], index=idx)
                        if st.form_submit_button("更新"):
                            update_data("persons", "person_id", row['person_id'], {'現在の状態': e_stat}, MAP_PERSONS)
                            st.rerun()

    # --- 5. 帳票作成 ---
    elif menu == "帳票作成":
        custom_header("帳票作成")
        uploaded = st.file_uploader("Excelテンプレート")
        if not df_persons.empty:
            target = st.selectbox("対象者", df_persons['氏名'])
            if st.button("作成") and uploaded:
                p_data = df_persons[df_persons['氏名'] == target].iloc[0].to_dict()
                excel = fill_excel_template(uploaded, p_data)
                st.download_button("ダウンロード", excel, f"{target}.xlsx")

    # --- 6. データ管理・移行 (CSVインポート) ---
    elif menu == "データ管理・移行":
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

    # --- 7. 初期設定 ---
    elif menu == "初期設定":
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
                # リスト表示
                current_opts = df_master[df_master['カテゴリ'] == cat_key].sort_values('順序')
                for _, row in current_opts.iterrows():
                    c1, c2 = st.columns([8, 2])
                    c1.write(f"{row['名称']} (順序:{row['順序']})")
                    if c2.button("削除", key=f"del_mst_{row['id']}"):
                        # 使用チェック
                        usage = check_usage_count(cat_key, row['名称'])
                        if usage > 0:
                            st.error(f"「{row['名称']}」は現在 {usage} 件のデータで使用されているため削除できません。")
                        else:
                            delete_data("master_options", "id", row['id'], MAP_MASTER)
                            st.rerun()

                # 追加フォーム
                with st.form(f"add_mst_{cat_key}"):
                    c_name = st.text_input("名称")
                    c_order = st.number_input("順序", min_value=0, value=100)
                    if st.form_submit_button("追加"):
                        if c_name:
                            insert_data("master_options", {'カテゴリ': cat_key, '名称': c_name, '順序': c_order}, MAP_MASTER)
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
                    update_data("app_system_user", "id", curr['id'], nd, MAP_SYSTEM)
                else:
                    insert_data("app_system_user", nd, MAP_SYSTEM)
                st.rerun()

if __name__ == "__main__":
    main()