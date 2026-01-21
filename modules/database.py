import streamlit as st
import pandas as pd
from supabase import create_client
from .constants import MAP_MASTER
from .utils import to_safe_id
import time

# --- Supabase接続設定 ---
def get_supabase_client():
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except KeyError:
        st.error("【設定エラー】Secretsが見つかりません。.streamlit/secrets.toml を確認してください。")
        st.stop()

@st.cache_resource
def init_supabase():
    """
    Supabaseクライアントを初期化してキャッシュする（互換性のため）
    """
    return get_supabase_client()

@st.cache_data(ttl=600)
def fetch_table(table_name, mapping_dict):
    """
    指定されたテーブルからデータを取得し、DataFrameとして返す
    """
    client = init_supabase()
    try:
        response = client.table(table_name).select("*").execute()
        data = response.data
    except Exception as e:
        # エラー発生時はユーザーに通知しないと原因不明になるため表示（本番ではログへ）
        st.error(f"データ取得エラー ({table_name}): {e}")
        return pd.DataFrame(columns=mapping_dict.keys())
    
    if not data:
        return pd.DataFrame(columns=mapping_dict.keys())
    
    df = pd.DataFrame(data)
    reverse_map = {v: k for k, v in mapping_dict.items()}
    df = df.rename(columns=reverse_map)
    
    for col in mapping_dict.keys():
        if col not in df.columns:
            df[col] = None
    
    id_cols = ['person_id', 'activity_id', 'asset_id', 'related_id', 'id']
    for col in id_cols:
        if col in df.columns:
            df[col] = df[col].apply(to_safe_id)
            
    return df

def get_master_list(category):
    """
    マスタデータから選択肢リストを取得する
    """
    try:
        df_master = fetch_table("master_options", MAP_MASTER)
        if df_master.empty: return []
        filtered = df_master[df_master['カテゴリ'] == category].copy()
        if filtered.empty: return []
        if '順序' in filtered.columns:
            filtered['順序'] = pd.to_numeric(filtered['順序'], errors='coerce')
            filtered = filtered.sort_values('順序')
        return filtered['名称'].tolist()
    except Exception:
        return []

def check_usage_count(category, option_name):
    """
    マスタデータの選択肢が使用されている数をチェックする
    """
    client = init_supabase()
    count = 0
    try:
        if category == 'activity':
            res = client.table('activities').select('activity_id', count='exact').eq('activity_type', option_name).execute()
            count = res.count
        elif category == 'asset':
            res = client.table('assets').select('asset_id', count='exact').eq('asset_type', option_name).execute()
            count = res.count
        elif category == 'relationship':
            res = client.table('related_parties').select('related_id', count='exact').eq('relationship', option_name).execute()
            count = res.count
        elif category == 'guardian_type':
            res = client.table('persons').select('person_id', count='exact').eq('guardianship_type', option_name).execute()
            count = res.count
    except Exception:
        pass
    return count

def insert_data(table_name, data_dict, mapping_dict):
    """
    データの新規登録を行う
    """
    client = init_supabase()
    db_data = {}
    for jp_key, val in data_dict.items():
        if jp_key in mapping_dict:
            if val == "": val = None
            db_data[mapping_dict[jp_key]] = val
    try:
        # print(f"DEBUG: DB Insert -> {table_name}, Data={db_data}")
        client.table(table_name).insert(db_data).execute()
        st.toast("登録しました", icon="✅")
        time.sleep(1) # DB反映待ち
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"登録エラー: {e}")
        return False

def update_data(table_name, id_col_jp, target_id, data_dict, mapping_dict):
    """
    データの更新を行う
    """
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
        time.sleep(1) # DB反映待ち
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"更新エラー: {e}")
        return False

def delete_data(table_name, id_col_jp, target_id, mapping_dict):
    """
    データの削除を行う
    """
    client = init_supabase()
    id_col_en = mapping_dict[id_col_jp]
    try:
        client.table(table_name).delete().eq(id_col_en, target_id).execute()
        st.toast("削除しました", icon="🗑️")
        time.sleep(1) # DB反映待ち
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"削除エラー: {e}")
        return False

def process_import(file_obj, table_name, mapping_dict, id_column=None):
    """
    CSV/Excelファイルからのインポート処理
    """
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
