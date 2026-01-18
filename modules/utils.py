import pandas as pd
import datetime
import re

def normalize_date_str(date_val):
    """
    日付文字列を正規化して 'YYYY-MM-DD' 形式で返す関数
    和暦（明治・大正・昭和・平成・令和）にも対応
    """
    if date_val is None: return ""
    text = str(date_val).strip()
    if not text or text.lower() == "nan": return ""
    
    # 全角数字を半角に
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    
    # 和暦対応
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
    except (ValueError, TypeError):
        return text

def calculate_age(born):
    """
    生年月日から年齢を計算する関数
    """
    if not born: return None
    born_str = normalize_date_str(born)
    if not born_str: return None
    try:
        born_date = pd.to_datetime(born_str, errors='coerce')
        if pd.isna(born_date): return None
        born_date = born_date.date()
        today = datetime.date.today()
        return today.year - born_date.year - ((today.month, today.day) < (born_date.month, born_date.day))
    except (ValueError, TypeError):
        return None

def to_safe_id(val):
    """
    IDを安全な文字列形式に変換する関数
    """
    if pd.isna(val) or val == "":
        return ""
    try:
        return str(int(float(val)))
    except (ValueError, TypeError, OverflowError):
        return str(val)
