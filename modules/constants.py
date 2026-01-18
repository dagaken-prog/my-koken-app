# 定数・マッピング定義

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

MAP_MASTER = {
    'id': 'id', 'カテゴリ': 'category', '名称': 'name', '順序': 'sort_order'
}

# 逆引き用辞書
R_MAP_PERSONS = {v: k for k, v in MAP_PERSONS.items()}
R_MAP_ACTIVITIES = {v: k for k, v in MAP_ACTIVITIES.items()}
R_MAP_ASSETS = {v: k for k, v in MAP_ASSETS.items()}
R_MAP_RELATED = {v: k for k, v in MAP_RELATED.items()}
R_MAP_SYSTEM = {v: k for k, v in MAP_SYSTEM.items()}
