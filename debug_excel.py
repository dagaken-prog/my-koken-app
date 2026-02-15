
import io
import openpyxl
from modules.report_generator import create_periodic_report, ROW_ASSET_BANK_START

# ダミーデータ
person_data = {
    '氏名': 'テスト 太郎',
    '住所': '東京都千代田区1-1',
    '居所': 'テスト施設',
    '〒': '100-0001',
    '家裁報告月': '8月'
}
guardian_data = {
    '氏名': '後見 花子',
    '住所': '東京都港区',
    '連絡先電話番号': '03-1234-5678'
}
asset_list = [
    {'財産種別': '預貯金', '名称・機関名': 'テスト銀行', '支店・詳細': '本店', '口座番号・記号': '1111', '評価額・残高': 1000000},
    {'財産種別': '預貯金', '名称・機関名': 'サンプル信用金庫', '支店・詳細': '駅前', '口座番号・記号': '2222', '評価額・残高': 50000},
    {'財産種別': '現金', '名称・機関名': '現金', '評価額・残高': 30000},
]

template_path = "template_report.xlsx"

try:
    # 生成実行
    output, err = create_periodic_report(person_data, guardian_data, asset_list, [], template_path)
    
    if err:
        print(f"Error: {err}")
    else:
        # ファイル保存
        with open("debug_output.xlsx", "wb") as f:
            f.write(output.getvalue())
        print("Success: debug_output.xlsx created.")
        
        # 検証 (読み直して値確認)
        wb = openpyxl.load_workbook("debug_output.xlsx", data_only=False)
        
        # 基本情報シート
        ws_rep = wb["後見事務報告書(定期報告)"]
        print(f"Name (V5): {ws_rep['V5'].value}")
        print(f"Address (S1): {ws_rep['S1'].value}")
        
        # 財産目録シート
        ws_ast = wb["財産目録"]
        print(f"Asset Name (W2): {ws_ast['W2'].value}")
        
        # 銀行1 (Row 25)
        print(f"Bank1 Name (C25): {ws_ast['C25'].value}")
        print(f"Bank1 Val (X25): {ws_ast['X25'].value}")
        
        # 現金 (Row 37) - *まだ実装していないので None か Templateの値が出るはず*
        print(f"Cash (X37): {ws_ast['X37'].value}")
        
        # クリア確認 (Row 27以降、データは2つなのでRow 29は空のはず)
        # TemplateにはRow 29にデータがあった場合、クリアされていればNone
        print(f"Empty Row (C29): {ws_ast['C29'].value}")

except Exception as e:
    print(f"Exception: {e}")
