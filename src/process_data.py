import os
import pandas as pd

# プロジェクトのルートディレクトリを取得
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 絶対パスを作成
file_path = os.path.join(project_root, 'data', 'estimate_progress.xlsx')
output_path = os.path.join(project_root, 'data', 'output', 'cleaned_data.xlsx')

# 実際のシート名を使用
sheet_name = '単発見積依頼_from依頼者 のコピー'

# ヘッダー行を指定してデータを読み込む
data = pd.read_excel(file_path, sheet_name=sheet_name, header=3)

# データの最初の数行を表示
print(data.head())
print(data.columns)

# データをフィルタリングおよび整形
filtered_data = data[data['希望メニュー'] == 'ルームクリーニング'].copy()

def convert_to_sqm(value):
    if isinstance(value, str):
        if '坪' in value:
            try:
                num = float(value.replace('坪', '').strip())
                return num * 3.30579
            except ValueError:
                return None
        elif '畳' in value:
            try:
                num = float(value.replace('畳', '').strip())
                return num * 1.65289
            except ValueError:
                return None
        else:
            try:
                return float(value.strip())
            except ValueError:
                return None
    return value

filtered_data['対象面積(バルコニー含む) ※単位㎡'] = filtered_data['対象面積(バルコニー含む) ※単位㎡'].apply(convert_to_sqm)

# 整形したデータを出力
filtered_data.to_excel(output_path, index=False)
print("Data processed and saved to", output_path)
