import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
import joblib

# データの読み込み
file_path = 'data/output/cleaned_data.xlsx'
data = pd.read_excel(file_path)

# 使用するカラムを選択
features = [
    "現調必要\n有無", "間取り", "対象面積(バルコニー含む) ※単位㎡", "部屋番号", "洗浄希望_エアコン台数",
    "ロフトの有無", "施工完了期日\n（計算式）", "施工完了期日（ある場合）", "見積\n担当", "ステータス",
    "既存予算（見積金額など）（税別）", "他社見積金額 （税別）", "現地写真や他社見積\n（↓URLを記入）",
    "備考（部屋の状況や進捗状況）", "見積期限", "見積提出日", "更新日", "日付＆パートナー打診状況"
]

target = "見積(クルー)金額(税別)"

# データのクリーニング
def convert_layout_to_numeric(layout):
    if isinstance(layout, str):
        if layout.endswith('K') or layout.endswith('D') or layout.endswith('L') or layout.endswith('S'):
            return layout[:-1]
    return layout

data['間取り'] = data['間取り'].apply(convert_layout_to_numeric)

# 空のセルを0で埋める
data.fillna(0, inplace=True)

# 数値データの変換
for feature in features:
    if data[feature].dtype == 'object':
        data[feature] = data[feature].astype('str').str.extract('(\d+)').astype(float)

# 特徴量と目的変数の分割
X = data[features]
y = data[target]

# 訓練データとテストデータの分割
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# モデルの作成
model = RandomForestRegressor(n_estimators=100, random_state=42)

# モデルの訓練
model.fit(X_train, y_train)

# モデルの評価
y_pred = model.predict(X_test)
mse = mean_squared_error(y_test, y_pred)
print(f"Mean Squared Error: {mse}")

# モデルの保存
model_path = 'models/crew_estimate_model.pkl'
joblib.dump(model, model_path)
print(f"Model saved to {model_path}")
