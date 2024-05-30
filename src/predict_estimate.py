import pandas as pd
import numpy as np
import joblib

# モデルの読み込み
model_path = 'models/crew_estimate_model.pkl'
model = joblib.load(model_path)

new_data_path = 'data/new_estimate_progress.xlsx'
sheet_name = '単発見積依頼_from依頼者 のコピー'
data = pd.read_excel(new_data_path, sheet_name=sheet_name, header=3)

# データのカラム名を表示
print(data.columns)

# データの先頭を確認
print(data.head())

# 必要なカラムを選択
columns = [
    '現調必要\n有無', '間取り', '対象面積(バルコニー含む) ※単位㎡', '部屋番号', '洗浄希望_エアコン台数',
    'ロフトの有無', '施工完了期日\n（計算式）', '施工完了期日（ある場合）', '見積\n担当', 'ステータス',
    '既存予算（見積金額など）（税別）', '他社見積金額 （税別）', '現地写真や他社見積\n（↓URLを記入）',
    '備考（部屋の状況や進捗状況）', '見積期限', '見積\n提出日', '更新日', '日付＆パートナー打診状況'
]

new_data = data[columns]

# カラム名の修正
new_data.columns = [
    "inspection_required", "layout", "area_sqm", "room_number", "ac_units",
    "loft", "completion_date_formula", "completion_date", "estimator", "status",
    "existing_budget", "competitor_estimate", "photo_or_other_estimate",
    "remarks", "estimate_deadline", "estimate_submission_date", "update_date", "partner_contact_status"
]

# カテゴリカルデータのエンコード
categorical_columns = ["inspection_required", "loft", "status", "estimator"]
new_data = pd.get_dummies(new_data, columns=categorical_columns, drop_first=True)

# モデルの訓練時に使用した特徴量と同じカラムを揃える
model_features = model.feature_names_in_
for col in model_features:
    if col not in new_data.columns:
        new_data[col] = 0

new_data = new_data[model_features]

# 「見積(クルー)金額(税別)」の値を予測
predicted_prices = model.predict(new_data)

# 予測結果を元のデータに追加
data['見積(クルー)金額(税別)'] = predicted_prices

# 結果を新しいExcelファイルに保存
output_path = 'data/predicted_estimates.xlsx'
data.to_excel(output_path, index=False)

print(f"Predicted estimates saved to {output_path}")
