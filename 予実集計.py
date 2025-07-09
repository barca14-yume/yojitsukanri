import pandas as pd
import glob
import os
import traceback
import sys

print("=== 予実集計スクリプト 開始 ===")
try:
    print("1. 予算ファイル読込中...")
    budget_file = "2025予算.xlsx"
    budget_df = pd.read_excel(budget_file, skiprows=6)
    print(f"   予算カラム名一覧: {budget_df.columns.tolist()}")
    # 科目名列を自動検出
    budget_subject_col = [col for col in budget_df.columns if '科目' in str(col)]
    if budget_subject_col:
        budget_subject_col = budget_subject_col[0]
    else:
        raise ValueError(f"予算ファイルに科目名列が見つかりません: {budget_df.columns.tolist()}")
    print("2. 実績ファイル検索中...")
    dir_path = os.path.dirname(os.path.abspath(__file__))
    actual_files = sorted(glob.glob(os.path.join(dir_path, "PL_2025年*.xlsx")))
    print(f"   検出ファイル: {actual_files}")

    def extract_month(filename):
        basename = os.path.basename(filename)
        return basename.replace("PL_", "").replace(".xlsx", "")

    actual_data = {}
    for afile in actual_files:
        month = extract_month(afile)
        print(f"3. 実績ファイル読込中: {afile} → {month}")
        df = pd.read_excel(afile, skiprows=6)
        print(f"   カラム名一覧: {df.columns.tolist()}")
        # 「科目名」列を自動検出（空白や表記揺れ対応）
        col_candidates = [col for col in df.columns if '科目' in str(col)]
        if col_candidates:
            subject_col = col_candidates[0]
        else:
            raise ValueError(f"科目名列が見つかりません: {df.columns.tolist()}")
        actual_data[month] = df.set_index(subject_col)

    # 月カラム抽出（例：4月, 5月, ...）
    months = [col for col in budget_df.columns if col not in [budget_subject_col, 'Unnamed: 13']]
    print(f"4. 月リスト: {months}")

    # 月ごとに該当する実績ファイル（actual_dataのキー）をマッピング
    actual_file_map = {}
    actual_col_map = {}
    for k, df in actual_data.items():
        for month in months:
            # まず2025年のカラムを優先
            found = False
            for col in df.columns:
                if '2025年' in col and month in col and '実績金額(発生)' in col:
                    actual_file_map[month] = k
                    actual_col_map[month] = col
                    found = True
                    break
            # 2025年がなければ2024年を使う
            if not found:
                for col in df.columns:
                    if '2024年' in col and month in col and '実績金額(発生)' in col:
                        actual_file_map[month] = k
                        actual_col_map[month] = col
                        break
    print(f"5. 実績カラム対応: {actual_col_map}")

    def to_thousand_yen(x):
        try:
            return round(float(x) / 1000)
        except:
            return ""

    result = []
    for subject in budget_df[budget_subject_col]:
        row = {"科目名": subject}
        for month in months:
            budget = budget_df.loc[budget_df[budget_subject_col] == subject, month].values[0]
            actual = ""
            # 実績ファイル・カラムがあれば取得（なければ空欄）
            if month in actual_file_map and month in actual_col_map:
                df = actual_data[actual_file_map[month]]
                col = actual_col_map[month]
                if subject in df.index:
                    try:
                        actual = df.at[subject, col]
                    except Exception:
                        actual = ""
            budget_ = to_thousand_yen(budget)
            actual_ = to_thousand_yen(actual)
            diff = "" if budget_ == "" or actual_ == "" else actual_ - budget_
            rate = "" if budget_ in ["", 0] or actual_ == "" else round(actual_ / budget_ * 100, 1)
            row[f"{month}_予算"] = budget_
            row[f"{month}_実績"] = actual_
            row[f"{month}_差額"] = diff
            row[f"{month}_達成率"] = rate
        result.append(row)

    result_df = pd.DataFrame(result)
    result_df.to_excel("月次予実表.xlsx", index=False)
    print("=== 予実集計スクリプト 正常終了 ===")

except Exception as e:
    print("エラーが発生しました:", file=sys.stderr)
    traceback.print_exc()

# 四半期集計
def quarter_sum(df, months, subject_col="科目名"):
    q = [months[i:i+3] for i in range(0, len(months), 3)]
    q_result = []
    for subject in df[subject_col]:
        row = {subject_col: subject}
        for idx, q_months in enumerate(q, 1):
            budget_sum = actual_sum = 0
            for m in q_months:
                b = df.loc[df[subject_col] == subject, f"{m}_予算"].values[0]
                a = df.loc[df[subject_col] == subject, f"{m}_実績"].values[0]
                budget_sum += b if b != "" else 0
                actual_sum += a if a != "" else 0
            diff = actual_sum - budget_sum if budget_sum != 0 else ""
            rate = round(actual_sum / budget_sum * 100, 1) if budget_sum != 0 else ""
            row[f"Q{idx}_予算合計"] = budget_sum if budget_sum != 0 else ""
            row[f"Q{idx}_実績合計"] = actual_sum if actual_sum != 0 else ""
            row[f"Q{idx}_差額"] = diff
            row[f"Q{idx}_達成率"] = rate
        q_result.append(row)
    return pd.DataFrame(q_result)

quarter_df = quarter_sum(result_df, months)
quarter_df.to_excel("四半期予実集計.xlsx", index=False)

# 年間進捗
annual_result = []
for subject in result_df["科目名"]:
    budget_total = actual_total = 0
    for month in months:
        b = result_df.loc[result_df["科目名"] == subject, f"{month}_予算"].values[0]
        a = result_df.loc[result_df["科目名"] == subject, f"{month}_実績"].values[0]
        budget_total += b if b != "" else 0
        actual_total += a if a != "" else 0
    diff = actual_total - budget_total if budget_total != 0 else ""
    rate = round(actual_total / budget_total * 100, 1) if budget_total != 0 else ""
    # 年間見込み（単純月平均×12）
    month_count = sum(1 for m in months if result_df.loc[result_df["科目名"] == subject, f"{m}_実績"].values[0] != "")
    avg = actual_total / month_count if month_count else 0
    forecast = round(avg * 12) if avg else ""
    annual_result.append({
        "科目名": subject,
        "年間予算": budget_total if budget_total != 0 else "",
        "実績累計": actual_total if actual_total != 0 else "",
        "差額": diff,
        "進捗率": rate,
        "年間見込み（単純月平均×12）": forecast
    })
annual_df = pd.DataFrame(annual_result)
annual_df.to_excel("年間進捗集計.xlsx", index=False)
