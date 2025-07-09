import streamlit as st
import pandas as pd
import io

def to_thousand_yen(x):
    try:
        return round(float(x) / 1000)
    except:
        return ""

def main():
    st.markdown("""
# 予算・実績 自動集計システム
---
""")

    with st.expander("❓ 使い方ガイド", expanded=True):
        st.markdown("""
        1. **予算ファイル（1つ）・実績ファイル（複数）をアップロード**
        2. 保存済みファイルの確認や削除も可能
        3. ファイルが揃うと自動で集計・プレビュー
        4. 集計結果はExcelでダウンロードできます
        """)

    import os
    BUDGET_SAVE_PATH = "予算保存用.xlsx"
    actual_dir = "actuals"
    os.makedirs(actual_dir, exist_ok=True)

    # --- ファイルアップロードUI ---
    with st.expander("📤 ファイルアップロード・管理", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("予算ファイル")
            budget_file = st.file_uploader("予算ファイルをアップロード", type=["xlsx"], key="budget")
            use_saved_budget = os.path.exists(BUDGET_SAVE_PATH)
            if use_saved_budget:
                st.success(f"現在の予算ファイル: {BUDGET_SAVE_PATH}")
            if budget_file:
                with open(BUDGET_SAVE_PATH, "wb") as f:
                    f.write(budget_file.getbuffer())
                st.success(f"予算ファイルを保存しました: {BUDGET_SAVE_PATH}")
                use_saved_budget = True
        with col2:
            st.subheader("実績ファイル（複数可）")
            actual_file = st.file_uploader("実績ファイルをアップロード", type=["xlsx"], accept_multiple_files=True, key="actual")
            if actual_file:
                for afile in actual_file:
                    save_path = os.path.join(actual_dir, afile.name)
                    with open(save_path, "wb") as f:
                        f.write(afile.getbuffer())
                st.success(f"{len(actual_file)}件の実績ファイルを保存しました。")

        st.markdown("---")
        saved_actual_files = [os.path.join(actual_dir, f) for f in os.listdir(actual_dir) if f.endswith(".xlsx")]
        st.info(f"保存済み実績ファイル: {[os.path.basename(f) for f in saved_actual_files]}")
        if saved_actual_files:
            st.subheader("実績ファイルの削除")
            files_to_delete = st.multiselect("削除したい実績ファイルを選択", [os.path.basename(f) for f in saved_actual_files])
            if st.button("選択したファイルを削除"):
                for fname in files_to_delete:
                    fpath = os.path.join(actual_dir, fname)
                    if os.path.exists(fpath):
                        os.remove(fpath)
                st.success(f"{len(files_to_delete)}件のファイルを削除しました。画面を再読み込みしてください。")
        st.markdown("---")

    if use_saved_budget and saved_actual_files:
        st.success("ファイルがアップロードされました。自動集計を開始します。")
        st.markdown("---")
        st.subheader("集計結果プレビュー（4月・5月のみ）")
        st.markdown(
            "<div style='background-color:#f0f2f6;border-radius:8px;padding:10px 16px 10px 16px;margin-bottom:8px;'>"
            "<b>アップロード済みのファイルに基づき、4月・5月の主要指標を集計しています。下記テーブルは横スクロール・高さ制限付きで閲覧できます。</b>"
            "</div>",
            unsafe_allow_html=True
        )
        # 予算データ読込
        budget_df = pd.read_excel(BUDGET_SAVE_PATH, skiprows=6)
        budget_subject_col = [col for col in budget_df.columns if '科目' in str(col)]
        if budget_subject_col:
            budget_subject_col = budget_subject_col[0]
        else:
            st.error(f"予算ファイルに科目名列が見つかりません: {budget_df.columns.tolist()}")
            return
        months = [col for col in budget_df.columns if col not in [budget_subject_col, 'Unnamed: 13']]
        # 実績データ読込
        actual_data = {}
        import re
        for afile in saved_actual_files:
            df = pd.read_excel(afile, skiprows=6)
            col_candidates = [col for col in df.columns if '科目' in str(col).replace(' ', '').replace('　', '')]
            if col_candidates:
                subject_col = col_candidates[0]
            else:
                st.error(f"実績ファイルに科目名列が見つかりません: {df.columns.tolist()}")
                return
            # インデックス正規化
            df[subject_col] = df[subject_col].astype(str).str.strip().str.replace(' ', '').str.replace('　', '')
            df = df.set_index(subject_col)
            # デバッグ: インデックス一覧と販売費および一般管理費の存在確認
            
            
            m = re.search(r'PL_(\d{4})年(\d{1,2})月', os.path.basename(afile))
            if m:
                month = f"{int(m.group(2))}月"
            else:
                month = os.path.basename(afile)
            actual_data[month] = df
        # 実績カラム名マッピング
        actual_file_map = {}
        actual_col_map = {}
        for k, df in actual_data.items():
            for month in months:
                found = False
                for col in df.columns:
                    if '2025年' in col and month in col and '実績金額(発生)' in col:
                        actual_file_map[month] = k
                        actual_col_map[month] = col
                        found = True
                        break
                if not found:
                    for col in df.columns:
                        if '2024年' in col and month in col and '実績金額(発生)' in col:
                            actual_file_map[month] = k
                            actual_col_map[month] = col
                            break
        # 集計
        # 必要な科目名だけ抽出
        needed_subjects = ["売上高", "売上総利益", "販売費及び一般管理費", "経常利益"]
        result = []
        # 予算ファイルにある全科目名を取得
        all_subjects = list(budget_df[budget_subject_col].unique())
        # 必要な科目が抜けていれば追加
        for subj in needed_subjects:
            if subj not in all_subjects:
                all_subjects.append(subj)
        # 月ごとの前年科目名マッピング
        prev_actual_col_map = {}
        for k, df in actual_data.items():
            for month in months:
                for col in df.columns:
                    if '2024年' in col and month in col and '実績金額(発生)' in col:
                        prev_actual_col_map[(month, k)] = col
        for subject in needed_subjects:
            row = {"科目名": subject}
            for month in months:
                # 予算
                budget = budget_df.loc[budget_df[budget_subject_col] == subject, month].values[0] if subject in list(budget_df[budget_subject_col]) else ""
                # 実績（当年）
                actual = ""
                # 実績ファイル・カラムがあれば取得（なければ空欄）
                if month in actual_file_map and month in actual_col_map:
                    df = actual_data[actual_file_map[month]]
                    col = actual_col_map[month]
                    # インデックスに完全一致しない場合、stripや全角半角・空白除去で部分一致を試みる
                    subject_norm = subject.strip().replace(' ', '').replace('　', '').replace('および', '及び')
                    matched_index = None
                    for idx in df.index:
                        idx_norm = str(idx).strip().replace(' ', '').replace('　', '').replace('および', '及び')
                        if subject_norm == idx_norm:
                            matched_index = idx
                            break
                    if matched_index is not None:
                        try:
                            actual = df.at[matched_index, col]
                        except:
                            actual = ""
                    else:
                        actual = ""
                # 実績（前年）
                prev_actual = ""
                if (month, actual_file_map.get(month, None)) in prev_actual_col_map:
                    prev_df = actual_data[actual_file_map[month]]
                    prev_col = prev_actual_col_map[(month, actual_file_map[month])]
                    subject_norm = subject.strip().replace(' ', '').replace('　', '')
                    matched_index = None
                    for idx in prev_df.index:
                        idx_norm = str(idx).strip().replace(' ', '').replace('　', '')
                        if subject_norm == idx_norm:
                            matched_index = idx
                            break
                    if matched_index is not None:
                        try:
                            prev_actual = prev_df.at[matched_index, prev_col]
                        except Exception as e:
                            st.error(f"前年実績データ取得エラー: {e}")
                            prev_actual = ""
                budget_ = to_thousand_yen(budget)
                actual_ = to_thousand_yen(actual)
                prev_actual_ = to_thousand_yen(prev_actual)
                diff = "" if budget_ == "" or actual_ == "" else actual_ - budget_
                rate = "" if budget_ in ["", 0] or actual_ == "" else round(actual_ / budget_ * 100, 1)
                yoy = "" if prev_actual_ in ["", 0] or actual_ == "" else round(actual_ / prev_actual_ * 100, 1)
                row[f"{month}_予算"] = budget_
                row[f"{month}_実績"] = actual_
                row[f"{month}_対予算比%"] = rate
                row[f"{month}_前年比%"] = yoy
                row[f"{month}_差額"] = diff
            result.append(row)
        result_df = pd.DataFrame(result)

        # 原価率・販管費率の計算
        def safe_div(num, denom):
            try:
                if denom in ["", 0, None] or num in ["", None]:
                    return ""
                if (isinstance(num, float) and math.isnan(num)) or (isinstance(denom, float) and math.isnan(denom)):
                    return ""
                return f"{round(float(num) / float(denom) * 100, 1)}%"
            except:
                return ""
        def safe_zero(x):
            return 0 if x in ["", None] else x
        summary_rows = []
        # サマリー用カラムをDataFrameに必ず追加
        for month in months:
            if f"{month}_原価率%" not in result_df.columns:
                result_df[f"{month}_原価率%"] = ""
            if f"{month}_販管費率%" not in result_df.columns:
                result_df[f"{month}_販管費率%"] = ""
        # サマリー行用の全カラムで空文字初期化
        all_columns = list(result_df.columns)
        genka_row = {col: "" for col in all_columns}
        han_kan_row = {col: "" for col in all_columns}
        genka_row["科目名"] = "原価率(%)"
        han_kan_row["科目名"] = "販管費率(%)"
        for month in months:
            if f"{month}_実績" in result_df.columns:
                try:
                    sales = safe_zero(result_df[result_df["科目名"] == "売上高"][f"{month}_実績"].values[0])
                    gross_profit = safe_zero(result_df[result_df["科目名"] == "売上総利益"][f"{month}_実績"].values[0])
                    cost = sales - gross_profit
                    vals = result_df[result_df["科目名"] == "販売費及び一般管理費"][f"{month}_実績"].values
                    sg_and_a = vals[0] if len(vals) > 0 and vals[0] not in [None, "", 0] else 0
                    genka = safe_div(cost, sales)
                    han_kan = safe_div(sg_and_a, sales)
                    genka_row[f"{month}_原価率%"] = genka
                    han_kan_row[f"{month}_販管費率%"] = han_kan
                    # 4月・5月の実績列にも値を入れる
                    if month in ["4月", "5月"]:
                        genka_row[f"{month}_実績"] = genka
                        han_kan_row[f"{month}_実績"] = han_kan
                except Exception as e:
                    genka_row[f"{month}_原価率%"] = ""
                    han_kan_row[f"{month}_販管費率%"] = ""
        # 経常利益の直後（index=4,5）にサマリー行を挿入
        keijou_idx = result_df.index[result_df["科目名"] == "経常利益"].tolist()
        if keijou_idx:
            insert_at = keijou_idx[0] + 1
        else:
            insert_at = 4
        # サマリー行をresult_df.columns順に並べてから挿入
        summary_df = pd.DataFrame([genka_row, han_kan_row])[result_df.columns]
        result_df = pd.concat([
            result_df.iloc[:insert_at],
            summary_df,
            result_df.iloc[insert_at:]
        ], ignore_index=True)

        # 4月・5月のみのカラムを抽出し、「原価率%」「販管費率%」カラムは除外
        keep_months = ["4月", "5月"]
        keep_cols = ["科目名"]
        for m in keep_months:
            keep_cols += [c for c in result_df.columns if c.startswith(m+"_") and not (c.endswith("原価率%") or c.endswith("販管費率%"))]
        # カラムを絞り込む
        result_df = result_df[keep_cols]
        st.markdown(
            "<div style='border:2px solid #A3BFFA;border-radius:8px;padding:8px 12px 8px 12px;background-color:#ffffff;'>",
            unsafe_allow_html=True
        )
        st.dataframe(result_df, use_container_width=True, height=360)
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown(":blue[↓ 集計結果をExcelでダウンロード ↓]")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
        st.download_button(
            label="集計結果をExcelでダウンロード",
            data=output.getvalue(),
            file_name="月次予実表集計結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown("---")

if __name__ == "__main__":
    main()
