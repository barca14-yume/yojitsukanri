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
        # 集計処理開始（デバッグ表示削除済み）
        # 予算データ読込
        try:
            budget_df = pd.read_excel(BUDGET_SAVE_PATH, skiprows=6)
            # st.write("DEBUG: 予算データ読み込み", budget_df.head())
        except Exception as e:
            st.error(f"予算ファイル読込エラー: {e}")
            return
        budget_subject_col = [col for col in budget_df.columns if '科目' in str(col)]

        if budget_subject_col:
            budget_subject_col = budget_subject_col[0]

        else:
            st.error(f"予算ファイルに科目名列が見つかりません: {budget_df.columns.tolist()}")
            return
        months = [col for col in budget_df.columns if col not in [budget_subject_col, 'Unnamed: 13']]
        # デバッグ用出力（削除済み）
        # 実績データ読込
        actual_data = {}
        import re
        for afile in saved_actual_files:
            try:
                df = pd.read_excel(afile, skiprows=6)
                # デバッグ用出力（削除済み）
            except Exception as e:
                st.error(f"実績ファイル読込エラー({afile}): {e}")
                return
            col_candidates = [col for col in df.columns if '科目' in str(col).replace(' ', '').replace('　', '')]

            if col_candidates:
                subject_col = col_candidates[0]

            else:
                st.error(f"実績ファイルに科目名列が見つかりません: {df.columns.tolist()}")
                return
            # インデックス正規化
            df[subject_col] = df[subject_col].astype(str).str.strip().str.replace(' ', '').str.replace('　', '')
            df = df.set_index(subject_col)

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
        # --- 販売費及び一般管理費の明細分析 ---
        # 5月の実績データから「販売費及び一般管理費」配下のみ抽出
        sg_subject = "販売費及び一般管理費"
        may_file = actual_file_map.get("5月", None)
        may_col = actual_col_map.get("5月", None)
        apr_file = actual_file_map.get("4月", None)

        def render_card(subject, row4, row5):
            icon_map = {
                '売上高': '💸',
                '売上総利益': '📈',
                '販売費及び一般管理費': '💼',
                '経常利益': '📈',
                '原価率(%)': '⚙️',
                '販管費率(%)': '🧾',
            }
            badge = ''
            if subject in ['売上高', '経常利益']:
                badge = '<span style="background:#ffd700;color:#444;font-size:0.85em;padding:2px 8px 2px 8px;border-radius:10px;margin-left:8px;">重要</span>'
            icon = icon_map.get(subject, '')
            def format_block(label, value, color, is_rate=False, is_diff=False):
                empty = (value is None or value == '' or value == 'None')
                # 色は全て黒（#222）、ただし差額でマイナスのみ赤
                base_color = '#222'
                if is_diff and not empty:
                    try:
                        v = float(value)
                        if v < 0:
                            base_color = '#ff1744'  # 鮮やかな赤
                    except:
                        pass
                style = f"background:{'#f0f5fa' if empty else '#fff'};border-radius:8px;padding:8px 12px;margin-bottom:3px;min-width:90px;box-shadow:0 1px 3px #e3e8f0;"
                val_style = f"font-size:1.25rem;font-weight:bold;color:{'#aaa' if empty else base_color};display:flex;align-items:center;gap:2px;"
                label_style = "font-size:0.93rem;color:#555;letter-spacing:0.01em;"
                # 金額系はカンマ区切り
                if not empty and not is_rate:
                    try:
                        val = f"{int(float(value)):,}"
                    except:
                        val = value
                else:
                    val = value if not empty else "-"
                # 率系は%を強調
                if is_rate and not empty:
                    val = f"{value}<span style='font-size:1.08rem;color:{base_color};margin-left:2px;'>%</span>"
                return f'<div style="{style}"><div style="{label_style}">{label}</div><div style="{val_style}">{val}</div></div>'
            return f'<div class="card-hover" style="border-radius:13px;padding:22px 18px 18px 18px;margin-bottom:22px;background:linear-gradient(90deg,#eaf2fb 60%,#f8faff 100%);box-shadow:0 3px 12px #a3bffa18;max-width:560px;margin-left:auto;margin-right:auto;"><div style="font-size:1.17rem;font-weight:700;color:#28427a;margin-bottom:12px;letter-spacing:0.01em;display:flex;align-items:center;gap:6px;">{icon} {subject}{badge}</div><div class="card-flex"><div style="flex:1;min-width:170px;"><div style="font-size:1.01rem;color:#2b7cff;font-weight:600;margin-bottom:4px;">4月</div>{format_block('実績', row4.get('実績'), '#2b7cff')}{format_block('予算', row4.get('予算'), '#28427a')}{format_block('差額', row4.get('差額'), '#c0392b', is_diff=True)}{format_block('対予算比', row4.get('対予算比'), '#1abc9c', is_rate=True)}{format_block('前年比', row4.get('前年比'), '#8e44ad', is_rate=True)}</div><div style="flex:1;min-width:170px;"><div style="font-size:1.01rem;color:#00b383;font-weight:600;margin-bottom:4px;">5月</div>{format_block('実績', row5.get('実績'), '#00b383')}{format_block('予算', row5.get('予算'), '#28427a')}{format_block('差額', row5.get('差額'), '#c0392b', is_diff=True)}{format_block('対予算比', row5.get('対予算比'), '#1abc9c', is_rate=True)}{format_block('前年比', row5.get('前年比'), '#8e44ad', is_rate=True)}</div></div></div>'
        # --- CSSを1回だけグローバルに出す ---
        st.markdown("""
        <style>
        .card-flex {display:flex;gap:18px;justify-content:space-between;flex-wrap:wrap;}
        @media (max-width: 600px) { .card-flex {flex-direction:column;gap:8px;} }
        .card-hover:hover {box-shadow:0 6px 24px #6a8cff33;transform:translateY(-2px);transition:0.2s;}
        </style>
        """, unsafe_allow_html=True)
        # 指標ごとにカードで表示
        html_cards = ""
        for idx, row in result_df.iterrows():
            subject = row["科目名"]
            row4 = {
                '実績': row.get('4月_実績', ''),
                '予算': row.get('4月_予算', ''),
                '差額': row.get('4月_差額', ''),
                '対予算比': row.get('4月_対予算比%', ''),
                '前年比': row.get('4月_前年比%', ''),
            }
            row5 = {
                '実績': row.get('5月_実績', ''),
                '予算': row.get('5月_予算', ''),
                '差額': row.get('5月_差額', ''),
                '対予算比': row.get('5月_対予算比%', ''),
                '前年比': row.get('5月_前年比%', ''),
            }
            html_cards += render_card(subject, row4, row5)
        # st.write("DEBUG: html_cards 内容", html_cards)
        html = "<div style='display:grid;gap:8px;'>" + html_cards + "</div>"
        if html_cards.strip():
            st.markdown(html, unsafe_allow_html=True)
        # st.write("DEBUG: st.markdown(unsafe_allow_html=True) 実行済み")
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

        # --- 販売費及び一般管理費の明細ピックアップ（月別・前年比/前月比110%以上・90%以下） ---
        for target_month, target_file, target_col, prev_col in [
            ("4月", actual_file_map.get("4月", None), actual_col_map.get("4月", None), None),
            ("5月", actual_file_map.get("5月", None), actual_col_map.get("5月", None), prev_actual_col_map.get(("5月", actual_file_map.get("5月", None)), None))
        ]:
            pick_rows = []
            if target_file and target_col:
                df = actual_data[target_file]
                # 前年データは5月のみ
                df_prev = actual_data[target_file] if prev_col and target_month == "5月" else None
                # 前月データ（4月はなし）
                df_prev_month = actual_data[actual_file_map.get("4月", None)] if target_month == "5月" and actual_file_map.get("4月", None) else None
                prev_month_col = actual_col_map.get("4月", None) if target_month == "5月" else None
                # 役員報酬～雑費の範囲だけピックアップ
                idx_list = list(df.index)
                try:
                    start_idx = idx_list.index("役員報酬")
                    end_idx = idx_list.index("雑費")
                    if start_idx > end_idx:
                        detail_subjects = []
                    else:
                        detail_subjects = idx_list[start_idx:end_idx+1]
                except ValueError:
                    detail_subjects = []
                for subject in detail_subjects:
                    try:
                        val = df.at[subject, target_col]
                        # 前年比（5月のみ）
                        if target_month == "5月" and df_prev is not None and prev_col and subject in df_prev.index:
                            prev_val = df_prev.at[subject, prev_col]
                            yoy = round(float(val) / float(prev_val) * 100, 1) if prev_val not in [None, "", 0] and val not in [None, "", 0] else None
                        else:
                            yoy = None
                        # 前月比（5月のみ）
                        if target_month == "5月" and df_prev_month is not None and prev_month_col and subject in df_prev_month.index:
                            prev_month_val = df_prev_month.at[subject, prev_month_col]
                            mom = round(float(val) / float(prev_month_val) * 100, 1) if prev_month_val not in [None, "", 0] and val not in [None, "", 0] else None
                        else:
                            mom = None
                        pick_rows.append({
                            "科目名": subject,
                            "金額": f"{int(float(val)):,}" if val not in [None, "", 0] else None,
                            "前年比": yoy,
                            "前月比": mom
                        })
                    except:
                        continue
            # 高い順・低い順でソート
            if pick_rows:
                df_pick = pd.DataFrame(pick_rows)
                st.markdown(f"#### 販売費及び一般管理費の明細（{target_month}） 前年比・前月比ピックアップ")
                col1, col2 = st.columns(2)
                with col1:
                    if target_month == "4月":
                        st.markdown("**金額が110%以上**")
                        df_over_110 = df_pick[(df_pick["金額"].notnull()) & (df_pick["金額"].apply(lambda x: int(str(x).replace(',', '')) if x not in [None, "", 0] else 0) >= 110)]
                        st.dataframe(df_over_110.sort_values("金額", ascending=False)[["科目名", "金額"]], use_container_width=True)
                        st.markdown("**金額が90%以下**")
                        df_under_90 = df_pick[(df_pick["金額"].notnull()) & (df_pick["金額"].apply(lambda x: int(str(x).replace(',', '')) if x not in [None, "", 0] else 0) <= 90)]
                        st.dataframe(df_under_90.sort_values("金額", ascending=True)[["科目名", "金額"]], use_container_width=True)
                    else:
                        st.markdown("**前年比が110%以上**")
                        df_over_110 = df_pick[(df_pick["前年比"].notnull()) & (df_pick["前年比"] >= 110)]
                        st.dataframe(df_over_110.sort_values("前年比", ascending=False)[["科目名", "金額", "前年比", "前月比"]], use_container_width=True)
                        st.markdown("**前年比が90%以下**")
                        df_under_90 = df_pick[(df_pick["前年比"].notnull()) & (df_pick["前年比"] <= 90)]
                        st.dataframe(df_under_90.sort_values("前年比", ascending=True)[["科目名", "金額", "前年比", "前月比"]], use_container_width=True)
                with col2:
                    if target_month == "4月":
                        st.markdown("")
                    else:
                        st.markdown("**前月比が110%以上**")
                        df_mom_over_110 = df_pick[(df_pick["前月比"].notnull()) & (df_pick["前月比"] >= 110)]
                        st.dataframe(df_mom_over_110.sort_values("前月比", ascending=False)[["科目名", "金額", "前年比", "前月比"]], use_container_width=True)
                        st.markdown("**前月比が90%以下**")
                        df_mom_under_90 = df_pick[(df_pick["前月比"].notnull()) & (df_pick["前月比"] <= 90)]
                        st.dataframe(df_mom_under_90.sort_values("前月比", ascending=True)[["科目名", "金額", "前年比", "前月比"]], use_container_width=True)

if __name__ == "__main__":
    main()
