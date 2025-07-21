import streamlit as st
import os
import glob
import pandas as pd  # コメントアウトされていたので有効化

# 結果保持用
final_result = pd.DataFrame()

# タイトルと説明
st.title("コストダウン対象品目 抽出ツール")
st.caption("CSVファイルをもとに条件に合致する品目を抽出・集計します")

# 入力フォーム
with st.form("input_form"):
    col1, col2, col3 = st.columns(3)
    with col1:
        min_count = st.number_input("最小発注回数（年）", min_value=1, value=3, step=1)
    with col2:
        min_quantity = st.number_input("最小発注数量（年）", min_value=1, value=100, step=1)
    with col3:
        min_suppliers = st.number_input("最小仕入先数（年）", min_value=1, value=2, step=1)

    submitted = st.form_submit_button("抽出実行")

if submitted:
    try:
        # CSVファイルを1件取得
        csv_files = glob.glob('./data/*.csv')
        if not csv_files:
            raise FileNotFoundError("CSVファイルが見つかりません。./data フォルダを確認してください。")
        file_path = csv_files[0]  # 最初のCSVファイルを使用

        df = pd.read_csv(file_path, encoding='cp932', sep=',')

        # クォート・空白削除
        df.columns = df.columns.str.strip().str.replace('"', '')

        # フィルタ処理
        df = df[~df['購入先'].isin(['10258XD', '14264XH', '13007XI'])]
        df = df[~df['品目説明'].str.contains('送料', na=False)]

        # 日付変換＆年度追加
        df['納期'] = pd.to_datetime(df['納期'], errors='coerce')
        df.dropna(subset=['納期'], inplace=True)
        df['年度'] = df['納期'].apply(lambda x: x.year if x.month >= 10 else x.year - 1)

        # 発注金額
        df['発注金額'] = df['発注'] * df['品目原価']

        # 年度別に集計
        grouped = df.groupby(['品目説明', '年度']).agg(
            発注回数=('納期', 'count'),
            発注数量=('発注', 'sum'),
            仕入先数=('購入先', 'nunique')
        ).reset_index()

        qualified = grouped[
            (grouped['発注回数'] >= min_count) &
            (grouped['発注数量'] >= min_quantity) &
            (grouped['仕入先数'] >= min_suppliers)
        ]['品目説明'].unique()

        filtered_df = df[df['品目説明'].isin(qualified)]

        summary = filtered_df.groupby('品目説明').agg(
            仕入先の数=('購入先', 'nunique'),
            合計発注回数=('納期', 'count'),
            合計発注数量=('発注', 'sum'),
            合計発注金額=('発注金額', 'sum')
        ).reset_index()

        summary.sort_values(by='合計発注金額', ascending=False, inplace=True)

        # 表示
        st.success(f"{len(summary)} 件の品目が抽出されました（合計発注金額の降順で表示）")
        st.dataframe(summary.style.format({
            '合計発注数量': '{:,.0f}',
            '合計発注金額': '{:,.0f}'
        }))

        # Excel出力
        if st.button("Excelに出力"):
            os.makedirs('./output', exist_ok=True)
            file_name = f"対象一覧_回数{min_count}_数量{min_quantity}_仕入先{min_suppliers}.xlsx"
            output_path = os.path.join('./output', file_name)
            summary.to_excel(output_path, index=False)
            st.success(f"Excelファイルを保存しました： {output_path}")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
