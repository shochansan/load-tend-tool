# -*- coding: utf-8 -*-
"""
1週間 負荷推移 可視化 自動化 — アプリ(ドラッグ&ドロップUI)改良版
開発指示書(第1版)準拠 + 改良要望反映。

改良点:
  1. total: SPD MX のみ「曜日合計」ではなく「曜日平均」で算出(エンジン側)。
  2. total に含める曜日をサイドバーで選択可能。
  3. 棒(平均/total)のみ数値ラベルを表示(エンジン側)。
  4. データ未投入の曜日も負荷量0として全曜日(月〜日)をグラフ表示。
  5. 縦軸の目盛り幅(刻み)を指標ごとに指定可能。

起動:  streamlit run app.py
"""
import os, io, zipfile, tempfile
import streamlit as st
import pandas as pd

import load_engine as eng

st.set_page_config(page_title="1週間 負荷推移 可視化", layout="wide")
st.title("1週間 負荷推移 可視化 自動化")
st.caption("曜日の箱にデータ(Excel / Numbers)をドラッグして投入してください。日付はファイル名から自動取得します。")

WEEKDAYS = eng.WEEKDAYS_JP  # ['月','火','水','木','金','土','日']

# ---- サイドバー: 設定(指示書 6章) ----------------------------------------
st.sidebar.header("設定")
threshold_pct = st.sidebar.number_input(
    "外れ値しきい値 ±X% (Distance)", min_value=0.0, max_value=100.0,
    value=30.0, step=1.0,
    help="完了時点の平均Aに対する変化率 V=(P−A)/A がこの範囲を超えた行を外れ値として除外します。")

# ---- total 設定(改良点2) -------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("total の設定")
total_weekdays_jp = st.sidebar.multiselect(
    "total に含める曜日", options=WEEKDAYS, default=WEEKDAYS,
    help="ここで選んだ曜日のデータだけを total に集計します。"
         "選んだ全曜日に有効データが揃った選手のみが total 対象になります。"
         "(SPD MX は曜日平均、その他の指標は曜日合計で算出します)")
total_weekdays = [WEEKDAYS.index(w) for w in total_weekdays_jp]
if not total_weekdays:
    st.sidebar.warning("total に含める曜日が未選択です。total は空になります。")

# ---- グラフ表示設定(改良点4) --------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("グラフ表示")
show_all_weekdays = st.sidebar.checkbox(
    "全曜日(月〜日)を表示し、未投入日は0とする", value=True,
    help="オフにすると、データを投入した曜日だけをグラフに表示します。")

# ---- 縦軸の手動指定(改良点5: 目盛り幅を追加) -----------------------------
st.sidebar.subheader("グラフ縦軸(任意)")
st.sidebar.caption("空欄なら自動。下限・上限・目盛り幅(刻み)を指標ごとに指定できます。"
                   "目盛り幅だけの指定も可能です。")
yaxis_overrides = {}
ytick_overrides = {}
with st.sidebar.expander("縦軸の手動指定を開く"):
    for m in eng.METRIC_COLUMNS:
        st.markdown(f"**{m}**")
        c1, c2, c3 = st.columns(3)
        lo = c1.text_input("下限", key=f"lo_{m}", placeholder="auto")
        hi = c2.text_input("上限", key=f"hi_{m}", placeholder="auto")
        step = c3.text_input("目盛り幅", key=f"st_{m}", placeholder="auto")

        if lo.strip() != "" and hi.strip() != "":
            try:
                yaxis_overrides[m] = (float(lo), float(hi))
            except ValueError:
                st.sidebar.warning(f"{m} の下限・上限が数値ではありません。無視します。")
        elif lo.strip() != "" or hi.strip() != "":
            st.sidebar.warning(f"{m} は下限・上限を両方指定してください。片方のみは無視します。")

        if step.strip() != "":
            try:
                sv = float(step)
                if sv > 0:
                    ytick_overrides[m] = sv
                else:
                    st.sidebar.warning(f"{m} の目盛り幅は正の数で指定してください。無視します。")
            except ValueError:
                st.sidebar.warning(f"{m} の目盛り幅が数値ではありません。無視します。")

# ---- 7曜日のドラッグ&ドロップ箱(指示書 2.2) -------------------------------
st.subheader("データ投入(月〜日)")
uploaded = {}
cols = st.columns(7)
for i, wd in enumerate(WEEKDAYS):
    with cols[i]:
        st.markdown(f"**{wd}**")
        uploaded[wd] = st.file_uploader(
            wd, type=["xlsx", "xls", "numbers", "csv", "tsv"],
            key=f"up_{wd}", label_visibility="collapsed")

run = st.button("▶ 集計・可視化を実行", type="primary")

# ---- 実行 -----------------------------------------------------------------
if run:
    files = [f for f in uploaded.values() if f is not None]
    if not files:
        st.error("少なくとも1つのファイルを投入してください。")
        st.stop()

    # アップロードを一時ディレクトリに元のファイル名で保存(日付取得のため)
    tmpdir = tempfile.mkdtemp(prefix="load_")
    paths = []
    for f in files:
        p = os.path.join(tmpdir, f.name)
        with open(p, "wb") as out:
            out.write(f.getbuffer())
        paths.append(p)

    outdir = os.path.join(tmpdir, "output")
    with st.spinner("処理中..."):
        res = eng.process(paths, threshold=threshold_pct / 100.0,
                          yaxis_overrides=yaxis_overrides,
                          ytick_overrides=ytick_overrides,
                          total_weekdays=total_weekdays,
                          show_all_weekdays=show_all_weekdays,
                          outdir=outdir)

    # エラー(処理中断)
    if res.errors:
        st.error("エラーのため中断しました:")
        for e in res.errors:
            st.write("・", e)
        st.stop()

    # 警告
    for w in res.warnings:
        st.warning(w)

    st.success(f"完了: {len(res.day_order)}日分 / 採用 {len(res.used_df)} 行 / "
               f"total対象 {len(res.total_players)} 名")
    if res.total_weekdays:
        st.caption("total 集計曜日: " + "、".join(WEEKDAYS[w] for w in res.total_weekdays))
    if res.total_players:
        st.caption("total対象選手(total集計曜日すべてに有効データが残った選手): "
                   + "、".join(res.total_players))

    # --- CSVプレビュー & ダウンロード ---
    st.subheader("表(CSV)")
    df = pd.read_csv(res.csv_path)
    st.dataframe(df, use_container_width=True, height=360)
    with open(res.csv_path, "rb") as fh:
        st.download_button("⬇ CSVをダウンロード", fh.read(),
                           file_name="負荷推移_集計.csv", mime="text/csv")

    # --- グラフ ---
    st.subheader("グラフ")
    with open(res.pdf_path, "rb") as fh:
        st.download_button("⬇ グラフ(PDF・12指標まとめ)をダウンロード", fh.read(),
                           file_name="負荷推移_グラフ.pdf", mime="application/pdf")

    # PNG一括ZIP
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in res.png_paths:
            zf.write(p, arcname=os.path.basename(p))
    st.download_button("⬇ グラフ(PNG・12枚)をZIPでダウンロード", zbuf.getvalue(),
                       file_name="負荷推移_グラフPNG.zip", mime="application/zip")

    # 画面表示(2列)
    gcols = st.columns(2)
    for idx, p in enumerate(res.png_paths):
        gcols[idx % 2].image(p, use_container_width=True)
