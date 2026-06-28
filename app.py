# -*- coding: utf-8 -*-
"""
1週間 負荷推移 可視化 自動化 — アプリ(ドラッグ&ドロップUI)改良版
開発指示書(第1版)準拠 + 改良要望反映。

改良点(従来):
  1. total: SPD MX のみ「曜日合計」ではなく「曜日平均」で算出(エンジン側)。
  2. total に含める曜日をサイドバーで選択可能。
  3. 棒(平均/total)のみ数値ラベルを表示(エンジン側)。
  4. データ未投入の曜日も負荷量0として全曜日(月〜日)をグラフ表示。
  5. 縦軸の目盛り幅(刻み)を指標ごとに指定可能。

改良点(今回追加):
  A. 各曜日で「採用するSession名」を選べるようにした。
     - 投入したファイルから Session の一覧を自動検出し、曜日ごとに選択(複数可)。
     - 既定では all / full 系のSessionを選択。無ければ全Sessionを既定選択。
     - 選択行だけを抽出してエンジンへ渡すため、ROUND ALL など任意のSession名で集計可能。
  B. Name列(その他の列も)が検出されないことがある問題を改善。
     - CSVのBOM、前後空白、引用符、大文字小文字の差を吸収して正規化。
     - これにより「'\\ufeff"Name"'」のような列名でも Name として認識される。

※ load_engine.py は変更していません。app.py 側の前処理で上記A・Bを実現しています。
   選択した行は、エンジンの採用条件に合わせて Session を 'all' に書き換えてから渡します
   (このため出力CSVの明細では Session 列が 'all' と表示されます。元のSession名は画面で確認できます)。

起動:  streamlit run app.py
"""
import os, io, re, zipfile, tempfile
import streamlit as st
import pandas as pd

import load_engine as eng

st.set_page_config(page_title="1週間 負荷推移 可視化", layout="wide")
st.title("1週間 負荷推移 可視化 自動化")
st.caption("曜日の箱にデータ(Excel / Numbers / CSV)をドラッグして投入してください。"
           "日付はファイル名から自動取得します。投入後、曜日ごとに採用するSessionを選べます。")

WEEKDAYS = eng.WEEKDAYS_JP          # ['月','火','水','木','金','土','日']
RAW_COLUMNS = eng.RAW_COLUMNS       # 必須22列


# ======================================================================
#  前処理ヘルパー(改良点A・B)
# ======================================================================
def _norm_key(s) -> str:
    """列名・Session名の比較用キー。BOM・引用符・前後空白・連続空白・大小文字を吸収。"""
    if s is None:
        return ""
    t = str(s).replace("\ufeff", "")          # BOM除去
    t = t.strip().strip('"').strip("'").strip()  # 引用符・空白除去
    t = re.sub(r"\s+", " ", t)                 # 連続空白を1つに
    return t.lower()


def _read_raw(path: str, ext: str) -> pd.DataFrame:
    """拡張子ごとに素のDataFrameを読み込む。CSV/TSVはBOM対応(utf-8-sig)。"""
    if ext == ".numbers":
        from numbers_parser import Document
        doc = Document(path)
        rows = doc.sheets[0].tables[0].rows(values_only=True)
        header = list(rows[0])
        data = [list(r) for r in rows[1:]]
        return pd.DataFrame(data, columns=header)
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, engine="openpyxl" if ext != ".xls" else None)
    if ext == ".tsv":
        return pd.read_csv(path, sep="\t", encoding="utf-8-sig")
    if ext == ".csv":
        # encoding='utf-8-sig' でBOMを除去 → 先頭列の引用符も正しく解釈される
        return pd.read_csv(path, encoding="utf-8-sig")
    raise ValueError(f"未対応のファイル形式です: {ext}")


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """実列名を正規化し、必須22列(RAW_COLUMNS)の正式名へマッピングして揃える。"""
    canon = {_norm_key(c): c for c in RAW_COLUMNS}   # 正規化キー -> 正式列名
    rename = {}
    used_targets = set()
    for c in df.columns:
        key = _norm_key(c)
        if key in canon and canon[key] not in used_targets:
            rename[c] = canon[key]
            used_targets.add(canon[key])
    return df.rename(columns=rename)


@st.cache_data(show_spinner=False)
def read_normalized(name: str, data: bytes) -> pd.DataFrame:
    """アップロード内容を読み込み、列名を正規化して返す(bytesをキーにキャッシュ)。"""
    ext = os.path.splitext(name)[1].lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    try:
        tmp.write(data)
        tmp.close()
        df = _read_raw(tmp.name, ext)
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
    return _normalize_columns(df)


def sessions_of(df: pd.DataFrame):
    """DataFrameに含まれるSession名の一覧(出現順・重複/空除外)。"""
    if "Session" not in df.columns:
        return []
    out, seen = [], set()
    for v in df["Session"].tolist():
        if v is None:
            continue
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            continue
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def default_sessions(sessions):
    """既定選択: all完全一致 または full前方一致。無ければ全Session。"""
    match = [s for s in sessions
             if _norm_key(s) == "all" or _norm_key(s).startswith("full")]
    return match if match else list(sessions)


# ======================================================================
#  サイドバー: 設定(指示書 6章)
# ======================================================================
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


# ======================================================================
#  7曜日のドラッグ&ドロップ箱 + Session選択(改良点A)
# ======================================================================
st.subheader("データ投入(月〜日)")
uploaded = {}
selections = {}    # 曜日 -> 選んだSession名のリスト
read_errors = {}   # 曜日 -> 読み込み時の問題メッセージ

cols = st.columns(7)
for i, wd in enumerate(WEEKDAYS):
    with cols[i]:
        st.markdown(f"**{wd}**")
        f = st.file_uploader(
            wd, type=["xlsx", "xls", "numbers", "csv", "tsv"],
            key=f"up_{wd}", label_visibility="collapsed")
        uploaded[wd] = f
        selections[wd] = []

        if f is None:
            continue

        # 投入ファイルを読み込み、Session一覧を提示
        try:
            df = read_normalized(f.name, bytes(f.getbuffer()))
        except Exception as e:
            read_errors[wd] = f"{wd}曜({f.name}) の読み込みに失敗: {e}"
            st.error("読み込み失敗")
            continue

        missing = [c for c in RAW_COLUMNS if c not in df.columns]
        if missing:
            read_errors[wd] = (f"{wd}曜({f.name}) に必要な列が見つかりません: "
                               f"{'、'.join(missing)}")
            st.warning("列不足")

        sess = sessions_of(df)
        if not sess:
            st.caption("Sessionが検出できませんでした")
            continue

        fid = f"{f.name}_{getattr(f, 'size', len(df))}"  # ファイルが変わると選択をリセット
        selections[wd] = st.multiselect(
            "採用Session", options=sess,
            default=[s for s in default_sessions(sess) if s in sess],
            key=f"sess_{wd}_{fid}",
            help="この曜日で集計に使うSessionを選びます(複数選択可)。")

run = st.button("▶ 集計・可視化を実行", type="primary")


# ======================================================================
#  実行
# ======================================================================
if run:
    # 読み込み段階のエラーがあれば中断
    if read_errors:
        st.error("ファイルに問題があります。修正してください:")
        for msg in read_errors.values():
            st.write("・", msg)
        st.stop()

    if not any(uploaded.values()):
        st.error("少なくとも1つのファイルを投入してください。")
        st.stop()

    tmpdir = tempfile.mkdtemp(prefix="load_")
    paths = []
    used_sessions_note = []   # 画面表示用(曜日: 採用Session)
    pre_warnings = []

    for wd in WEEKDAYS:
        f = uploaded[wd]
        if f is None:
            continue
        sel_sessions = selections.get(wd, [])
        if not sel_sessions:
            pre_warnings.append(f"{wd}曜: 採用Sessionが未選択のためスキップしました。")
            continue

        df = read_normalized(f.name, bytes(f.getbuffer()))

        # 選択Sessionの行を抽出(前後空白を無視して一致判定)
        sel_keys = {s.strip() for s in sel_sessions}
        mask = df["Session"].apply(lambda v: str(v).strip() in sel_keys)
        sub = df[mask].copy()
        if sub.empty:
            pre_warnings.append(
                f"{wd}曜: 選択したSession({'、'.join(sel_sessions)})の行が0件でした。スキップします。")
            continue

        # エンジンの採用条件(all/full)に合わせて Session を 'all' に統一
        sub["Session"] = "all"

        # 日付トークン(_YYYYMMDD_)を保持するため、元のファイル名(stem)で .xlsx 出力
        stem = os.path.splitext(f.name)[0]
        outp = os.path.join(tmpdir, f"{stem}.xlsx")
        try:
            sub.to_excel(outp, index=False)
        except Exception as e:
            st.error(f"{wd}曜({f.name}) の中間ファイル書き出しに失敗: {e}")
            st.stop()
        paths.append(outp)
        used_sessions_note.append(f"{wd}: {'、'.join(sel_sessions)}")

    if not paths:
        st.error("集計対象のデータがありません。"
                 "各曜日で採用するSessionが選択されているか確認してください。")
        for w in pre_warnings:
            st.warning(w)
        st.stop()

    outdir = os.path.join(tmpdir, "output")
    with st.spinner("処理中..."):
        res = eng.process(paths, threshold=threshold_pct / 100.0,
                          yaxis_overrides=yaxis_overrides,
                          ytick_overrides=ytick_overrides,
                          total_weekdays=total_weekdays,
                          show_all_weekdays=show_all_weekdays,
                          outdir=outdir)

    # 前処理段階の注意
    for w in pre_warnings:
        st.warning(w)

    # エンジンのエラー(処理中断)
    if res.errors:
        st.error("エラーのため中断しました:")
        for e in res.errors:
            st.write("・", e)
        st.stop()

    # エンジンの警告
    for w in res.warnings:
        st.warning(w)

    st.success(f"完了: {len(res.day_order)}日分 / 採用 {len(res.used_df)} 行 / "
               f"total対象 {len(res.total_players)} 名")
    if used_sessions_note:
        st.caption("採用Session — " + " / ".join(used_sessions_note))
    if res.total_weekdays:
        st.caption("total 集計曜日: " + "、".join(WEEKDAYS[w] for w in res.total_weekdays))
    if res.total_players:
        st.caption("total対象選手(total集計曜日すべてに有効データが残った選手): "
                   + "、".join(res.total_players))

    # --- CSVプレビュー & ダウンロード ---
    st.subheader("表(CSV)")
    df_out = pd.read_csv(res.csv_path)
    st.dataframe(df_out, use_container_width=True, height=360)
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
