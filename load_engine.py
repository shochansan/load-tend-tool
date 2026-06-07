# -*- coding: utf-8 -*-
"""
1週間 負荷推移 可視化 自動化 — 処理エンジン(改良版)
開発指示書(第1版)準拠 + 改良要望反映。

改良点:
  1. total: SPD MX のみ「各曜日の合計」ではなく「各曜日の平均」で算出。
  2. total: total に含める曜日を選択可能(total_weekdays)。
  3. グラフ: 棒(平均/total)にのみ数値ラベルを表示。
  4. グラフ: データ未投入の曜日も負荷量0として全曜日(月〜日)を表示。
  5. グラフ: 縦軸の目盛り幅(刻み)を指標ごとに指定可能(ytick_overrides)。

依存: pandas, numpy, matplotlib, openpyxl, numbers-parser
"""
from __future__ import annotations
import matplotlib
matplotlib.rcParams['pdf.fonttype'] = 42
matplotlib.rcParams['ps.fonttype'] = 42
import os, re, datetime
from dataclasses import dataclass, field
import numpy as np
import pandas as pd

# ---- 定数 -----------------------------------------------------------------
RAW_COLUMNS = ['Name','Position','Session','Duration','Duration_TF','Distance',
               'SPD MX','SI_D','HI_D','Sprint','Accel_Z2','Accel_Z3','Decel_Z2',
               'Decel_Z3','DIST/min','Rest Time','Dist Mode','Sprint Reach Time',
               'Peak Sprint','DEVICE','DEVICE_VER','CSV_VERSION']

DERIVED_COLUMNS = ['Accel_All', 'Decel_All', 'HighAgility']

# グラフ・集計対象の12指標
METRIC_COLUMNS = ['Distance','SPD MX','SI_D','HI_D','Sprint','Accel_Z2','Accel_Z3',
                  'Decel_Z2','Decel_Z3','Accel_All','Decel_All','HighAgility']

# total を「合計」ではなく「平均」で求める指標(改良点1)
# 例: SPD MX(最高速度)は週合計に意味がないため、選手内では曜日平均を採用する。
TOTAL_MEAN_METRICS = {'SPD MX'}

WEEKDAYS_JP = ['月','火','水','木','金','土','日']   # 0=月 ... 6=日
DEFAULT_THRESHOLD = 0.30  # ±30%

FNAME_DATE_RE = re.compile(r'_(\d{8})_')


# ---- 戻り値の入れ物 --------------------------------------------------------
@dataclass
class ProcessResult:
    csv_path: str = ""
    pdf_path: str = ""
    png_paths: list = field(default_factory=list)
    used_df: pd.DataFrame = None        # 採用された明細(外れ値除外後)
    daily_avg: pd.DataFrame = None       # 曜日別平均
    total: pd.Series = None              # total(選手間平均)
    total_players: list = field(default_factory=list)
    total_weekdays: list = field(default_factory=list)  # total に採用した曜日(週順index)
    day_order: list = field(default_factory=list)   # データのある曜日(週順)
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)


# ---- ファイル読み込み ------------------------------------------------------
def _read_numbers(path):
    from numbers_parser import Document
    doc = Document(path)
    rows = doc.sheets[0].tables[0].rows(values_only=True)
    header = list(rows[0])
    data = [list(r) for r in rows[1:]]
    return pd.DataFrame(data, columns=header)


def _read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.numbers':
        return _read_numbers(path)
    if ext in ('.xlsx', '.xlsm', '.xls'):
        return pd.read_excel(path, engine='openpyxl' if ext != '.xls' else None)
    if ext in ('.csv', '.tsv'):
        sep = '\t' if ext == '.tsv' else None
        return pd.read_csv(path, sep=sep, engine='python')
    raise ValueError(f"未対応のファイル形式です: {ext}")


def _date_from_filename(path):
    m = FNAME_DATE_RE.search(os.path.basename(path))
    if not m:
        raise ValueError(f"ファイル名から日付(YYYYMMDD)を取得できません: {os.path.basename(path)}")
    d = m.group(1)
    return datetime.date(int(d[:4]), int(d[4:6]), int(d[6:8]))


# ---- セッション判定(指示書 2.4) -------------------------------------------
def _is_target_session(value) -> bool:
    if value is None:
        return False
    s = str(value).strip().lower()
    return s == 'all' or s.startswith('full')


# ---- メイン処理 ------------------------------------------------------------
def process(files, threshold=DEFAULT_THRESHOLD, yaxis_overrides=None,
            ytick_overrides=None, total_weekdays=None, show_all_weekdays=True,
            outdir='output', basename='負荷推移'):
    """
    files            : 入力ファイルパスのリスト
    threshold        : 外れ値しきい値(小数。0.30 = ±30%)
    yaxis_overrides  : {指標名: (下限, 上限)} の辞書。指定指標のみ縦軸固定
    ytick_overrides  : {指標名: 目盛り幅} の辞書。指定指標のみ縦軸の刻みを固定(改良点5)
    total_weekdays   : total に含める曜日index(0=月..6=日)の集合/リスト。
                       None なら全曜日。投入の無い曜日は自動的に対象外(改良点2)。
    show_all_weekdays: True で月〜日の全曜日をグラフに表示し、未投入日は0とする(改良点4)
    """
    os.makedirs(outdir, exist_ok=True)
    yaxis_overrides = yaxis_overrides or {}
    ytick_overrides = ytick_overrides or {}
    res = ProcessResult()

    # --- 1. 読み込み & 曜日割り当て -------------------------------------
    per_day_used = {}   # weekday_idx -> DataFrame(採用・外れ値除外後)
    day_dates = {}      # weekday_idx -> date
    seen_weekday = {}

    for path in files:
        try:
            d = _date_from_filename(path)
        except ValueError as e:
            res.errors.append(str(e)); continue
        wd = d.weekday()  # 0=月

        # 同一曜日への複数投入はエラー(指示書 2.2)
        if wd in seen_weekday:
            res.errors.append(
                f"同一曜日({WEEKDAYS_JP[wd]})に複数ファイルが投入されました: "
                f"{os.path.basename(seen_weekday[wd])} と {os.path.basename(path)}")
            continue
        seen_weekday[wd] = path

        try:
            df = _read_table(path)
        except Exception as e:
            res.errors.append(f"{os.path.basename(path)} の読み込みに失敗: {e}"); continue

        # 必須列チェック
        missing = [c for c in RAW_COLUMNS if c not in df.columns]
        if missing:
            res.errors.append(f"{os.path.basename(path)} に必要な列がありません: {missing}")
            continue

        # --- 対象セッション行の抽出(2.4) ---
        sel = df[df['Session'].apply(_is_target_session)].copy()
        if sel.empty:
            res.warnings.append(f"{WEEKDAYS_JP[wd]}({d}) は対象セッション(all/full)が0件のため、データなし扱いです。")
            continue

        # 同一選手に複数採用行 → エラー検知(2.4)
        dup = sel['Name'][sel['Name'].duplicated()].unique().tolist()
        if dup:
            res.errors.append(f"{WEEKDAYS_JP[wd]}({d}) で同一選手に複数の採用行があります: {dup}")
            continue

        # --- 追加列(3.1) ---
        for c in METRIC_COLUMNS:
            if c in RAW_COLUMNS:
                sel[c] = pd.to_numeric(sel[c], errors='coerce')
        sel['Accel_Z2'] = pd.to_numeric(sel['Accel_Z2'], errors='coerce')
        sel['Accel_Z3'] = pd.to_numeric(sel['Accel_Z3'], errors='coerce')
        sel['Decel_Z2'] = pd.to_numeric(sel['Decel_Z2'], errors='coerce')
        sel['Decel_Z3'] = pd.to_numeric(sel['Decel_Z3'], errors='coerce')
        sel['Accel_All']   = sel['Accel_Z2'] + sel['Accel_Z3']
        sel['Decel_All']   = sel['Decel_Z2'] + sel['Decel_Z3']
        sel['HighAgility'] = sel['Accel_Z3'] + sel['Decel_Z3']

        # --- 外れ値除外(3.2) ---
        sel['Distance'] = pd.to_numeric(sel['Distance'], errors='coerce')
        valid = sel[sel['Distance'].notna()].copy()
        dropped_nan = len(sel) - len(valid)
        if dropped_nan:
            res.warnings.append(f"{WEEKDAYS_JP[wd]}({d}) で Distance が数値でない採用行を {dropped_nan} 件除外しました。")

        A = valid['Distance'].mean()              # 除外前の全採用行平均(1回算出)
        if A and not np.isnan(A) and A != 0:
            V = (valid['Distance'] - A) / A
            used = valid[V.abs() <= threshold].copy()
        else:
            used = valid.copy()                   # A=0等の異常時は除外しない

        used['日付'] = d
        used['曜日'] = WEEKDAYS_JP[wd]
        used['_wd'] = wd
        per_day_used[wd] = used
        day_dates[wd] = d

    if res.errors:
        return res  # エラーがあれば中断(指示書 2.2/7)

    if not per_day_used:
        res.errors.append("有効なデータが1日もありません。")
        return res

    # --- 表示する曜日(週順) ---
    day_order = sorted(per_day_used.keys())          # 0..6(データのある曜日)
    res.day_order = day_order

    used_all = pd.concat([per_day_used[w] for w in day_order], ignore_index=True)
    res.used_df = used_all

    # --- 3. 曜日別平均(3.3) ---
    daily_avg = (used_all.groupby('_wd')[METRIC_COLUMNS].mean()
                 .reindex(day_order))
    res.daily_avg = daily_avg

    # --- 4. total(3.4 + 改良点1,2) ---
    if total_weekdays is None:
        total_set = set(range(7))
    else:
        total_set = set(int(w) for w in total_weekdays)
    # 実際に total に使う曜日 = 「選択された曜日」かつ「データがある曜日」
    total_days = [w for w in day_order if w in total_set]
    res.total_weekdays = total_days

    if not total_days:
        res.total = pd.Series({c: np.nan for c in METRIC_COLUMNS})
        res.total_players = []
        res.warnings.append("total 対象の曜日が選択されていない(または該当日にデータがない)ため、total は空(NaN)です。")
    else:
        # total 対象曜日すべてに有効データが残る選手のみ対象
        players_per_day = {w: set(per_day_used[w]['Name']) for w in total_days}
        full_players = set.intersection(*players_per_day.values()) if players_per_day else set()
        res.total_players = sorted(full_players)

        if full_players:
            sub = used_all[(used_all['Name'].isin(full_players)) &
                           (used_all['_wd'].isin(total_days))]
            grp = sub.groupby('Name')
            tvals = {}
            for c in METRIC_COLUMNS:
                # SPD MX 等は選手内で「曜日平均」、それ以外は「曜日合計」
                per_player = grp[c].mean() if c in TOTAL_MEAN_METRICS else grp[c].sum()
                tvals[c] = per_player.mean()        # 選手間平均
            res.total = pd.Series(tvals)
        else:
            res.total = pd.Series({c: np.nan for c in METRIC_COLUMNS})
            res.warnings.append("total対象曜日すべてに揃った選手がいないため、total は空(NaN)です。")

    # --- 5. CSV出力(4章) ---
    res.csv_path = _write_csv(res, day_order, day_dates, outdir, basename)

    # --- 6. グラフ出力(5章 + 改良点3,4,5) ---
    res.pdf_path, res.png_paths = _draw_graphs(
        res, day_order, day_dates, yaxis_overrides, ytick_overrides,
        show_all_weekdays, outdir, basename)
    return res


# ---- CSV(指示書 4章) ------------------------------------------------------
def _write_csv(res, day_order, day_dates, outdir, basename):
    out_cols = ['区分', '日付', '曜日'] + RAW_COLUMNS + DERIVED_COLUMNS

    # 明細
    detail = res.used_df.copy()
    detail['区分'] = '明細'
    detail = detail[out_cols]

    # 平均(曜日ごと)
    avg_rows = []
    for w in day_order:
        row = {c: '' for c in out_cols}
        row['区分'] = '平均'; row['日付'] = day_dates[w]; row['曜日'] = WEEKDAYS_JP[w]
        for c in METRIC_COLUMNS:
            row[c] = res.daily_avg.loc[w, c]
        avg_rows.append(row)
    avg_df = pd.DataFrame(avg_rows, columns=out_cols)

    # total(最終行)
    trow = {c: '' for c in out_cols}
    trow['区分'] = 'total'
    # total の曜日欄に採用曜日を控えめに記録(可読性のため)
    trow['曜日'] = '+'.join(WEEKDAYS_JP[w] for w in res.total_weekdays) if res.total_weekdays else ''
    for c in METRIC_COLUMNS:
        trow[c] = res.total[c]
    total_df = pd.DataFrame([trow], columns=out_cols)

    full = pd.concat([detail, avg_df, total_df], ignore_index=True)
    path = os.path.join(outdir, f'{basename}_集計.csv')
    full.to_csv(path, index=False, encoding='utf-8-sig')   # BOM付きUTF-8
    return path


# ---- グラフ(指示書 5章 + 改良) --------------------------------------------
def _draw_graphs(res, day_order, day_dates, yaxis_overrides, ytick_overrides,
                 show_all_weekdays, outdir, basename):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from matplotlib.ticker import MultipleLocator
    import matplotlib.font_manager as fm

    # 日本語フォント
    for cand in ['Hiragino Sans', 'Hiragino Kaku Gothic Pro', 'Hiragino Maru Gothic Pro',
                 'Noto Sans CJK JP', 'IPAexGothic', 'TakaoPGothic', 'Yu Gothic',
                 'YuGothic', 'Meiryo', 'MS Gothic']:
        if any(cand in f.name for f in fm.fontManager.ttflist):
            plt.rcParams['font.family'] = cand
            break
    plt.rcParams['axes.unicode_minus'] = False

    # 表示する曜日: 全曜日(月〜日) or データのある曜日のみ
    graph_days = list(range(7)) if show_all_weekdays else list(day_order)
    wd_to_x = {w: i for i, w in enumerate(graph_days)}      # 曜日index -> x位置
    x_labels = [WEEKDAYS_JP[w] for w in graph_days] + ['total']
    x_pos = list(range(len(graph_days) + 1))
    total_idx = len(graph_days)  # totalのx位置

    players = sorted(res.used_df['Name'].unique())
    cmap = plt.get_cmap('tab20')
    color_of = {p: cmap(i % 20) for i, p in enumerate(players)}

    def _isnan(v):
        try:
            return v is None or np.isnan(v)
        except TypeError:
            return v is None

    def _fmt(v):
        """棒ラベル用の数値整形。"""
        if _isnan(v):
            return ''
        av = abs(v)
        if av >= 1000:
            return f'{v:,.0f}'
        if av >= 100:
            return f'{v:.0f}'
        if av >= 10:
            return f'{v:.1f}'
        return f'{v:.2f}'

    png_paths = []
    pdf_path = os.path.join(outdir, f'{basename}_グラフ.pdf')
    pdf = PdfPages(pdf_path)

    for metric in METRIC_COLUMNS:
        fig, ax = plt.subplots(figsize=(10, 5.6))

        # --- 棒: 曜日別平均(未投入日は0) + total ---
        bar_vals = []
        for w in graph_days:
            if res.daily_avg is not None and w in res.daily_avg.index:
                v = res.daily_avg.loc[w, metric]
                bar_vals.append(0.0 if pd.isna(v) else float(v))
            else:
                bar_vals.append(0.0)   # データ未投入の曜日は負荷量0(改良点4)
        tv = res.total[metric] if (res.total is not None and metric in res.total.index) else np.nan
        bar_vals.append(tv)

        bar_heights = [0.0 if _isnan(b) else b for b in bar_vals]
        ax.bar(x_pos, bar_heights, color='#9DBBD6', alpha=0.55, width=0.6,
               zorder=1, label='平均 / total')

        # --- 棒の数値ラベル(棒グラフのみ。改良点3) ---
        for xi, v in zip(x_pos, bar_vals):
            label = _fmt(v)
            if label == '':
                continue
            ax.annotate(label, (xi, v if not _isnan(v) else 0.0),
                        textcoords='offset points', xytext=(0, 3),
                        ha='center', va='bottom', fontsize=8,
                        color='#33526e', zorder=4)

        # --- 個人: 点+線(区間ごとに実線/点線を切替) ---
        for p in players:
            prow = res.used_df[res.used_df['Name'] == p]
            val_at = {}   # x位置 -> 値
            for _, r in prow.iterrows():
                wd = int(r['_wd'])
                if wd in wd_to_x:
                    val_at[wd_to_x[wd]] = r[metric]
            present = sorted(val_at.keys())
            if not present:
                continue
            col = color_of[p]
            # 点
            ax.scatter(present, [val_at[i] for i in present],
                       color=col, s=22, zorder=3)
            # 線(隣接する表示位置=実線, 飛んでいる=薄い点線)
            for a, bx in zip(present[:-1], present[1:]):
                style = '-' if (bx - a) == 1 else ':'
                alpha = 0.95 if style == '-' else 0.35
                ax.plot([a, bx], [val_at[a], val_at[bx]], style,
                        color=col, alpha=alpha, linewidth=1.6, zorder=2)
            # 凡例用ダミー(点と同色)
            ax.plot([], [], '-', color=col, label=p)

        ax.set_title(metric, fontsize=14, fontweight='bold')
        ax.set_xticks(x_pos); ax.set_xticklabels(x_labels)
        ax.set_ylabel('実数値')
        ax.set_xlim(-0.6, total_idx + 0.6)
        ax.axvline(total_idx - 0.5, color='#bbbbbb', linestyle='--', linewidth=0.8)

        # 縦軸 下限・上限(任意)
        has_ylim = bool(yaxis_overrides.get(metric))
        if has_ylim:
            lo, hi = yaxis_overrides[metric]
            ax.set_ylim(lo, hi)
        else:
            # ラベルが上枠で切れないよう少し余白を確保
            ymin, ymax = ax.get_ylim()
            if ymax > 0:
                ax.set_ylim(ymin, ymax * 1.12)

        # 縦軸 目盛り幅(任意。改良点5)
        step = ytick_overrides.get(metric)
        if step:
            try:
                step = float(step)
                if step > 0:
                    ax.yaxis.set_major_locator(MultipleLocator(step))
            except (TypeError, ValueError):
                pass

        ax.legend(fontsize=7, ncol=2, loc='upper left', bbox_to_anchor=(1.01, 1.0),
                  borderaxespad=0, framealpha=0.9)
        ax.grid(axis='y', alpha=0.25)
        fig.tight_layout()

        png = os.path.join(outdir, f'{basename}_{metric.replace("/", "_").replace(" ", "_")}.png')
        fig.savefig(png, dpi=150, bbox_inches='tight')
        png_paths.append(png)
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

    pdf.close()
    return pdf_path, png_paths


if __name__ == '__main__':
    import glob, sys
    files = sys.argv[1:] or glob.glob('/mnt/user-data/uploads/学習院大学_*')
    r = process(files, outdir='/home/claude/test_out')
    print("ERRORS  :", r.errors)
    print("WARNINGS:", r.warnings)
    print("曜日     :", [WEEKDAYS_JP[w] for w in r.day_order])
    print("採用明細 :", None if r.used_df is None else len(r.used_df), "行")
    print("total曜日:", [WEEKDAYS_JP[w] for w in r.total_weekdays])
    print("total対象:", r.total_players)
    print("CSV      :", r.csv_path)
    print("PDF      :", r.pdf_path)
    print("PNG      :", len(r.png_paths), "枚")
