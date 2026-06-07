# 1週間 負荷推移 可視化 自動化

開発指示書(第1版)に準拠した実装です。GPSデータ(Excel / Numbers)を曜日ごとに読み込み、
外れ値を除いた使用データから「表(CSV)」と「グラフ(PDF・PNG)」を自動生成します。

## 構成

| ファイル | 役割 |
|----------|------|
| `app.py` | ドラッグ&ドロップUI(Streamlit)。月〜日の7箱・設定・ダウンロード |
| `load_engine.py` | 処理エンジン本体(読み込み→抽出→外れ値除外→集計→CSV/グラフ生成) |
| `requirements.txt` | 依存ライブラリ |

## セットアップ

Python 3.10 以降を推奨。

```bash
pip install -r requirements.txt
```

## 使い方A：アプリ(推奨)

```bash
streamlit run app.py
```

ブラウザが開きます。

1. 月〜日の箱に、その日のデータファイルをドラッグして投入(対応: `.xlsx` `.xls` `.numbers` `.csv`)
2. 左の「設定」で外れ値しきい値(既定 ±30%)や、必要なら指標ごとの縦軸範囲を指定
3. 「▶ 集計・可視化を実行」を押す
4. 表(CSV)・グラフ(PDF / PNG)をダウンロード

> 日付はファイル名 `..._YYYYMMDD_...` から自動取得します。どの箱に入れても、
> 実際の曜日はファイル名の日付で判定します。同一曜日に2ファイルが来た場合はエラーになります。

## 使い方B：コマンドライン(UIなしで一括処理)

```bash
python load_engine.py ファイル1 ファイル2 ...
# 例: python load_engine.py data/学習院大学_20260505_*.numbers data/...xlsx
```

出力は `output/` に生成されます(`負荷推移_集計.csv` / `負荷推移_グラフ.pdf` / 各指標PNG)。
エンジンを自作スクリプトから呼ぶ場合:

```python
import load_engine as eng
res = eng.process(
    files=["a.numbers", "b.xlsx"],
    threshold=0.30,                       # ±30%
    yaxis_overrides={"Distance": (0, 12000)},  # 指標ごとの縦軸固定(任意)
    outdir="output",
)
print(res.csv_path, res.pdf_path, res.png_paths)
```

## 仕様の要点(指示書との対応)

- **対象行(2.4)**: Session を正規化(trim+小文字)し、`all` 完全一致 または `full` 前方一致のみ採用。`ROUND ALL` 等は除外。
- **追加列(3.1)**: `Accel_All=Accel_Z2+Accel_Z3` / `Decel_All=Decel_Z2+Decel_Z3` / `HighAgility=Accel_Z3+Decel_Z3`。
- **外れ値(3.2)**: 日ごとに採用行のDistance平均Aを1回算出 → `V=(P−A)/A` → `|V|>しきい値` を除外。
- **total(3.4)**: 全アップロード日に有効データが残る選手のみ対象。各選手の週合計の選手間平均。
- **CSV(4章)**: `区分,日付,曜日 + 生データ22列 + 追加3列`。明細→曜日別平均→total の順。UTF-8(BOM付き)。
- **グラフ(5章)**: 12指標。横軸=曜日+total、縦軸=実数値。平均/totalは棒、個人は点+線(連続区間は実線・欠損をまたぐ区間は薄い点線)。選手は色分け+凡例。PDF1枚 + PNG12枚。

## 注意・前提

- Numbers読み込みには `numbers-parser` を使用します(暗号化されたNumbersファイルは非対応)。
- 個人の折れ線は曜日位置にのみ描画し、total列は棒のみ(集計値)です。
- 日本語フォントは環境にあるもの(Hiragino / Yu Gothic / Noto など)を自動選択します。
