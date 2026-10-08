# 1次データ生成ツール

## 結論：実データをどこまで使えるか

| データ | 実データ化 | 理由 |
|---|---|---|
| レジの売上明細（注文・明細・会計） | **不可** | 飲食店の取引明細を公開している API は確認できていない。実店舗の POS データは各社の非公開データで、入手には契約が必要 |
| 天気（日別降水量） | 可能 | Open-Meteo Historical Weather API（無料・登録不要・CC BY 4.0）。気象庁の過去データは公式 API がなく、CSV の手動ダウンロードのみ |
| 祝日 | 可能 | 内閣府「国民の祝日」CSV、または holidays-jp API |

そのため「取引明細は架空、来客数を左右する外部要因（天気・祝日）は実データ」とする。雨の日に伝票数が減る等の関係は生成上の仮定であり、実測ではない。

参考（未採用）：Kaggle「Recruit Restaurant Visitor Forecasting」は実在する日本の飲食店の日別来客数を含むが、競技用データで利用条件の確認が必要なうえ、明細（何が売れたか）は含まない。

## 手順

インターネットに接続できる PC で実行する（Python 3.9 以上、`pip install openpyxl`）。

```bash
cd tools/datagen
python fetch_external.py   # 天気・祝日を data/external/ に保存（実データ）
python generate.py         # data/source/ と data/answer/ を再生成
```

- `fetch_external.py` を実行しない場合、天気は乱数の代替データで生成される（`data/answer/README.md` に `weather: synthetic` と出る）
- 現在コミットされている生成物は代替データ版。実データ版にするには上記2コマンドを実行してコミットし直す
- 同じシード・同じ外部データなら出力は毎回同一。研修生ごとにデータを変える場合は `--seed` を変える
- 天気データを使う場合、研修資料に「天気データ：Open-Meteo（CC BY 4.0）」と出典を記載すること

## 出力

| パス | 内容 |
|---|---|
| `data/source/pos_sales_YYYYMMDD.csv` | レジ売上 CSV（49 営業日分） |
| `data/source/pos_sales_20260904 (1).csv` | 不備 D-03：同日ファイルの複製 |
| `data/source/menu.xlsx` / `staff.xlsx` / `seats.txt` | マスタ系の 1次データ |
| `data/answer/README.md` | 不備 D-01〜D-07 の該当箇所、設計上の論点、生成の仮定 |
| `data/answer/R0x_*.csv` | 要件 R-01〜R-06 の正解値 |

## 変更する場合

期間・値上げ日・メニュー・価格・スタッフ・不備の内容は `generate.py` 冒頭の定数で変える。変更したら `docs/data-flow.md` 2章・4章も合わせて直すこと。
