# 1次データ生成ツール

みやこ食堂の 1次データ（ダミー）と講師用の正解データを生成する。**すべて架空のデータ**で、実データ（外部 API 等）は使わない。

## 手順

Python 3.9 以上、`pip install openpyxl`。

```bash
cd tools/datagen
python generate.py            # data/source/ と data/answer/ を再生成
python generate.py --seed 7   # 乱数を変える
```

- 同じシードなら出力は毎回同一
- 天気（降水量）は乱数で作り、「雨の日は来客が減る」隠れた要因としてのみ使う。1次データには含めない
- 祝日は 2026年8〜9月分をスクリプト内に定義（山の日・敬老の日・国民の休日・秋分の日）

## 出力

| パス | 内容 | 研修生に渡すか |
|---|---|---|
| `data/source/pos_sales_YYYYMMDD.csv` | レジ売上 CSV（49 営業日分） | 渡す |
| `data/source/pos_sales_20260904 (1).csv` | 不備 D-03：同日ファイルの複製 | 渡す |
| `data/source/menu.xlsx` / `staff.xlsx` / `seats.txt` | マスタ系の 1次データ | 渡す |
| `data/answer/README.md` | 不備 D-01〜D-07 の該当箇所、設計上の論点、生成の仮定 | 渡さない |
| `data/answer/R0x_*.csv` | 要件 R-01〜R-06 の正解値 | 渡さない |

## 変更する場合

期間・値上げ日・メニュー・価格・スタッフ・不備の内容は `generate.py` 冒頭の定数で変える。変更したら `docs/data-flow.md` 2章・4章も合わせて直すこと。
