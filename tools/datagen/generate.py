"""みやこ食堂の1次データ（ダミー）と正解データを生成する。

  python generate.py            # data/source/ と data/answer/ を再生成
  python generate.py --seed 7   # 乱数を変える

天気（降水量）は乱数で作る架空の値で、来客数を増減させる隠れた要因として使う。
研修生に渡す 1次データには含めない。出力は seed が同じなら毎回同一になる。
"""
import argparse
import csv
import datetime as dt
import pathlib
import random
import shutil
from collections import defaultdict

from openpyxl import Workbook

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = ROOT / "data" / "source"
ANS = ROOT / "data" / "answer"

START = dt.date(2026, 8, 1)
END = dt.date(2026, 9, 30)
PRICE_UP = dt.date(2026, 9, 1)
CLOSED = {dt.date(2026, 8, 13), dt.date(2026, 8, 14), dt.date(2026, 8, 15)}  # お盆の臨時休業
# 内閣府「国民の祝日」2026年8〜9月分
HOLIDAYS = {"2026-08-11": "山の日", "2026-09-21": "敬老の日",
                    "2026-09-22": "国民の休日", "2026-09-23": "秋分の日"}

# (コード, 商品名, カテゴリ, 旧価格, 新価格, 種別, 販売開始, 販売終了)
# 種別: main=1人1品の主食 / side=大盛など / dish=一品 / drink
MENU = [
    ("A001", "からあげ定食",     "定食",     900,  950, "main", None, None),
    ("A002", "生姜焼き定食",     "定食",     950, 1000, "main", None, None),
    ("A003", "さば味噌定食",     "定食",     950, 1000, "main", None, None),
    ("A004", "ハンバーグ定食",   "定食",    1000, 1080, "main", None, None),
    ("A005", "日替わり定食",     "定食",     850,  900, "main", None, None),
    ("B001", "親子丼",           "丼",       800,  850, "main", None, None),
    ("B002", "カツ丼",           "丼",       900,  950, "main", None, None),
    ("C001", "ざるそば",         "麺",       700,  750, "main", None, None),
    ("C002", "冷やし中華",       "麺",       800,  800, "main", None, dt.date(2026, 8, 31)),
    ("D001", "ご飯大盛",         "サイド",   100,  100, "side", None, None),
    ("E001", "冷奴",             "一品",     300,  300, "dish", None, None),
    ("E002", "唐揚げ（単品）",   "一品",     450,  500, "dish", None, None),
    ("E003", "ポテトサラダ",     "一品",     350,  380, "dish", None, None),
    ("F001", "生ビール",         "ドリンク", 550,  600, "drink", None, None),
    ("F002", "ハイボール",       "ドリンク", 450,  480, "drink", None, None),
    ("F003", "ウーロン茶",       "ドリンク", 250,  250, "drink", None, None),
]
# D-05: メニュー表に載っていない期間限定メニュー
LIMITED = ("L901", "さんま定食", "定食", 1100, 1100, "main", dt.date(2026, 9, 10), None)
# D-01: 店主がレジ側の値上げを忘れた商品（レジは旧価格のまま）
POS_PRICE_NOT_UPDATED = "A004"

MAIN_WEIGHT = {"A001": 22, "A002": 16, "A003": 12, "A004": 14, "A005": 18,
               "B001": 10, "B002": 9, "C001": 6, "C002": 8, "L901": 12}

STAFF = [  # (No, 氏名, 雇用区分, 入社日, レジ表記, 出勤の重み)
    ("S01", "宮古 誠",   "店主",       "2015/04/01", "宮古", 1),
    ("S02", "佐藤 健太", "正社員",     "2021/04/01", "佐藤", 4),
    ("S03", "佐藤 美咲", "アルバイト", "2025/06/15", "佐藤", 3),  # D-06: レジ表記が同じ
    ("S04", "鈴木 彩",   "アルバイト", "2024/10/01", "鈴木", 3),
    ("S05", "高橋 翔",   "アルバイト", "2026/04/01", "高橋", 2),
]
SEATS = [(f"T{i}", "テーブル", 4) for i in range(1, 5)] + \
        [(f"T{i}", "テーブル", 2) for i in range(5, 9)] + \
        [(f"C{i}", "カウンター", 1) for i in range(1, 7)]

POS_HEADER = ["伝票番号", "明細番号", "来店日時", "卓番", "客数", "担当者", "商品コード", "商品名",
              "カテゴリ", "単価", "数量", "小計", "取消区分", "支払方法", "伝票合計"]
DUPLICATED_DAY = dt.date(2026, 9, 4)  # D-03


def synthetic_weather(rng):
    out, d = {}, START
    while d <= END:
        out[d.isoformat()] = round(rng.expovariate(1 / 12), 1) if rng.random() < 0.3 else 0.0
        d += dt.timedelta(days=1)
    return out


def price(item, day, for_pos):
    code, _, _, old, new = item[:5]
    if day < PRICE_UP or (for_pos and code == POS_PRICE_NOT_UPDATED):
        return old
    return new


def available(item, day):
    start, end = item[6], item[7]
    return (start is None or day >= start) and (end is None or day <= end)


def slip_count(rng, day, session, rain_mm, is_holiday):
    wd = day.weekday()  # 0=月
    if session == "lunch":
        base = 32 if wd == 5 else 48
        if is_holiday:
            base *= 0.6  # 平日昼はオフィス客中心という設定
    else:
        base = {4: 34, 5: 30}.get(wd, 26)
    if rain_mm >= 20:
        base *= 0.7
    elif rain_mm >= 5:
        base *= 0.85
    if day >= PRICE_UP:
        base *= 0.95
    return max(5, round(rng.gauss(base, base * 0.12)))


def arrival(rng, day, session):
    if session == "lunch":
        m = rng.triangular(0, 210, 75)    # 11:00〜14:30、ピーク 12:15
        t0 = dt.datetime(day.year, day.month, day.day, 11)
    else:
        m = rng.triangular(0, 210, 120)   # 17:00〜20:30、ピーク 19:00
        t0 = dt.datetime(day.year, day.month, day.day, 17)
    return t0 + dt.timedelta(minutes=int(m))


def pick_seat(rng, guests):
    if guests == 1 and rng.random() < 0.7:
        pool = [s for s in SEATS if s[2] == 1]
    elif guests <= 2:
        pool = [s for s in SEATS if s[2] == 2]
    else:
        pool = [s for s in SEATS if s[2] == 4]
    return rng.choice(pool)[0]


def build_orders(rng, day, session, menu):
    """1伝票の明細を (商品, 数量) のリストで返す。追加注文は別行になる。"""
    mains = [m for m in menu if m[5] == "main"]
    by = {m[0]: m for m in menu}
    guests = rng.choices([1, 2, 3, 4], weights=[55, 30, 8, 7] if session == "lunch" else [20, 45, 15, 20])[0]
    lines = []
    picked = rng.choices(mains, weights=[MAIN_WEIGHT[m[0]] for m in mains],
                         k=guests if session == "lunch" else sum(rng.random() < 0.6 for _ in range(guests)) or 1)
    counts = defaultdict(int)
    for m in picked:
        counts[m[0]] += 1
    lines += [(by[c], q) for c, q in counts.items()]
    if session == "lunch":
        big = sum(rng.random() < 0.15 for _ in range(guests))
        if big:
            lines.append((by["D001"], big))
        if rng.random() < 0.05:
            lines.append((by["F003"], 1))
    else:
        drinks = rng.choices(["F001", "F002", "F003"], weights=[50, 30, 20], k=guests)
        for c in sorted(set(drinks)):
            lines.append((by[c], drinks.count(c)))
        for c in rng.sample(["E001", "E002", "E003"], k=rng.choice([0, 1, 1, 2, 3])):
            lines.append((by[c], 1))
        for _ in range(rng.choice([0, 0, 1, 2])):  # 追加注文は同じ商品でも別行
            lines.append((by[rng.choice(["F001", "F001", "F002"])], 1))
    return guests, lines


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", type=int, default=20261101)
    args = ap.parse_args()
    rng = random.Random(args.seed)

    weather = synthetic_weather(random.Random(args.seed + 1))

    for d in (SRC, ANS):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir(parents=True)

    menu_all = MENU + [LIMITED]
    staff_codes = [s[4] for s in STAFF]
    staff_w = [s[5] for s in STAFF]

    rows_all = []  # 正解計算用（重複ファイルを含まない真のデータ）
    defects = defaultdict(list)
    day = START
    while day <= END:
        if day.weekday() == 6 or day in CLOSED:
            day += dt.timedelta(days=1)
            continue
        iso = day.isoformat()
        menu_today = [m for m in menu_all if available(m, day)]
        slips = []
        for session in ("lunch", "dinner"):
            for _ in range(slip_count(rng, day, session, weather.get(iso, 0.0), iso in HOLIDAYS)):
                slips.append((arrival(rng, day, session), session))
        slips.sort()

        day_rows = []
        for no, (ts, session) in enumerate(slips, start=1):
            slip_id = f"{day:%Y%m%d}-{no:04d}"
            guests, lines = build_orders(rng, day, session, menu_today)
            staff = "" if rng.random() < 0.02 else rng.choices(staff_codes, weights=staff_w)[0]
            if staff == "":
                defects["D-04"].append(slip_id)
            pay = rng.choices(["現金", "クレジットカード", "QRコード決済"],
                              weights=[55, 15, 30] if session == "lunch" else [40, 35, 25])[0]
            seat = pick_seat(rng, guests)
            detail = []
            for item, qty in lines:
                p = price(item, day, for_pos=True)
                detail.append([item, p, qty, 0])
                if item[0] == POS_PRICE_NOT_UPDATED and day >= PRICE_UP:
                    defects["D-01"].append(slip_id)
                if item[0] == LIMITED[0]:
                    defects["D-05"].append(slip_id)
            if rng.random() < 0.01:  # D-02: 誤って入力した商品をレジで取り消した
                item = rng.choice(menu_today)
                p = price(item, day, for_pos=True)
                detail += [[item, p, 1, 0], [item, p, -1, 1]]
                defects["D-02"].append(slip_id)
            total = sum(p * q for _, p, q, _ in detail)
            for line_no, (item, p, q, cancel) in enumerate(detail, start=1):
                day_rows.append({
                    "伝票番号": slip_id, "明細番号": line_no, "来店日時": ts.strftime("%Y/%m/%d %H:%M"),
                    "卓番": seat, "客数": guests, "担当者": staff, "商品コード": item[0],
                    "商品名": item[1], "カテゴリ": item[2], "単価": p, "数量": q,
                    "小計": p * q, "取消区分": cancel, "支払方法": pay, "伝票合計": total,
                    "_session": session,
                })
        rows_all += day_rows
        write_pos(SRC / f"pos_sales_{day:%Y%m%d}.csv", day_rows)
        day += dt.timedelta(days=1)

    # D-03: 店主が同じ日の CSV を2回保存した
    shutil.copy(SRC / f"pos_sales_{DUPLICATED_DAY:%Y%m%d}.csv",
                SRC / f"pos_sales_{DUPLICATED_DAY:%Y%m%d} (1).csv")

    write_menu(SRC / "menu.xlsx")
    write_staff(SRC / "staff.xlsx")
    (SRC / "seats.txt").write_text("".join(f"{c},{k},{n}\r\n" for c, k, n in SEATS), encoding="utf-8")

    write_answers(rows_all, defects, weather, args.seed)
    print(f"days={len({r['伝票番号'][:8] for r in rows_all})} slips={len({r['伝票番号'] for r in rows_all})} "
          f"rows={len(rows_all)}")


def fmt_money(v):
    return f"{v:,}"


def write_pos(path, rows):
    # Windows のレジを想定し UTF-8(BOM付き)・CRLF。金額はカンマ付き（D-07）
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(POS_HEADER)
        for r in rows:
            w.writerow([fmt_money(r[h]) if h in ("小計", "伝票合計") else r[h] for h in POS_HEADER])


def write_menu(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "メニュー"
    ws.append(["商品コード", "商品名", "カテゴリ", "価格", "旧価格", "改定日", "販売状況"])
    for code, name, cat, old, new, _, _, end in MENU:  # L901 は載せない（D-05）
        changed = old != new
        ws.append([code, name, cat, new if end is None else old, old if changed and end is None else None,
                   PRICE_UP if changed and end is None else None, "終了" if end else "販売中"])
    for c in ws["F"][1:]:
        c.number_format = "yyyy/mm/dd"
    wb.save(path)


def write_staff(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "スタッフ"
    ws.append(["スタッフNo", "氏名", "雇用区分", "入社日"])
    for no, name, kind, joined, _, _ in STAFF:
        ws.append([no, name, kind, joined])
    wb.save(path)


def write_csv(path, header, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def write_answers(rows, defects, weather, seed):
    slips = {}
    for r in rows:
        slips.setdefault(r["伝票番号"], r)

    # R-01 日別売上（取消行を含めて合算＝取消は相殺される）
    daily = defaultdict(lambda: [0, 0, 0])
    for r in rows:
        daily[r["来店日時"][:10]][0] += r["小計"]
    for s in slips.values():
        daily[s["来店日時"][:10]][1] += 1
        daily[s["来店日時"][:10]][2] += s["客数"]
    write_csv(ANS / "R01_daily_sales.csv", ["日付", "売上", "伝票数", "客数", "降水量mm"],
              [[d, v[0], v[1], v[2], weather.get(d.replace("/", "-"), "")] for d, v in sorted(daily.items())])

    # R-02 メニュー別
    menu = defaultdict(lambda: [0, 0])
    names = {}
    for r in rows:
        menu[r["商品コード"]][0] += r["数量"]
        menu[r["商品コード"]][1] += r["小計"]
        names[r["商品コード"]] = r["商品名"]
    write_csv(ANS / "R02_menu_sales.csv", ["商品コード", "商品名", "数量", "売上"],
              [[c, names[c], v[0], v[1]] for c, v in sorted(menu.items(), key=lambda x: -x[1][1])])

    # R-03 値上げ前後（月別）
    month = defaultdict(lambda: [0, 0, 0, set()])
    for r in rows:
        month[r["来店日時"][:7]][0] += r["小計"]
    for s in slips.values():
        m = month[s["来店日時"][:7]]
        m[1] += 1
        m[2] += s["客数"]
        m[3].add(s["来店日時"][:10])
    write_csv(ANS / "R03_before_after.csv", ["年月", "売上", "伝票数", "客数", "営業日数", "1日平均売上", "客単価"],
              [[k, v[0], v[1], v[2], len(v[3]), round(v[0] / len(v[3])), round(v[0] / v[2])]
               for k, v in sorted(month.items())])

    # R-04 昼夜（11:00〜15:00 昼 / 17:00〜21:00 夜）
    sess = defaultdict(lambda: [0, 0])
    for r in rows:
        sess[r["_session"]][0] += r["小計"]
    for s in slips.values():
        sess[s["_session"]][1] += 1
    write_csv(ANS / "R04_lunch_dinner.csv", ["区分", "売上", "伝票数"],
              [["昼", *sess["lunch"]], ["夜", *sess["dinner"]]])

    # R-05 支払方法（伝票単位）
    pay = defaultdict(lambda: [0, 0])
    for s in slips.values():
        pay[s["支払方法"]][0] += 1
        pay[s["支払方法"]][1] += int(s["伝票合計"])
    total = sum(v[1] for v in pay.values())
    write_csv(ANS / "R05_payment.csv", ["支払方法", "伝票数", "金額", "金額構成比%"],
              [[k, v[0], v[1], round(v[1] * 100 / total, 1)] for k, v in sorted(pay.items(), key=lambda x: -x[1][1])])

    # R-06 担当者別（レジ表記単位。佐藤は2名を判別できない）
    st = defaultdict(int)
    for s in slips.values():
        st[s["担当者"] or "（空欄）"] += 1
    write_csv(ANS / "R06_staff.csv", ["担当者（レジ表記）", "伝票数", "備考"],
              [[k, v, "佐藤 健太・佐藤 美咲の合算。判別不能" if k == "佐藤" else ""]
               for k, v in sorted(st.items(), key=lambda x: -x[1])])

    sales_total = sum(r["小計"] for r in rows)
    lines = [
        "# 正解データ・不備一覧（講師用。研修生に渡さない）", "",
        f"- 生成シード: {seed}",
        "- 天気（降水量）: 乱数による架空の値。来客数の増減要因としてのみ使用し、1次データには含めない",
        f"- 祝日: {', '.join(f'{k} {v}' for k, v in sorted(HOLIDAYS.items()))}",
        f"- 期間売上合計（D-03 の重複ファイルを除外、取消は相殺）: {sales_total:,} 円",
        f"- 伝票数: {len(slips):,} / 明細行: {len(rows):,}", "",
        "## 不備（data-flow.md 4章）", "",
        "| # | 内容 | 該当 |", "|---|---|---|",
        f"| D-01 | ハンバーグ定食(A004) はメニュー表 1,080円だがレジは 1,000円のまま。売上はレジ単価（実際に受け取った金額）が正 | 9月の {len(set(defects['D-01']))} 伝票 |",
        f"| D-02 | 誤入力した商品の行（取消区分=0・数量1）と取消行（取消区分=1・数量-1）が両方残る。全行合算で相殺される。取消区分=0 だけで集計すると過大になる | {', '.join(defects['D-02'])} |",
        f"| D-03 | `pos_sales_{DUPLICATED_DAY:%Y%m%d} (1).csv` は同日ファイルの完全な複製。両方取り込むと二重計上 | 1 ファイル |",
        f"| D-04 | 担当者が空欄 | {len(defects['D-04'])} 伝票 |",
        f"| D-05 | さんま定食(L901) がメニュー表にない（9/10〜の期間限定） | {len(set(defects['D-05']))} 伝票 |",
        "| D-06 | 名簿に佐藤が2名（健太・美咲）、レジは姓のみ。データからは判別不能で、店主への確認事項として挙げられるかを見る | R06_staff.csv |",
        "| D-07 | 小計・伝票合計がカンマ付き文字列（例: \"1,900\"） | 全ファイル |", "",
        "## 不備ではないが設計上の論点", "",
        "- 夜の追加注文は同じ商品でも別の明細行になる。明細の主キーを（伝票番号, 商品コード）にすると重複エラーになる",
        "- 冷やし中華(C002) は 8/31 で販売終了。メニュー表に「終了」として残っている",
        "- 8/13〜8/15 は臨時休業、日曜は定休でファイルが存在しない。売上0の日を日別一覧に出すかは要件次第",
        "- 来店日時は日本時間の文字列。DB の GETDATE()（UTC）と比較すると9時間ずれる", "",
        "## 生成上の仮定（実データに基づかない）", "",
        "- 伝票数: 平日昼 48・土曜昼 32・祝日昼×0.6、夜 月〜木 26・金 34・土 30（±12%）",
        "- 降水量 5mm 以上で×0.85、20mm 以上で×0.7、値上げ後×0.95",
        "- 人数構成・注文内容・支払方法の比率はすべて作成者の仮定",
    ]
    (ANS / "README.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
