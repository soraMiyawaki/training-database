"""外部APIから実データ（天気・祝日）を取得し data/external/ に保存する。

社内PCなどインターネットに接続できる環境で実行する。標準ライブラリのみ使用。
  python fetch_external.py

取得元:
  - 天気: Open-Meteo Historical Weather API（CC BY 4.0。出典表記が必要）
  - 祝日: holidays-jp（内閣府「国民の祝日」CSV を元にした非公式API）
"""
import json
import pathlib
import urllib.request

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "external"

START, END = "2026-08-01", "2026-09-30"
# 店舗所在地の仮定：東京駅周辺
LAT, LON = 35.6812, 139.7671

WEATHER_URL = (
    "https://archive-api.open-meteo.com/v1/archive"
    f"?latitude={LAT}&longitude={LON}&start_date={START}&end_date={END}"
    "&daily=precipitation_sum,temperature_2m_max&timezone=Asia%2FTokyo"
)
HOLIDAY_URL = "https://holidays-jp.github.io/api/v1/2026/date.json"


def fetch(url):
    with urllib.request.urlopen(url, timeout=30) as res:
        return json.load(res)


def main():
    OUT.mkdir(parents=True, exist_ok=True)

    w = fetch(WEATHER_URL)["daily"]
    weather = {
        d: {"precipitation_mm": p, "temp_max_c": t}
        for d, p, t in zip(w["time"], w["precipitation_sum"], w["temperature_2m_max"])
    }
    (OUT / "weather.json").write_text(
        json.dumps({"source": "Open-Meteo (CC BY 4.0)", "lat": LAT, "lon": LON, "daily": weather},
                   ensure_ascii=False, indent=1), encoding="utf-8")

    holidays = fetch(HOLIDAY_URL)
    (OUT / "holidays.json").write_text(
        json.dumps({"source": "holidays-jp", "dates": holidays}, ensure_ascii=False, indent=1),
        encoding="utf-8")

    print(f"saved: {OUT / 'weather.json'}, {OUT / 'holidays.json'}")


if __name__ == "__main__":
    main()
