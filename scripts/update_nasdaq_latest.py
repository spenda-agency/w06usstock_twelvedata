"""Twelve Data から対象シート（gid=2055700762） A列の銘柄の最新値を取得し、B列に書き込む。

GitHub Actions から実行する想定。必要な環境変数:
  TWELVEDATA_API_KEY           Twelve Data の API キー
  GOOGLE_SERVICE_ACCOUNT_JSON  Google サービスアカウントの JSON キー（中身そのもの）
任意:
  SPREADSHEET_ID / SHEET_ID / CHUNK_SIZE
"""

import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from decimal import ROUND_HALF_UP, Decimal

import gspread
import requests

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1Q1xfafJ_lonuellGTdmGt0Mdjqas0NwXOOe0BJW0u9k")
# シート名は変わることがあるので、URL の gid（シートID）で特定する
SHEET_ID = int(os.environ.get("SHEET_ID", "2055700762"))
# 1分あたりのAPIクレジット上限（1銘柄=1クレジット）。Basic(無料)=8, Grow=55 など
CHUNK_SIZE = int(os.environ.get("CHUNK_SIZE", "8"))
WAIT_SEC = 61
JST = timezone(timedelta(hours=9))


def fetch_quotes(symbols, api_key):
    """/quote を複数銘柄まとめて呼び、{symbol: 最新値} を返す。

    quote の close は、取引終了後なら当日終値、取引中なら直近の最新値。
    """
    res = requests.get(
        "https://api.twelvedata.com/quote",
        params={"symbol": ",".join(symbols)},
        headers={"Authorization": f"apikey {api_key}"},  # URL・ログにキーを出さない
        timeout=60,
    )
    try:
        data = res.json()
    except ValueError:
        data = {}
    if res.status_code != 200:
        raise RuntimeError(f"Twelve Data HTTP {res.status_code}: {data.get('message', res.text[:300])}")

    if isinstance(data, dict) and data.get("status") == "error" and "code" in data:
        raise RuntimeError(f"Twelve Data error: {data.get('code')} {data.get('message')}")

    by_symbol = {symbols[0]: data} if len(symbols) == 1 else data
    out = {}
    for s in symbols:
        d = by_symbol.get(s) or {}
        close = d.get("close")
        if close not in (None, "") and d.get("status") != "error":
            # 整数に四捨五入（Python の round は偶数丸めなので Decimal で四捨五入）
            out[s] = int(Decimal(str(close)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        else:
            out[s] = f"ERR: {d.get('message', 'no data')}"
            print(f"[warn] {s}: {out[s]}", file=sys.stderr)
    return out


def main():
    api_key = os.environ["TWELVEDATA_API_KEY"].strip()
    if not api_key:
        sys.exit("TWELVEDATA_API_KEY が空です")
    print(f"API key length: {len(api_key)}")  # 通常は32文字
    creds = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])

    ws = gspread.service_account_from_dict(creds).open_by_key(SPREADSHEET_ID).get_worksheet_by_id(SHEET_ID)
    print(f"sheet: {ws.title}")
    codes = [c.strip() for c in ws.col_values(1)[1:]]  # 1行目は見出し
    symbols = [c for c in dict.fromkeys(codes) if c]
    print(f"{len(symbols)} symbols")

    results = {}
    for i in range(0, len(symbols), CHUNK_SIZE):
        if i:
            time.sleep(WAIT_SEC)  # 1分あたりのクレジット上限を回避
        chunk = symbols[i : i + CHUNK_SIZE]
        results.update(fetch_quotes(chunk, api_key))
        print(f"fetched {min(i + CHUNK_SIZE, len(symbols))}/{len(symbols)}")

    values = [[results.get(c, "") if c else ""] for c in codes]
    stamp = datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    ws.update(range_name="B1", values=[[f"最新値 {stamp} JST"]])
    if values:
        ws.update(range_name=f"B2:B{len(values) + 1}", values=values)
    print("done")


if __name__ == "__main__":
    main()
