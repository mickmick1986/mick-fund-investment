#!/usr/bin/env python3
"""Fetch current fund estimate changes and publish a same-origin JSON snapshot for GitHub Pages."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

ROOT = Path(__file__).resolve().parent
SOURCE_PATH = ROOT / "fund_data_enriched.json"
OUTPUT_PATH = ROOT / "deploy" / "live_estimates.json"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36",
    "Referer": "https://finance.sina.com.cn/",
}


def fetch_sse_index(session: requests.Session) -> dict | None:
    response = session.get(
        "https://hq.sinajs.cn/list=s_sh000001,sz399001,bj899050", headers=HEADERS, timeout=12
    )
    response.raise_for_status()
    text = response.content.decode("gbk", errors="replace")
    quotes = {}
    for line in text.splitlines():
        if '="' not in line or '"' not in line:
            continue
        key = line.split("hq_str_")[1].split("=")[0]
        quotes[key] = line.split('"')[1].split(",")

    shanghai = quotes.get("s_sh000001", [])
    shenzhen = quotes.get("sz399001", [])
    beijing = quotes.get("bj899050", [])
    if len(shanghai) < 6 or len(shenzhen) < 10:
        return None

    close = float(shanghai[1])
    change = float(shanghai[2])
    change_pct = float(shanghai[3])
    # 沪市摘要行情第6字段为成交额（万元）；深/北指数第10字段为成交额（元）。
    shanghai_turnover_yuan = float(shanghai[5]) * 10_000
    shenzhen_turnover_yuan = float(shenzhen[9])
    beijing_turnover_yuan = float(beijing[9]) if len(beijing) >= 10 else 0
    market_turnover_trillion = (
        shanghai_turnover_yuan + shenzhen_turnover_yuan + beijing_turnover_yuan
    ) / 1_000_000_000_000
    if close <= 0:
        return None
    return {
        "close": round(close, 2),
        "change": round(change, 2),
        "change_pct": round(change_pct, 4),
        "market_turnover_trillion": round(market_turnover_trillion, 2),
    }


def fetch_one(session: requests.Session, code: str) -> dict | None:
    response = session.get(
        f"https://hq.sinajs.cn/list=fu_{code}", headers=HEADERS, timeout=12
    )
    response.raise_for_status()
    text = response.content.decode("gbk", errors="replace")
    if f"hq_str_fu_{code}" not in text or '"' not in text:
        return None

    fields = text.split('"')[1].split(",")
    if len(fields) < 8:
        return None
    estimate_nav = float(fields[2])
    confirmed_nav = float(fields[3])
    if estimate_nav <= 0 or confirmed_nav <= 0:
        return None

    quote_date = fields[7].strip()
    quote_time = fields[1].strip()
    if not quote_date or not quote_time:
        return None
    return {
        "gszzl": round((estimate_nav - confirmed_nav) / confirmed_nav * 100, 4),
        "gztime": f"{quote_date} {quote_time}",
        "quote_date": quote_date,
    }


def build_live_fund_inputs(data: dict) -> dict[str, dict]:
    """Publish only the deterministic inputs required for browser-side intraday recomputation."""
    funds: dict[str, dict] = {}
    for item in data.get("funds", []):
        code = str(item.get("code", "")).zfill(6)
        if not code or not item.get("latest_nav"):
            continue
        funds[code] = {
            "confirmed_navs_desc": item.get("confirmed_navs_desc", []),
            "val_signal": item.get("val_signal", ""),
            "trend_20d_pct": item.get("trend_20d_pct"),
            "tp_display": item.get("tp_display"),
        }
    return funds


def main() -> int:
    data = json.loads(SOURCE_PATH.read_text(encoding="utf-8"))
    codes = [str(item["code"]).zfill(6) for item in data.get("funds", [])]
    estimates: dict[str, dict] = {}
    failures: list[str] = []

    with requests.Session() as session:
        for code in codes:
            try:
                item = fetch_one(session, code)
                if item:
                    estimates[code] = item
                else:
                    failures.append(code)
            except (requests.RequestException, ValueError, IndexError) as error:
                print(f"{code}: {error}")
                failures.append(code)

    try:
        with requests.Session() as session:
            sse_index = fetch_sse_index(session)
    except (requests.RequestException, ValueError, IndexError) as error:
        print(f"SSE index: {error}")
        sse_index = None

    now = datetime.now(ZoneInfo("Asia/Shanghai"))
    today = now.strftime("%Y-%m-%d")
    stale_quotes = [
        code for code, quote in estimates.items()
        if quote.get("quote_date") != today
    ]
    if stale_quotes:
        print("Stale quotes:", ", ".join(stale_quotes))
        return 3

    regime = data.get("market_regime", {})
    regime_params = data.get("market_regime_params") or {
        "rsi_tp_strong": 999 if regime.get("regime_key") == "correction" else 75,
        "rsi_tp_val": 999 if regime.get("regime_key") == "correction" else 70,
    }
    snapshot = {
        "trade_date": today,
        "updated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "source": "Sina fund estimate",
        "estimate_count": len(estimates),
        "estimates": estimates,
        "funds": build_live_fund_inputs(data),
        "market_regime": {
            "regime_key": regime.get("regime_key"),
            "params": regime_params,
        },
        "sse_index": sse_index,
    }
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    print(f"Published {len(estimates)}/{len(codes)} estimates to {OUTPUT_PATH}")
    if failures:
        print("Failed:", ", ".join(failures))
    return 0 if estimates else 2


if __name__ == "__main__":
    raise SystemExit(main())
