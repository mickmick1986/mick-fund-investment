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
        "https://hq.sinajs.cn/list=s_sh000001", headers=HEADERS, timeout=12
    )
    response.raise_for_status()
    text = response.content.decode("gbk", errors="replace")
    if '"' not in text:
        return None
    fields = text.split('"')[1].split(",")
    if len(fields) < 5:
        return None
    close = float(fields[1])
    change = float(fields[2])
    change_pct = float(fields[3])
    if close <= 0:
        return None
    return {
        "close": round(close, 2),
        "change": round(change, 2),
        "change_pct": round(change_pct, 4),
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

    return {
        "gszzl": round((estimate_nav - confirmed_nav) / confirmed_nav * 100, 4),
        "gztime": f"{fields[7]} {fields[1]}",
    }


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
    snapshot = {
        "trade_date": now.strftime("%Y-%m-%d"),
        "updated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "source": "Sina fund estimate",
        "estimate_count": len(estimates),
        "estimates": estimates,
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
