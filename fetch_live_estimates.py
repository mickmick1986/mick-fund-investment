#!/usr/bin/env python3
"""Fetch current fund estimate changes and publish a same-origin JSON snapshot for GitHub Pages."""
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

ROOT = Path(__file__).resolve().parent
SOURCE_PATH = ROOT / "fund_data_enriched.json"
OUTPUT_PATH = ROOT / "deploy" / "live_estimates.json"
CONFIG_PATH = ROOT / "config" / "live_index_map.json"

# 取值模式：legacy = 仅源端估值（历史行为，逐字节等价）；auto = 三级兜底（S1 → B → C → 暂无）。
# 用环境变量控制是为了可秒级回退：任何一帧异常都能把 LIVE_EST_SOURCE_MODE 改回 legacy。
SOURCE_MODE = os.environ.get("LIVE_EST_SOURCE_MODE", "auto").strip().lower()
PLACEHOLDER_TOL = 1e-9
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36",
    "Referer": "https://finance.sina.com.cn/",
}
STATUS_DIR = ROOT / ".workbuddy"
STATUS_PATH = STATUS_DIR / "live_estimates_status.json"
STATUS_HISTORY_PATH = STATUS_DIR / "live_estimates_status.jsonl"


def write_status(stage: str, outcome: str, **details: object) -> None:
    """Keep local machine-readable diagnostics outside the published data path."""
    payload = {
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "component": "fetch_live_estimates",
        "stage": stage,
        "outcome": outcome,
        **details,
    }
    STATUS_DIR.mkdir(parents=True, exist_ok=True)
    STATUS_PATH.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    with STATUS_HISTORY_PATH.open("a", encoding="utf-8") as history:
        history.write(json.dumps(payload, ensure_ascii=False) + "\n")


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


def fetch_one_fields(session: requests.Session, code: str) -> list[str] | None:
    """Return the raw Sina fund quote fields, or None when the payload is unusable."""
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
    return fields


def parse_confirmed_quote(fields: list[str] | None) -> dict | None:
    """Legacy semantic parse: both NAVs must be positive and the quote date must exist."""
    if not fields or len(fields) < 8:
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
        "estimate_nav": estimate_nav,
        "confirmed_nav": confirmed_nav,
        "quote_date": quote_date,
        "quote_time": quote_time,
    }


def fetch_one(session: requests.Session, code: str) -> dict | None:
    """Legacy path: 源端估算净值 / 确认净值 直接算当日涨跌幅。"""
    parsed = parse_confirmed_quote(fetch_one_fields(session, code))
    if not parsed:
        return None
    return {
        "gszzl": round(
            (parsed["estimate_nav"] - parsed["confirmed_nav"])
            / parsed["confirmed_nav"]
            * 100,
            4,
        ),
        "gztime": f"{parsed['quote_date']} {parsed['quote_time']}",
        "quote_date": parsed["quote_date"],
    }


def load_live_index_map() -> dict:
    """Load the shared fund→index map. A missing or broken file degrades to B/暂无 only."""
    fallback = {"funds": {}, "beta_default": 0.95, "max_abs_index_change_pct": 15.0}
    try:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        print(f"Live index map unavailable: {CONFIG_PATH}")
        return fallback
    if not isinstance(config, dict):
        return fallback
    config.setdefault("funds", {})
    config.setdefault("beta_default", 0.95)
    config.setdefault("max_abs_index_change_pct", 15.0)
    if not isinstance(config["funds"], dict):
        config["funds"] = {}
    return config


def fetch_index_quotes(session: requests.Session, symbols: list[str]) -> dict[str, dict]:
    """Batch fetch reference index/ETF quotes (full sina format: f[2]=昨收, f[3]=现价)."""
    if not symbols:
        return {}
    quotes: dict[str, dict] = {}
    chunk_size = 20
    for start in range(0, len(symbols), chunk_size):
        chunk = symbols[start:start + chunk_size]
        try:
            response = session.get(
                "https://hq.sinajs.cn/list=" + ",".join(chunk), headers=HEADERS, timeout=12
            )
            response.raise_for_status()
        except requests.RequestException as error:
            print(f"Index batch failed: {error}")
            continue
        text = response.content.decode("gbk", errors="replace")
        for line in text.splitlines():
            if '="' not in line:
                continue
            try:
                symbol = line.split("hq_str_")[1].split("=")[0]
                payload = line.split('"')[1]
            except IndexError:
                continue
            if not payload.strip():
                continue
            fields = payload.split(",")
            if len(fields) < 4:
                continue
            try:
                prev_close = float(fields[2])
                last_price = float(fields[3])
            except ValueError:
                continue
            if prev_close <= 0 or last_price <= 0:
                continue
            quotes[symbol] = {
                "name": fields[0].strip(),
                "change_pct": (last_price - prev_close) / prev_close * 100,
            }
    return quotes


def resolve_estimate(
    code: str, fields: list[str] | None, index_quotes: dict[str, dict], config: dict
) -> dict | None:
    """Three-level resolution: S1 源端估值 → B 源端估算涨幅 → C 指数×β → None（估值暂缺）。

    Returns None only when the underlying fund quote itself is unusable, so that a
    single broken fund still blocks the frame exactly as before (fail-closed).
    """
    parsed = parse_confirmed_quote(fields)
    if not parsed:
        return None

    base = {
        "gztime": f"{parsed['quote_date']} {parsed['quote_time']}",
        "quote_date": parsed["quote_date"],
    }
    diagnostics = {"s1": None, "b": None, "idx": None, "beta": None, "note": None}

    # S1：源端估算净值（此处只记录，是否采信交给后面的三级仲裁统一判定）
    nav_gap_limit = max(PLACEHOLDER_TOL, parsed["confirmed_nav"] * 1e-9)
    s1_value = None
    if abs(parsed["estimate_nav"] - parsed["confirmed_nav"]) > nav_gap_limit:
        s1_value = (
            (parsed["estimate_nav"] - parsed["confirmed_nav"])
            / parsed["confirmed_nav"]
            * 100
        )
        diagnostics["s1"] = round(s1_value, 4)

    # B：新浪自带的估算涨跌幅（估算净值冻结时仍会更新，但本身也可能滞后）
    b_value = None
    if len(fields) >= 10:
        try:
            candidate = float(fields[9])
        except (TypeError, ValueError):
            candidate = None
        if candidate is not None and abs(candidate) > 1e-4:
            b_value = candidate
    diagnostics["b"] = round(b_value, 4) if b_value is not None else None

    # C：指数/ETF 代理 × β（仅被动基金；必须通过名称校验与幅度校验）
    c_value = None
    index_note = None
    entry = config["funds"].get(code) or {}
    quote_symbol = entry.get("quote")
    if entry.get("passive") and quote_symbol:
        quote = index_quotes.get(quote_symbol)
        if quote:
            name = quote.get("name") or ""
            keywords = [k for k in (entry.get("keywords") or []) if k]
            if not any(keyword in name for keyword in keywords):
                index_note = f"名称校验失败({quote_symbol}={name})"
            elif abs(quote["change_pct"]) > config["max_abs_index_change_pct"]:
                index_note = f"指数涨跌幅异常({quote['change_pct']:.2f}%)"
            else:
                beta = entry.get("beta") or config["beta_default"]
                c_value = quote["change_pct"] * beta
                diagnostics["idx"] = round(quote["change_pct"], 4)
                diagnostics["beta"] = beta
                index_note = f"{quote_symbol}={name}"
        else:
            index_note = f"指数行情缺失({quote_symbol})"
    elif not entry.get("passive"):
        index_note = "主动基金不启用指数兜底"
    else:
        index_note = "无可用跟踪指数"

    tolerance = config.get("b_vs_index_tolerance_pct", 1.0)
    placeholder_abs = config.get("s1_placeholder_abs_pct", 0.05)
    placeholder_peer = config.get("s1_placeholder_min_peer_pct", 0.5)
    flat_peer = config.get("c_flat_peer_pct", 0.3)

    # S1 占位识别：源端有时把「估算净值」填成只差一个最小变动单位的近零值（如 +0.01%），
    # 而标的指数实际波动百分之几。近零 + 同源/指数明显非零 → 判为占位，不可采信。
    if s1_value is not None:
        peer = max(abs(b_value or 0.0), abs(c_value or 0.0))
        if abs(s1_value) <= placeholder_abs and peer >= placeholder_peer:
            diagnostics["note"] = (
                f"S1={s1_value:+.2f}% 近零，但同源估算/指数为 {peer:.2f}%，判定源端占位"
            )
            s1_value = None

    if s1_value is not None:
        return {**base, "gszzl": round(s1_value, 4), "src": "s1", "_diag": diagnostics}

    prefix = f"{diagnostics['note']}；" if diagnostics["note"] else ""

    # C 为主、B 兜底：指数是客观行情，B 可能滞后；但代理标的无成交（C 恰好为 0）时退回 B。
    c_is_flat_proxy = (
        c_value is not None
        and abs(c_value) < 1e-9
        and b_value is not None
        and abs(b_value) > flat_peer
    )
    if c_value is not None and not c_is_flat_proxy:
        note = index_note
        if b_value is not None and abs(b_value - c_value) > tolerance:
            note = (
                f"{index_note}；源端估算涨幅 {b_value:+.2f}% 与指数推算 {c_value:+.2f}% "
                f"偏差超过 {tolerance}pp，以指数为锚"
            )
        diagnostics["note"] = prefix + note
        return {**base, "gszzl": round(c_value, 4), "src": "idx", "_diag": diagnostics}

    if b_value is not None:
        diagnostics["note"] = prefix + f"无可用指数推算，采用源端估算涨幅（{index_note}）"
        return {**base, "gszzl": round(b_value, 4), "src": "b", "_diag": diagnostics}

    if c_value is not None:
        diagnostics["note"] = prefix + index_note
        return {**base, "gszzl": round(c_value, 4), "src": "idx", "_diag": diagnostics}

    # A：保底 —— 保留条目但明确为空，页面显示「估值暂缺」，绝不伪造 0.00%
    diagnostics["note"] = prefix + index_note
    return {**base, "gszzl": None, "src": "none", "_diag": diagnostics}


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


def build_market_status(data: dict) -> dict:
    """Expose real availability dates without presenting stale values as intraday data.

    The live snapshot may carry the latest confirmed market indicators as metadata,
    but it must explicitly preserve each indicator's own date and availability.
    Northbound intraday net flow is intentionally null because an older confirmed
    value must not be relabeled as today's official intraday figure.
    """
    regime = data.get("market_regime") or {}
    fear = regime.get("fear_greed") or {}
    hk = regime.get("hk_connect") or {}
    margin = regime.get("margin_balance") or {}

    south_value = hk.get("south_net")
    south_date = hk.get("south_date") or hk.get("date")
    margin_balance = margin.get("balance")
    margin_balance_trillion = margin_balance / 1e4 if margin_balance is not None else None

    return {
        "fear_greed": {
            "value": fear.get("value"),
            "label": fear.get("label"),
            "trade_date": fear.get("date"),
            "availability": "latest_available" if fear.get("value") is not None else "unavailable",
        },
        "north_flow": {
            "net_flow": None,
            "trade_date": None,
            "availability": "official_intraday_unavailable",
        },
        "south_flow": {
            "net_flow": south_value,
            "trade_date": south_date,
            "availability": "latest_available" if south_value is not None else "unavailable",
        },
        "margin_balance": {
            "balance_trillion": round(margin_balance_trillion, 4) if margin_balance_trillion is not None else None,
            "net_buy": margin.get("net_buy"),
            "trade_date": margin.get("date"),
            "availability": "latest_available_t_plus_1" if margin_balance is not None else "unavailable",
        },
    }


def main() -> int:
    write_status("start", "started", source=str(SOURCE_PATH), output=str(OUTPUT_PATH))
    try:
        data = json.loads(SOURCE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError) as error:
        write_status("source", "failed", error=str(error))
        print(f"Source data failed: {error}")
        return 2

    codes = [str(item["code"]).zfill(6) for item in data.get("funds", [])]
    estimates: dict[str, dict] = {}
    failures: list[str] = []
    diagnostics: dict[str, dict] = {}
    config = (
        load_live_index_map()
        if SOURCE_MODE == "auto"
        else {"funds": {}, "beta_default": 0.95, "max_abs_index_change_pct": 15.0}
    )

    with requests.Session() as session:
        index_quotes: dict[str, dict] = {}
        if SOURCE_MODE == "auto":
            symbols = sorted(
                {
                    entry.get("quote")
                    for entry in config["funds"].values()
                    if isinstance(entry, dict) and entry.get("quote")
                }
            )
            index_quotes = fetch_index_quotes(session, symbols)

        for code in codes:
            try:
                if SOURCE_MODE == "auto":
                    item = resolve_estimate(
                        code, fetch_one_fields(session, code), index_quotes, config
                    )
                else:
                    item = fetch_one(session, code)
                if not item:
                    failures.append(code)
                    continue
                detail = item.pop("_diag", None)
                estimates[code] = item
                if detail is not None:
                    diagnostics[code] = {**detail, "src": item.get("src")}
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
        write_status(
            "validate",
            "failed",
            reason="stale_quotes",
            expected_trade_date=today,
            estimate_count=len(estimates),
            expected_count=len(codes),
            stale_count=len(stale_quotes),
            stale_codes=stale_quotes,
        )
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
        "market_status": build_market_status(data),
        "sse_index": sse_index,
        "source_mode": SOURCE_MODE,
    }
    if diagnostics:
        snapshot["diagnostics"] = diagnostics
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")), encoding="utf-8"
    )
    unresolved = sorted(
        code for code, quote in estimates.items() if quote.get("gszzl") is None
    )
    write_status(
        "fetch_validate",
        "success" if len(estimates) == len(codes) and sse_index else "failed",
        trade_date=today,
        estimate_count=len(estimates),
        expected_count=len(codes),
        failure_count=len(failures),
        failed_codes=failures,
        turnover_trillion=(sse_index or {}).get("market_turnover_trillion"),
        snapshot_updated_at=snapshot["updated_at"],
        source_mode=SOURCE_MODE,
        unresolved=unresolved,
    )
    print(f"Published {len(estimates)}/{len(codes)} estimates to {OUTPUT_PATH}")
    if failures:
        print("Failed:", ", ".join(failures))
    return 0 if estimates else 2


if __name__ == "__main__":
    raise SystemExit(main())
