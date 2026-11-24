#!/usr/bin/env python3
"""2026-11-24 新增：stale 报价「按只降级」而不是「整帧拒收」的回归测试。

背景：10-30~11-24 共 46 次发布失败里 45 次集中在开盘 09:30-10:00，
根因是任何一只基金的 quote_date 还没翻到当日就 return 3 丢弃整帧。

新规则：
  · stale 只数 <= MAX_DEGRADABLE_STALE(8) → 该只 gszzl=None + src='stale'，其余照发，退出码 0
  · stale 只数 >  8                        → 仍按 fail-closed 整帧拒收，退出码 3，不写文件
"""
from __future__ import annotations

import importlib.util
import json
import shutil
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("fle", ROOT / "fetch_live_estimates.py")
fle = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fle)

RESULTS = []


def check(label, actual, expected):
    ok = actual == expected
    RESULTS.append(ok)
    print(f"  {'PASS' if ok else 'FAIL'}  {label} -> {actual!r}" + ("" if ok else f" (期望 {expected!r})"))


def run_case(stale_count: int):
    """伪造 25 只基金的抓取结果，其中 stale_count 只日期仍是昨天。"""
    codes = [str(item["code"]).zfill(6) for item in json.loads(fle.SOURCE_PATH.read_text(encoding="utf-8"))["funds"]]
    today = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d")
    yesterday = (datetime.now(ZoneInfo("Asia/Shanghai")) - timedelta(days=1)).strftime("%Y-%m-%d")
    stale_set = set(codes[:stale_count])

    def fake_fetch_one(session, code):
        return {
            "gztime": f"{today} 13:30:00",
            "quote_date": yesterday if code in stale_set else today,
            "gszzl": 1.2345,
            "src": "s1",
        }

    tmp = Path(tempfile.mkdtemp(prefix="live_stale_"))
    original = (fle.fetch_one, fle.fetch_sse_index, fle.OUTPUT_PATH, fle.SOURCE_MODE,
                fle.STATUS_DIR, fle.STATUS_PATH, fle.STATUS_HISTORY_PATH)
    fle.fetch_one = fake_fetch_one
    fle.fetch_sse_index = lambda session: {"market_turnover_trillion": 2.15, "market_turnover_trillion_prev": 2.0}
    fle.SOURCE_MODE = "legacy"
    fle.OUTPUT_PATH = tmp / "live_estimates.json"
    fle.STATUS_DIR = tmp
    fle.STATUS_PATH = tmp / "status.json"
    fle.STATUS_HISTORY_PATH = tmp / "status.jsonl"
    try:
        rc = fle.main()
        payload = json.loads(fle.OUTPUT_PATH.read_text(encoding="utf-8")) if fle.OUTPUT_PATH.exists() else None
    finally:
        (fle.fetch_one, fle.fetch_sse_index, fle.OUTPUT_PATH, fle.SOURCE_MODE,
         fle.STATUS_DIR, fle.STATUS_PATH, fle.STATUS_HISTORY_PATH) = original
    return rc, payload, codes, stale_set, tmp


print("== 1. 1 只 stale：降级发布，其余 24 只照发 ==")
rc, snap, codes, stale_set, tmp = run_case(1)
check("退出码", rc, 0)
check("快照已写出", snap is not None, True)
check("条目数保持 25", len(snap["estimates"]) if snap else None, 25)
check("estimate_count 仍为 25", snap["estimate_count"] if snap else None, 25)
bad = codes[0]
check("stale 只 gszzl 置空", snap["estimates"][bad]["gszzl"] if snap else "no-snap", None)
check("stale 只 src 标记", snap["estimates"][bad]["src"] if snap else "no-snap", "stale")
check("stale 只保留原报价日期", snap["estimates"][bad]["stale_quote_date"] if snap else "no-snap",
      snap["estimates"][bad]["quote_date"] if snap else "no-snap")
check("正常只不受影响", snap["estimates"][codes[1]]["gszzl"] if snap else None, 1.2345)
shutil.rmtree(tmp, ignore_errors=True)

print("== 2. 8 只 stale（阈值边界）：仍降级 ==")
rc, snap, codes, stale_set, tmp = run_case(8)
check("退出码", rc, 0)
check("快照已写出", snap is not None, True)
dropped = [c for c in codes if snap and snap["estimates"][c]["gszzl"] is None]
check("降级只数", len(dropped), 8)
shutil.rmtree(tmp, ignore_errors=True)

print("== 3. 9 只 stale（超阈值）：整帧拒收，绝不伪造 ==")
rc, snap, codes, stale_set, tmp = run_case(9)
check("退出码", rc, 3)
check("快照未写出", snap is None, True)
shutil.rmtree(tmp, ignore_errors=True)

failed = RESULTS.count(False)
print(f"\n结果：通过 {len(RESULTS) - failed} 项，失败 {failed} 项")
raise SystemExit(0 if failed == 0 else 1)
