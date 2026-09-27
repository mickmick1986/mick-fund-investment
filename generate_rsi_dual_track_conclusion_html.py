#!/usr/bin/env python3
"""Generate a 20-confirmed-trading-day RSI dual-track conclusion page.

This is a read-only reporting layer. It only reads the dual-track history and
records completed report cycles; it never changes the formal RSI, Excel workbook,
confirmed fund data, or manual anchors.
"""
from __future__ import annotations

import html
import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
HISTORY_PATH = ROOT / "rsi_dual_track_history.json"
STATE_PATH = ROOT / "rsi_dual_track_cycle_state.json"
DESKTOP_DIR = Path(r"C:\Users\13697\Desktop\金字塔丛林战法")
DESKTOP_LATEST_PATH = DESKTOP_DIR / "20天RSI双轨验证对比结论.html"
PUBLIC_LATEST_PATH = ROOT / "deploy" / "rsi_dual_track_20day_conclusion.html"
CYCLE_DAYS = 20


def esc(value: object) -> str:
    return html.escape("—" if value is None else str(value))


def confirmed_snapshots() -> list[dict[str, Any]]:
    if not HISTORY_PATH.exists():
        return []
    try:
        history = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    deduped: dict[str, dict[str, Any]] = {}
    for item in history:
        if item.get("snapshot_key", "").endswith("|final_nav") and item.get("snapshot_key"):
            deduped[item["snapshot_key"]] = item
    return sorted(deduped.values(), key=lambda item: item["snapshot_key"].split("|", 1)[0])


def load_state() -> dict[str, Any]:
    if not STATE_PATH.exists():
        return {"cycle_days": CYCLE_DAYS, "settled_cycle_end_keys": []}
    try:
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        state = {}
    state.setdefault("cycle_days", CYCLE_DAYS)
    state.setdefault("settled_cycle_end_keys", [])
    return state


def pending_cycles(snapshots: list[dict[str, Any]], state: dict[str, Any]) -> list[list[dict[str, Any]]]:
    settled = set(state.get("settled_cycle_end_keys", []))
    start = 0
    for index, snapshot in enumerate(snapshots):
        if snapshot.get("snapshot_key") in settled:
            start = index + 1
    cycles = []
    while len(snapshots) - start >= CYCLE_DAYS:
        cycles.append(snapshots[start:start + CYCLE_DAYS])
        start += CYCLE_DAYS
    return cycles


def amount(item: dict[str, Any], track: str) -> float:
    return float((item.get(track) or {}).get("amount") or 0)


def build_report(cycle: list[dict[str, Any]]) -> tuple[str, dict[str, Any]]:
    fund_stats: dict[str, dict[str, Any]] = defaultdict(lambda: {
        "name": "", "code": "", "action_days": [], "buy_days": [], "tp_days": [],
        "rsi_deltas": [], "examples": []
    })
    total_fund_days = 0
    action_events = buy_events = tp_events = 0
    for snapshot in cycle:
        date = snapshot["snapshot_key"].split("|", 1)[0]
        for fund in snapshot.get("funds", []):
            total_fund_days += 1
            code = str(fund.get("code", "")).zfill(6)
            stat = fund_stats[code]
            stat["name"] = fund.get("name", code)
            stat["code"] = code
            delta = fund.get("rsi_delta")
            if delta is not None:
                stat["rsi_deltas"].append(float(delta))
            changed = bool(fund.get("decision_changed"))
            buy_changed = bool(fund.get("buy_changed"))
            tp_changed = bool(fund.get("tp_trigger_changed"))
            if changed:
                action_events += 1
                stat["action_days"].append(date)
            if buy_changed:
                buy_events += 1
                stat["buy_days"].append(date)
            if tp_changed:
                tp_events += 1
                stat["tp_days"].append(date)
            if changed or buy_changed or tp_changed:
                stat["examples"].append({
                    "date": date,
                    "simple_rsi": fund.get("simple_rsi"),
                    "wilder_rsi": fund.get("wilder_rsi"),
                    "formal_action": fund.get("formal_action"),
                    "wilder_action": fund.get("wilder_action"),
                    "formal_amount": amount(fund, "simple_order"),
                    "wilder_amount": amount(fund, "wilder_order"),
                    "formal_tp": bool(fund.get("formal_tp_trigger")),
                    "wilder_tp": bool(fund.get("wilder_tp_trigger")),
                })

    start_date = cycle[0]["snapshot_key"].split("|", 1)[0]
    end_date = cycle[-1]["snapshot_key"].split("|", 1)[0]
    same_action_rate = (total_fund_days - action_events) / total_fund_days * 100 if total_fund_days else 100
    divergent = [item for item in fund_stats.values() if item["action_days"] or item["buy_days"] or item["tp_days"]]
    divergent.sort(key=lambda item: (len(item["action_days"]), len(item["buy_days"]), len(item["tp_days"])), reverse=True)

    if action_events == 0 and buy_events == 0 and tp_events == 0:
        recommendation = "本周期两种 RSI 在建议、补仓份数与止盈触发上完全一致。现阶段没有足够的实际差异证据支持切换正式轨，建议继续使用现行简单平均 RSI(14)，并继续观察下一周期。"
    elif action_events <= max(3, total_fund_days * 0.01) and tp_events == 0:
        recommendation = "Wilder RSI 已出现少量模拟差异，但频率较低，且尚未构成切换正式轨的充分证据。建议维持现行简单平均 RSI(14) 为正式轨，继续完成下一周期验证。"
    else:
        recommendation = "Wilder RSI 在本周期内已多次改变模拟建议、补仓份数或止盈触发。它值得进入人工复核，但结论页不会自动切换正式轨；请结合差异日期的后续净值表现再决定是否升级。"

    rows = []
    for item in divergent:
        examples = "<br>".join(
            f"{esc(example['date'])}：RSI {esc(example['simple_rsi'])} → {esc(example['wilder_rsi'])}；"
            f"建议 {esc(example['formal_action'])} → {esc(example['wilder_action'])}；"
            f"补仓 ¥{example['formal_amount']:,.0f} → ¥{example['wilder_amount']:,.0f}；"
            f"止盈 {'是' if example['formal_tp'] else '否'} → {'是' if example['wilder_tp'] else '否'}"
            for example in item["examples"][:4]
        )
        avg_delta = sum(item["rsi_deltas"]) / len(item["rsi_deltas"]) if item["rsi_deltas"] else 0
        rows.append(
            f"<tr><td><strong>{esc(item['name'])}</strong><br><span>{esc(item['code'])}</span></td>"
            f"<td>{len(item['action_days'])}</td><td>{len(item['buy_days'])}</td><td>{len(item['tp_days'])}</td>"
            f"<td>{avg_delta:+.1f}</td><td class=\"examples\">{examples}</td></tr>"
        )
    detail_html = "".join(rows) or '<tr><td colspan="6" class="empty">本周期没有任何基金出现实际决策、补仓份数或止盈触发差异。</td></tr>'
    report = {
        "cycle_start": start_date,
        "cycle_end": end_date,
        "confirmed_days": len(cycle),
        "fund_day_observations": total_fund_days,
        "action_events": action_events,
        "buy_events": buy_events,
        "tp_events": tp_events,
        "same_action_rate": round(same_action_rate, 2),
        "divergent_fund_count": len(divergent),
        "recommendation": recommendation,
    }
    page = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>20天 RSI 双轨验证对比结论</title><style>
*{{box-sizing:border-box}}body{{margin:0;background:#f3f6f8;color:#263238;font:15px/1.6 "Microsoft YaHei","PingFang SC",sans-serif}}.wrap{{max-width:1200px;margin:auto;padding:20px}}header{{background:#243b53;color:#fff;padding:24px;border-radius:8px 8px 0 0}}h1{{margin:0 0 6px;font-size:25px}}header p{{margin:0;color:#dce8f2}}.notice{{margin:15px 0;background:#fff8e1;border:1px solid #edd891;border-radius:6px;padding:13px 15px;color:#654f00}}.stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:11px;margin:15px 0}}.stat{{background:#fff;border:1px solid #d7e1e8;border-radius:6px;padding:15px}}.stat b{{display:block;font-size:27px;color:#1d5f82}}.stat span{{font-size:12px;color:#607d8b}}section{{margin:20px 0}}h2{{font-size:18px;margin:0 0 10px}}.conclusion{{background:#fff;border-left:4px solid #4a90d9;padding:15px 17px}}table{{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d7e1e8}}th{{background:#eaf1f6;color:#385166;text-align:left}}th,td{{border:1px solid #d7e1e8;padding:10px;vertical-align:top}}td span{{font:12px monospace;color:#78909c}}.examples{{font-size:13px;min-width:410px}}.empty{{text-align:center;color:#607d8b;padding:24px}}footer{{color:#607d8b;font-size:12px;padding:8px 0 20px}}@media(max-width:760px){{.wrap{{padding:10px}}.stats{{grid-template-columns:repeat(2,1fr)}}.examples{{min-width:260px}}table{{font-size:13px}}th,td{{padding:7px}}}}</style></head><body><main class="wrap"><header><h1>20 个确认交易日 RSI 双轨验证对比结论</h1><p>验证区间：{esc(start_date)} 至 {esc(end_date)} · 结论生成：{esc(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</p></header><div class="notice"><strong>边界：</strong>简单平均 RSI(14) 仍是唯一正式决策轨。Wilder 平滑 RSI(14) 仅用于本结论的回测式模拟统计；本页面不会改写 Excel、确认层数据、人工锚点或止盈锚点，也不会自动切换策略。</div><section class="stats"><div class="stat"><b>{len(cycle)}</b><span>确认交易日</span></div><div class="stat"><b>{same_action_rate:.2f}%</b><span>建议一致率</span></div><div class="stat"><b>{action_events}</b><span>建议差异基金日</span></div><div class="stat"><b>{buy_events + tp_events}</b><span>补仓/止盈差异基金日</span></div></section><section><h2>结论</h2><div class="conclusion">{esc(recommendation)}</div></section><section><h2>统计口径</h2><div class="conclusion">覆盖 <strong>{total_fund_days}</strong> 个“基金 × 确认交易日”观察值；建议差异 <strong>{action_events}</strong> 次，补仓份数差异 <strong>{buy_events}</strong> 次，止盈触发差异 <strong>{tp_events}</strong> 次，涉及 <strong>{len(divergent)}</strong> 只基金。</div></section><section><h2>出现差异的基金与日期</h2><table><thead><tr><th>基金</th><th>建议差异次数</th><th>补仓差异次数</th><th>止盈差异次数</th><th>平均 RSI 差值</th><th>差异样本（最多 4 条）</th></tr></thead><tbody>{detail_html}</tbody></table></section><footer>说明：统计仅纳入 `final_nav` 确认净值快照，自动忽略盘中估值快照。同一确认日期重复运行会按日期去重，不重复计入 20 天周期。</footer></main></body></html>'''
    return page, report


def main() -> int:
    snapshots = confirmed_snapshots()
    state = load_state()
    cycles = pending_cycles(snapshots, state)
    if not cycles:
        settled_count = len(state.get("settled_cycle_end_keys", []))
        completed = settled_count * CYCLE_DAYS
        print(f"RSI 20天结论尚未到期：已结算{completed}天，当前待累计{len(snapshots) - completed}个确认交易日。")
        return 0
    for cycle in cycles:
        page, report = build_report(cycle)
        end_date = report["cycle_end"]
        archive_path = DESKTOP_DIR / f"RSI双轨验证结论_{report['cycle_start']}_至_{end_date}.html"
        DESKTOP_DIR.mkdir(parents=True, exist_ok=True)
        PUBLIC_LATEST_PATH.parent.mkdir(parents=True, exist_ok=True)
        DESKTOP_LATEST_PATH.write_text(page, encoding="utf-8")
        archive_path.write_text(page, encoding="utf-8")
        PUBLIC_LATEST_PATH.write_text(page, encoding="utf-8")
        state["settled_cycle_end_keys"].append(cycle[-1]["snapshot_key"])
        state["last_report"] = {**report, "desktop_path": str(DESKTOP_LATEST_PATH), "archive_path": str(archive_path)}
        print(f"Generated 20-day conclusion: {archive_path}")
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
