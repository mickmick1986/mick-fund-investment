#!/usr/bin/env python3
"""盘中估值三源对照工具（只读，不发布）。

对 25 只基金同时算出：
  S1 = 源端估算净值口径
  B  = 源端估算涨跌幅（第 10 位）
  C  = 指数/ETF 行情 × β
  最终 = resolve_estimate 的仲裁结果（s1 / b / idx / none）

用途：① 上线当日验收报告；② 影子期逐日对比 β 是否合理。

用法：
  python tools/compare_live_sources.py            # 控制台表格
  python tools/compare_live_sources.py --html reports/xxx.html
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import requests  # noqa: E402

import fetch_live_estimates as fle  # noqa: E402

SRC_LABEL = {
    "s1": "源端估值",
    "b": "源端涨幅",
    "idx": "指数推算",
    "none": "估值暂缺",
}


def collect():
    config = fle.load_live_index_map()
    source_path = ROOT / "fund_data_enriched.json"
    data = json.loads(source_path.read_text(encoding="utf-8"))
    codes = [str(item["code"]).zfill(6) for item in data.get("funds", [])]
    names = {str(item["code"]).zfill(6): item.get("name", "") for item in data.get("funds", [])}

    with requests.Session() as session:
        symbols = sorted(
            {
                entry.get("quote")
                for entry in config["funds"].values()
                if isinstance(entry, dict) and entry.get("quote")
            }
        )
        quotes = fle.fetch_index_quotes(session, symbols)
        rows = []
        for code in codes:
            fields = fle.fetch_one_fields(session, code)
            parsed = fle.parse_confirmed_quote(fields)
            s1 = b = None
            if parsed:
                gap = parsed["estimate_nav"] - parsed["confirmed_nav"]
                if abs(gap) > max(1e-9, parsed["confirmed_nav"] * 1e-9):
                    s1 = gap / parsed["confirmed_nav"] * 100
                if len(fields) >= 10:
                    try:
                        cand = float(fields[9])
                    except (TypeError, ValueError):
                        cand = None
                    if cand is not None and abs(cand) > 1e-4:
                        b = cand
            entry = config["funds"].get(code) or {}
            q = quotes.get(entry.get("quote") or "")
            c = None
            if entry.get("passive") and q and any(
                k in (q.get("name") or "") for k in (entry.get("keywords") or [])
            ):
                c = q["change_pct"] * (entry.get("beta") or config["beta_default"])
            item = fle.resolve_estimate(code, fields, quotes, config)
            rows.append(
                {
                    "code": code,
                    "name": names.get(code, ""),
                    "passive": bool(entry.get("passive")),
                    "quote": entry.get("quote"),
                    "quote_name": (q or {}).get("name"),
                    "idx_change": round(q["change_pct"], 4) if q else None,
                    "s1": round(s1, 4) if s1 is not None else None,
                    "b": round(b, 4) if b is not None else None,
                    "c": round(c, 4) if c is not None else None,
                    "final": (item or {}).get("gszzl"),
                    "src": (item or {}).get("src"),
                    "note": ((item or {}).get("_diag") or {}).get("note"),
                }
            )
    return rows


def fmt(value, digits=2):
    return "—" if value is None else f"{value:+.{digits}f}%"


def render_console(rows):
    print(f"{'代码':7}{'名称':22}{'S1':>9}{'B':>9}{'C':>9}{'最终':>9}{'来源':>9}  备注")
    for r in rows:
        print(
            f"{r['code']:7}{r['name'][:20]:22}{fmt(r['s1']):>9}{fmt(r['b']):>9}"
            f"{fmt(r['c']):>9}{fmt(r['final']):>9}{SRC_LABEL.get(r['src'], '-'):>9}  {r['note'] or ''}"
        )
    diff = [r for r in rows if r["src"] not in ("s1", None)]
    print()
    print(f"总 {len(rows)} 只；非 S1 路径 {len(diff)} 只：" + "，".join(
        f"{r['code']}({SRC_LABEL.get(r['src'])})" for r in diff
    ))
    conflicts = [r for r in rows if r["b"] is not None and r["c"] is not None and abs(r["b"] - r["c"]) > 1.0]
    if conflicts:
        print("B 与指数推算偏差 >1pp：" + "，".join(
            f"{r['code']} B={r['b']:+.2f}% C={r['c']:+.2f}%" for r in conflicts
        ))


def render_html(rows, generated_at):
    color = {"s1": "#1f6feb", "b": "#e67e22", "idx": "#8e44ad", "none": "#95a5a6"}

    def cell(value, cls=""):
        if value is None:
            return f'<td class="na {cls}">—</td>'
        tone = "up" if value > 0 else ("down" if value < 0 else "flat")
        return f'<td class="num {tone} {cls}">{value:+.2f}%</td>'

    body = []
    for r in rows:
        badge = SRC_LABEL.get(r["src"], "-")
        body.append(
            "<tr>"
            f'<td class="code">{r["code"]}</td>'
            f'<td class="name">{r["name"]}</td>'
            f'<td class="center">{"被动" if r["passive"] else "主动"}</td>'
            f'<td class="center small">{r["quote"] or "—"}<br><span class="qn">{r["quote_name"] or ""}</span></td>'
            + cell(r["s1"]) + cell(r["b"]) + cell(r["c"])
            + cell(r["final"])
            + f'<td class="center"><span class="badge" style="background:{color.get(r["src"], "#95a5a6")}">{badge}</span></td>'
            f'<td class="note">{r["note"] or ""}</td>'
            "</tr>"
        )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>盘中估值三源对照 {generated_at}</title>
<style>
body{{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#0d1117;color:#e6edf3;margin:24px}}
h1{{font-size:22px;margin:0 0 6px}}
.sub{{color:#8b949e;font-size:13px;margin-bottom:16px}}
table{{border-collapse:collapse;width:100%;font-size:14px}}
th,td{{border:1px solid #30363d;padding:7px 9px}}
th{{background:#161b22;font-size:13px;color:#c9d1d9;position:sticky;top:0}}
tr:nth-child(even){{background:#0f141a}}
.num{{text-align:right;font-family:Consolas,monospace;font-weight:bold}}
.up{{color:#ff6b6b}} .down{{color:#3fb950}} .flat{{color:#95a5a6}}
.na{{text-align:right;color:#6e7681}}
.center{{text-align:center}} .small{{font-size:11px;color:#8b949e}}
.qn{{color:#6e7681;font-size:10px}}
.code{{font-family:Consolas,monospace;color:#79c0ff}}
.name{{font-weight:600}}
.note{{font-size:11px;color:#8b949e;max-width:320px}}
.badge{{color:#fff;padding:2px 8px;border-radius:10px;font-size:12px;white-space:nowrap}}
.legend{{margin-top:14px;font-size:12px;color:#8b949e;line-height:1.9}}
</style></head><body>
<h1>盘中估值三源对照 · {generated_at}</h1>
<div class="sub">S1 = 源端估算净值口径 ｜ B = 源端估算涨跌幅 ｜ C = 指数/ETF 行情 × β ｜ 最终 = 仲裁结果</div>
<table>
<thead><tr><th>代码</th><th>名称</th><th>类型</th><th>行情源</th><th>S1</th><th>B</th><th>C</th><th>最终取值</th><th>来源</th><th>仲裁说明</th></tr></thead>
<tbody>{''.join(body)}</tbody></table>
<div class="legend">
仲裁规则：S1 可用即用 S1；S1 为占位时，若 B 与 C 偏差 ≤1pp 用 B，否则以指数为锚（C）；C 不可用（主动基金/名称校验失败/行情缺失/幅度异常）时显示「估值暂缺」，绝不伪造 0.00%。<br>
颜色：红=涨，绿=跌（A 股习惯）。
</div>
</body></html>"""
    return html


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--html", help="导出 HTML 报告路径")
    args = parser.parse_args()

    rows = collect()
    generated_at = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M:%S")
    render_console(rows)
    if args.html:
        path = Path(args.html)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(render_html(rows, generated_at), encoding="utf-8")
        print(f"\nHTML 报告已写入 {path}")
    out = ROOT / ".workbuddy" / "live_source_compare" / f"{generated_at[:10]}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps({"generated_at": generated_at, "rows": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"对照数据已存档 {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
