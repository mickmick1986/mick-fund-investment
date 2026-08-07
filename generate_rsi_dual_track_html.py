#!/usr/bin/env python3
"""Render the standalone RSI dual-track validation page."""
from __future__ import annotations

import html
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
INPUT_PATH = ROOT / "rsi_dual_track_validation.json"
OUTPUT_PATH = ROOT / "deploy" / "rsi_dual_track_comparison.html"


def esc(value: object) -> str:
    return html.escape("—" if value is None else str(value))


def badge(text: object, tone: str = "neutral") -> str:
    return f'<span class="badge {tone}">{esc(text)}</span>'


def rsi_tone(value: float | None) -> str:
    if value is None:
        return "neutral"
    if value < 30:
        return "down"
    if value > 70:
        return "up"
    return "neutral"


def decision_tone(changed: bool) -> str:
    return "changed" if changed else "same"


def render_chart(item: dict) -> str:
    simple = item.get("simple_series") or []
    wilder = item.get("wilder_series") or []
    values = [float(v) for v in simple + wilder if v is not None]
    if not values:
        return '<div class="chart-empty">历史序列不足</div>'
    width, height = 520, 180
    left, right, top, bottom = 34, 12, 12, 24
    plot_w, plot_h = width - left - right, height - top - bottom
    count = max(len(simple), len(wilder), 2)

    def points(series: list[float]) -> str:
        if len(series) == 1:
            series = series * 2
        return " ".join(
            f"{left + i * plot_w / (count - 1):.1f},{top + plot_h - max(0, min(100, float(value))) / 100 * plot_h:.1f}"
            for i, value in enumerate(series)
        )

    y30 = top + plot_h * 0.7
    y40 = top + plot_h * 0.6
    y60 = top + plot_h * 0.4
    y70 = top + plot_h * 0.3
    labels = "".join([
        f'<line x1="{left}" y1="{y:.1f}" x2="{width-right}" y2="{y:.1f}" class="grid"/>'
        for y in (y30, y40, y60, y70)
    ])
    label_text = "".join([
        f'<text x="{left-7}" y="{y+3:.1f}" text-anchor="end" class="axis">{label}</text>'
        for y, label in ((y30, "30"), (y40, "40"), (y60, "60"), (y70, "70"))
    ])
    return f'''<svg viewBox="0 0 {width} {height}" role="img" aria-label="{esc(item['name'])} RSI双轨曲线">
      <rect x="{left}" y="{top}" width="{plot_w}" height="{plot_h}" class="plot"/>
      {labels}{label_text}
      <polyline points="{points(simple)}" class="simple-line"/>
      <polyline points="{points(wilder)}" class="wilder-line"/>
      <text x="{left}" y="{height-5}" class="legend simple-label">简单平均</text>
      <text x="{left+100}" y="{height-5}" class="legend wilder-label">Wilder平滑</text>
      <text x="{left+plot_w}" y="{height-5}" text-anchor="end" class="legend">最近 →</text>
    </svg>'''


def main() -> int:
    data = json.loads(INPUT_PATH.read_text(encoding="utf-8"))
    summary = data["summary"]
    funds = data.get("funds", [])
    changed = [item for item in funds if item.get("decision_changed")]
    biggest = sorted(
        [item for item in funds if item.get("rsi_delta") is not None],
        key=lambda item: abs(item["rsi_delta"]), reverse=True
    )
    cards = []
    for item in funds:
        diff_class = decision_tone(item.get("decision_changed", False))
        delta = item.get("rsi_delta")
        delta_text = f"{delta:+.1f}" if delta is not None else "—"
        cards.append(f'''<article class="fund-card {diff_class}">
          <div class="fund-head"><div><strong>{esc(item['name'])}</strong><span class="code">{esc(item['code'])}</span></div>{badge('决策发生变化' if item.get('decision_changed') else '决策一致', 'changed' if item.get('decision_changed') else 'same')}</div>
          <div class="rsi-pair"><div><small>简单平均 RSI</small><b class="{rsi_tone(item.get('simple_rsi'))}">{esc(item.get('simple_rsi'))}</b>{badge(item.get('simple_signal'))}</div><div class="arrow">→</div><div><small>Wilder RSI</small><b class="{rsi_tone(item.get('wilder_rsi'))}">{esc(item.get('wilder_rsi'))}</b>{badge(item.get('wilder_signal'))}</div><div class="delta">差值<br><strong>{delta_text}</strong></div></div>
          <div class="decision-grid"><div><small>正式建议</small><strong>{esc(item.get('formal_action'))}</strong><span>{item['simple_order']['shares']}份 / ¥{item['simple_order']['amount']:,.0f}</span></div><div><small>Wilder模拟</small><strong>{esc(item.get('wilder_action'))}</strong><span>{item['wilder_order']['shares']}份 / ¥{item['wilder_order']['amount']:,.0f}</span></div><div><small>止盈触发</small><span>{'是' if item.get('formal_tp_trigger') else '否'} → {'是' if item.get('wilder_tp_trigger') else '否'}</span></div></div>
          <details><summary>查看RSI走势（{item.get('history_point_count', 0)}个确认净值点，Wilder有效{item.get('wilder_usable_points', 0)}点）</summary>{render_chart(item)}</details>
        </article>''')

    highlight = biggest[:3]
    highlight_html = "".join(
        f'<li><strong>{esc(item["name"])}</strong>：{item.get("simple_rsi")} → {item.get("wilder_rsi")}（{item.get("rsi_delta"):+.1f}），最终建议仍为{esc(item.get("formal_action"))}</li>'
        for item in highlight
    )
    conclusion = (
        f"本次确认快照中，Wilder RSI 与简单 RSI 的<strong>数值差异已经存在</strong>。"
        f"其中 <strong>{len(changed)} 只</strong>基金的 Wilder 模拟建议与正式建议不同；"
        "这只用于验证，不会改写正式操作建议、补仓份数或止盈触发。"
    )
    html_text = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>RSI双轨验证对比</title><style>
      *{{box-sizing:border-box}}body{{margin:0;background:#f3f6f8;color:#263238;font:14px/1.55 "Microsoft YaHei","PingFang SC",sans-serif}}.wrap{{max-width:1180px;margin:auto;padding:18px}}header{{background:#243b53;color:#fff;padding:20px 22px;border-radius:8px 8px 0 0}}h1{{margin:0 0 5px;font-size:22px}}header p{{margin:0;color:#dbe7f2}}.note{{background:#fff8e1;border:1px solid #eed58a;padding:12px 14px;margin:12px 0;border-radius:6px;color:#654f00}}.stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0}}.stat{{background:#fff;border:1px solid #d9e2ec;padding:14px;border-radius:6px}}.stat b{{display:block;font-size:25px;color:#1d5f82}}.stat span{{color:#607d8b;font-size:12px}}.section{{margin:18px 0}}h2{{font-size:17px;margin:0 0 9px}}.insight{{background:#fff;border-left:4px solid #4a90d9;padding:12px 15px}}.insight ul{{margin:6px 0 0;padding-left:20px}}.fund-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}}.fund-card{{background:#fff;border:1px solid #d9e2ec;border-radius:6px;padding:14px}}.fund-card.changed{{border-color:#e0a33a;background:#fffdf5}}.fund-head,.rsi-pair,.decision-grid{{display:flex;align-items:center;gap:10px}}.fund-head{{justify-content:space-between;border-bottom:1px solid #edf1f4;padding-bottom:9px}}.code{{margin-left:8px;color:#78909c;font:12px monospace}}.badge{{display:inline-block;padding:2px 7px;border-radius:10px;background:#eceff1;color:#546e7a;font-size:11px;white-space:nowrap}}.badge.same{{background:#e8f5e9;color:#2e7d32}}.badge.changed{{background:#fff0c2;color:#8a5a00}}.rsi-pair{{padding:13px 0;justify-content:space-between}}.rsi-pair>div:not(.arrow):not(.delta){{display:flex;flex-direction:column;gap:3px}}small{{color:#78909c;font-size:11px}}.rsi-pair b{{font:700 24px monospace}}.up{{color:#c62828}}.down{{color:#2e7d32}}.neutral{{color:#546e7a}}.arrow{{font-size:22px;color:#90a4ae}}.delta{{text-align:center;color:#78909c;font-size:11px}}.delta strong{{font:700 16px monospace;color:#455a64}}.decision-grid{{align-items:stretch;border-top:1px solid #edf1f4;padding-top:10px}}.decision-grid>div{{display:flex;flex-direction:column;gap:3px;flex:1}}.decision-grid strong{{color:#37474f}}.decision-grid span{{font-size:12px;color:#607d8b}}details{{margin-top:11px;border-top:1px solid #edf1f4;padding-top:8px}}summary{{cursor:pointer;color:#2c6e9f;font-size:12px}}svg{{width:100%;height:auto;margin-top:8px;background:#fbfcfd;border:1px solid #edf1f4}}.plot{{fill:#fff}}.grid{{stroke:#cfd8dc;stroke-dasharray:3 3;stroke-width:1}}.axis,.legend{{fill:#78909c;font-size:11px}}.simple-line{{fill:none;stroke:#e56b5d;stroke-width:2.4}}.wilder-line{{fill:none;stroke:#3d8f9d;stroke-width:2.4}}.simple-label{{fill:#e56b5d}}.wilder-label{{fill:#3d8f9d}}.chart-empty{{padding:15px;color:#90a4ae}}footer{{color:#78909c;font-size:12px;margin-top:16px}}@media(max-width:760px){{.wrap{{padding:8px}}h1{{font-size:18px}}.stats{{grid-template-columns:repeat(2,1fr)}}.fund-grid{{grid-template-columns:1fr}}.rsi-pair{{gap:6px}}.rsi-pair b{{font-size:21px}}}}
    </style></head><body><main class="wrap"><header><h1>RSI(14) 双轨验证对比</h1><p>确认日期：{esc(data.get('as_of_date'))} · 生成时间：{esc(data.get('generated_at'))}</p></header><div class="note"><strong>阅读方法：</strong>简单平均 RSI 仍是正式决策轨；Wilder RSI 只做模拟验证。页面中的“Wilder模拟”不会写回 Excel、确认层 JSON，也不会改变止盈锚点。当前数据源为{esc(data.get('source_mode'))}，盘中实时估值需单独生成实时验证快照。</div><section class="stats"><div class="stat"><b>{summary['fund_count']}</b><span>参与基金</span></div><div class="stat"><b>{summary['same_action_count']}</b><span>两轨建议一致</span></div><div class="stat"><b>{summary['action_changed_count']}</b><span>建议发生变化</span></div><div class="stat"><b>{summary['tp_trigger_changed_count']}</b><span>止盈触发变化</span></div></section><section class="section"><h2>先看结论</h2><div class="insight"><div>{conclusion}</div><ul>{highlight_html}</ul></div></section><section class="section"><h2>逐基金对比</h2><div class="fund-grid">{''.join(cards)}</div></section><footer>说明：Wilder RSI 采用经典递推平滑。当前确认快照保留45个净值点，Wilder有效历史点数会随后续确认刷新持续增加。</footer></main></body></html>'''
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(html_text, encoding="utf-8")
    print(f"Generated {OUTPUT_PATH} ({len(html_text):,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
