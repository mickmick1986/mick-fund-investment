#!/usr/bin/env python3
"""盘中估值三级兜底（S1 → B → C → 暂无）的回放单测。

不依赖网络：用手工构造的新浪 fu_ 字段样本验证仲裁逻辑与各拒绝分支。
运行： python tests/test_live_resolution.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import fetch_live_estimates as fle  # noqa: E402

TODAY = "2026-10-29"
CFG = fle.load_live_index_map()

passed = 0
failed = 0


def make_fields(name, est, conf, pct, date=TODAY, time="13:30:00", prev="1.0000"):
    """构造新浪 fu_ 返回的字段序列（索引 2=估算净值 3=确认净值 7=日期 9=估算涨跌幅）。"""
    return [
        name, time, str(est), str(conf), str(conf), "0", "0.0000",
        date, str(prev), str(pct),
    ]


def check(label, actual, expected):
    global passed, failed
    if actual == expected:
        passed += 1
        print(f"  PASS  {label} -> {actual}")
    else:
        failed += 1
        print(f"  FAIL  {label} -> got {actual!r}, want {expected!r}")


def resolve(code, fields, quotes=None):
    return fle.resolve_estimate(code, fields, quotes or {}, CFG)


print("== 1. S1：源端估算净值可用 ==")
item = resolve("000218", make_fields("国泰黄金ETF联接A", 3.1366, 3.1320, 0.1617, prev="3.1320"))
check("src", item["src"], "s1")
check("gszzl", item["gszzl"], 0.1469)
check("quote_date", item["quote_date"], TODAY)

print("== 2. B：估算净值被占位，退到源端估算涨幅 ==")
item = resolve("161725", make_fields("招商中证白酒A", 0.5285, 0.5285, 1.7988, time="13:33:00"))
check("src", item["src"], "b")
check("gszzl", item["gszzl"], 1.7988)
check("b 诊断值", item["_diag"]["b"], 1.7988)

print("== 3. C：B 也为 0，用指数×β ==")
quotes = {"sz980018": {"name": "卫星通信", "change_pct": 2.28}}
item = resolve("024194", make_fields("永赢国证商用卫星通信产业ETF联接A", 1.1151, 1.1151, 0), quotes)
check("src", item["src"], "idx")
check("gszzl = 2.28x0.95", item["gszzl"], 2.166)
check("idx 诊断值", item["_diag"]["idx"], 2.28)
check("beta", item["_diag"]["beta"], 0.95)

print("== 4. C 拒绝：名称校验失败（防前缀猜错命中个股）==")
bad = {"sz980018": {"name": "华菱钢铁", "change_pct": -0.87}}
item = resolve("024194", make_fields("永赢国证商用卫星通信产业ETF联接A", 1.1151, 1.1151, 0), bad)
check("src", item["src"], "none")
check("gszzl 必须为 None（不得伪造 0）", item["gszzl"], None)
check("note 含名称校验", "名称校验失败" in (item["_diag"]["note"] or ""), True)

print("== 5. C 拒绝：指数涨幅异常（>15%）==")
wild = {"sz980018": {"name": "卫星通信", "change_pct": 22.0}}
item = resolve("024194", make_fields("永赢国证商用卫星通信产业ETF联接A", 1.1151, 1.1151, 0), wild)
check("src", item["src"], "none")
check("note 含异常", "异常" in (item["_diag"]["note"] or ""), True)

print("== 6. C 拒绝：指数行情缺失 ==")
item = resolve("024194", make_fields("永赢国证商用卫星通信产业ETF联接A", 1.1151, 1.1151, 0), {})
check("src", item["src"], "none")
check("note 含缺失", "缺失" in (item["_diag"]["note"] or ""), True)

print("== 7. 主动基金不启用指数兜底 ==")
item = resolve("003984", make_fields("嘉实新能源新材料股票A", 1.8566, 1.8566, 0),
               {"sz399808": {"name": "中证新能", "change_pct": -1.03}})
check("src", item["src"], "none")
check("note 含主动基金", "主动基金" in (item["_diag"]["note"] or ""), True)

print("== 7b. 仲裁：B 与指数推算冲突（>1pp）→ 以指数为锚 ==")
conflict = {"sz159995": {"name": "芯片ETF", "change_pct": -3.69}}
item = resolve("008887", make_fields("华夏国证半导体芯片ETF", 1.5366, 1.5366, -1.2642), conflict)
check("src 应为 idx", item["src"], "idx")
check("gszzl = -3.69x0.95", item["gszzl"], -3.5055)
check("note 说明冲突", "以指数为锚" in (item["_diag"]["note"] or ""), True)

print("== 7c. C 为主：B 与指数同时可用时优先指数 ==")
agree = {"sz399997": {"name": "中证白酒", "change_pct": 1.32}}
item = resolve("161725", make_fields("招商中证白酒A", 0.5285, 0.5285, 1.50), agree)
check("src 应为 idx", item["src"], "idx")
check("gszzl = 1.32x0.95", item["gszzl"], 1.254)

print("== 7d. 代理标的无成交（C=0）且 B 明显非零 → 退回 B ==")
flat_proxy = {"sz159825": {"name": "农业ETF", "change_pct": 0.0}}
item = resolve("010769", make_fields("天弘中证农业主题ETF联接A", 0.7085, 0.7085, 0.60), flat_proxy)
check("src 应为 b", item["src"], "b")
check("gszzl 取 B", item["gszzl"], 0.6)

print("== 7e. S1 近零占位（+0.01%）而指数 -3.6% → 判占位并改走指数 ==")
chip = {"sz399006": {"name": "创业板指", "change_pct": -3.30}}
item = resolve("003765", make_fields("广发创业板ETF联接A", 1.0100, 1.0099, -3.13), chip)
check("src 应为 idx", item["src"], "idx")
check("gszzl = -3.30x0.95", item["gszzl"], -3.135)
check("note 含占位判定", "判定源端占位" in (item["_diag"]["note"] or ""), True)

print("== 7f. S1 小幅波动且指数也平（无背离）→ 仍采信 S1 ==")
item = resolve("003765", make_fields("广发创业板ETF联接A", 1.0103, 1.0100, 0.03),
               {"sz399006": {"name": "创业板指", "change_pct": 0.04}})
check("src 应为 s1", item["src"], "s1")

print("== 8. 条目必须保留（gszzl=None 而非丢弃）==")
item = resolve("024202", make_fields("永赢制造升级智选混合A", 1.3167, 1.3167, 0))
check("条目非 None", item is not None, True)
check("含 quote_date", item["quote_date"], TODAY)

print("== 9. 底层行情不可用时仍返回 None（沿用原有 fail-closed）==")
check("空字段", resolve("024194", None), None)
check("净值非正", resolve("024194", make_fields("x", 0, 1.1151, 0)), None)
check("字段过短", resolve("024194", ["x", "1", "1"]), None)
check("日期缺失", resolve("024194", make_fields("x", 1.1, 1.0, 0, date="")), None)

print("== 10. 历史行为对照（说明 0.00% 的来源）==")
legacy = fle.parse_confirmed_quote(make_fields("永赢国证商用卫星通信产业ETF联接A", 1.1151, 1.1151, 0))
legacy_pct = round((legacy["estimate_nav"] - legacy["confirmed_nav"]) / legacy["confirmed_nav"] * 100, 4)
check("legacy 口径算出的值", legacy_pct, 0.0)

print("== 11. 配置表完整性 ==")
check("配置表基金数", len(CFG["funds"]), 25)
passive_with_quote = [c for c, e in CFG["funds"].items() if e.get("passive") and e.get("quote")]
check("启用指数兜底的被动基金数", len(passive_with_quote), 16)
for code, entry in CFG["funds"].items():
    if entry.get("quote"):
        if not entry.get("keywords"):
            failed += 1
            print(f"  FAIL  {code} 有 quote 但缺 keywords（名称校验会失效）")

print()
print(f"结果：通过 {passed} 项，失败 {failed} 项")
raise SystemExit(0 if failed == 0 else 1)
