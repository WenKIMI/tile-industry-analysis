#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
瓷砖景气评分脚本 —— 读 data/raw.csv，输出「建议得分」和景气分。

定位：**辅助打分，不替代判断**。
脚本只做算术（同比、变化、加权），方向判断和最终得分必须你自己确认后
填回 scorecard.xlsx。这样做的原因：行业研究的价值在判断，不在算数。

用法
----
    py -3.13 src/score.py
    py -3.13 src/score.py --raw data/raw.csv --out data/scorecard_auto.csv

输出
----
    data/scorecard_auto.csv   建议得分表（可直接对照填 Excel）
    控制台                    景气分 + 分级 + 缺失指标清单

打分规则（可改 SCORE_RULES）
---------------------------
指标值本身是同比（%），所以"变化"用百分点（pp）差：
    正向指标：变化 ≥ +2pp → +2 ；≥ +0.5pp → +1 ；≤ -2pp → -2 ；≤ -0.5pp → -1 ；其余 0
    负向指标：方向取反
另外：正向指标从负转正（跨过 0）额外 +1，从正转负额外 -1。
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

# 与 indicator_tree.xlsx 保持一致（改权重时两边都要改）
INDICATORS = [
    # (类别, 指标, 方向, 权重, 是否手工)
    ("需求", "地产竣工面积同比", "正", 0.15, False),
    ("需求", "商品房销售面积同比", "正", 0.05, False),
    ("需求", "瓷砖产量同比", "正", 0.15, False),
    ("需求", "瓷砖出口金额同比", "正", 0.10, False),
    ("供给", "行业开工率", "正", 0.05, True),
    ("供给", "库存天数", "负", 0.05, True),
    ("价格", "瓷砖价格指数", "正", 0.10, True),
    ("成本", "能源成本（天然气）", "负", 0.05, False),
    ("盈利", "上市公司毛利率", "正", 0.10, False),
    ("信用", "应收账款周转天数", "负", 0.05, False),
    ("政策", "地产政策强度", "正", 0.10, True),
    ("政策", "环保/能耗政策", "负", 0.05, True),
]

SCORE_RULES = {
    "big_up": 2.0,     # 变化 ≥ +2pp
    "up": 0.5,         # 变化 ≥ +0.5pp
    "big_down": -2.0,  # 变化 ≤ -2pp
    "down": -0.5,      # 变化 ≤ -0.5pp
}


def suggest_score(delta_pp: float, direction: str, latest: float) -> tuple[int, str]:
    """按变化幅度给出建议得分，返回 (得分, 依据)。"""
    d = delta_pp if direction == "正" else -delta_pp
    r = SCORE_RULES

    if d >= r["big_up"]:
        score, why = 2, f"变化 {delta_pp:+.2f}pp，显著改善"
    elif d >= r["up"]:
        score, why = 1, f"变化 {delta_pp:+.2f}pp，温和改善"
    elif d <= r["big_down"]:
        score, why = -2, f"变化 {delta_pp:+.2f}pp，显著恶化"
    elif d <= r["down"]:
        score, why = -1, f"变化 {delta_pp:+.2f}pp，温和恶化"
    else:
        score, why = 0, f"变化 {delta_pp:+.2f}pp，基本持平"

    # 方向翻转额外加减一分（"同比转正/转负"往往比幅度更重要）
    if direction == "正":
        if latest > 0 >= latest - delta_pp:
            score, why = min(score + 1, 2), why + "；且由负转正"
        elif latest < 0 <= latest - delta_pp:
            score, why = max(score - 1, -2), why + "；且由正转负"
    else:
        if latest < 0 <= latest - delta_pp:
            score, why = min(score + 1, 2), why + "；负向指标转负（利好）"
        elif latest > 0 >= latest - delta_pp:
            score, why = max(score - 1, -2), why + "；负向指标转正（利空）"

    return int(score), why


def load_raw(raw_path: Path) -> pd.DataFrame:
    if not raw_path.exists():
        raise SystemExit(f"找不到 {raw_path}，请先按 data/raw.csv 的格式录入数据")
    df = pd.read_csv(raw_path, encoding="utf-8-sig")
    need = {"日期", "指标", "值"}
    missing = need - set(df.columns)
    if missing:
        raise SystemExit(f"raw.csv 缺列：{missing}（需要 日期, 指标, 值, 单位, 来源, 链接, 备注）")
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df["值"] = pd.to_numeric(df["值"], errors="coerce")
    df = df.dropna(subset=["日期", "值"])
    return df.sort_values("日期")


def main() -> None:
    ap = argparse.ArgumentParser(description="瓷砖景气评分（辅助）")
    ap.add_argument("--raw", default="data/raw.csv")
    ap.add_argument("--out", default="data/scorecard_auto.csv")
    args = ap.parse_args()

    df = load_raw(Path(args.raw))

    if (df["来源"].astype(str).str.contains("示例").any() if "来源" in df.columns else False):
        print("⚠️  raw.csv 里还有「示例」行，结果不可用，请先替换成真实数据\n")

    rows, missing = [], []
    total = 0.0

    for cat, name, direction, weight, manual in INDICATORS:
        sub = df[df["指标"] == name]
        if sub.empty:
            missing.append(name)
            rows.append({
                "类别": cat, "指标": name, "方向": direction, "权重": weight,
                "最新值": None, "上期值": None, "变化pp": None,
                "建议得分": 0, "依据": "无数据，按 0 处理（需补充）",
            })
            continue

        latest_row = sub.iloc[-1]
        latest = float(latest_row["值"])
        prev = float(sub.iloc[-2]["值"]) if len(sub) >= 2 else None

        if prev is None:
            rows.append({
                "类别": cat, "指标": name, "方向": direction, "权重": weight,
                "最新值": latest, "上期值": None, "变化pp": None,
                "建议得分": 0, "依据": "只有一期数据，无法算变化",
            })
            continue

        delta = latest - prev
        score, why = suggest_score(delta, direction, latest)
        total += weight * score
        rows.append({
            "类别": cat, "指标": name, "方向": direction, "权重": weight,
            "最新值": latest, "上期值": prev, "变化pp": round(delta, 2),
            "建议得分": score, "依据": why,
        })

    out = pd.DataFrame(rows)
    out["加权分"] = (out["权重"] * out["建议得分"]).round(3)

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(args.out, index=False, encoding="utf-8-sig")

    # ---- 控制台输出 ----
    pd.set_option("display.width", 200)
    pd.set_option("display.max_colwidth", 30)
    show = out[["指标", "方向", "权重", "最新值", "上期值", "变化pp", "建议得分", "依据"]]
    print(show.to_string(index=False, na_rep="—"))

    print("\n" + "=" * 70)
    print(f"景气分（建议）: {total:+.3f}")
    level = "景气上行" if total > 0.5 else ("景气下行" if total < -0.5 else "中性震荡")
    print(f"景气分级      : {level}")
    print("=" * 70)

    if missing:
        print("\n⚠️ 以下指标完全没有数据，已按 0 计入（会稀释景气分）：")
        for m in missing:
            print("   -", m)

    print(f"\n建议得分表已写出: {args.out}")
    print("下一步：在 scorecard.xlsx 的「评分卡」Sheet 里逐个确认得分，")
    print("        不同意的直接改，并在「打分依据」写清为什么改。")


if __name__ == "__main__":
    main()
