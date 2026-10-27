#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
按人工判断调整 scorecard.xlsx 的得分，并用 Python 独立复算全部结果。

为什么要有这个脚本：
    .xlsx 里的公式只有 Excel 打开时才算。我们没有 Excel 引擎，
    所以要自己把「加权分 / 景气分 / 磨具两分」再算一遍，
    确保写进文件的东西和预期一致，也让你在没打开 Excel 前就知道结果。
"""

import re
import shutil
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "scorecard.xlsx"

# 人工调整：行号 -> (新得分, 新依据)
OVERRIDES = {
    14: (1, "天数 73.97→58.39 降 15.6 天，但同期营收同比 -15.1%，分母缩水效应显著，"
            "不能直接视为回款改善，由 +2 下调至 +1（人工判断 2026-10-27）"),
    # 政策强度是「水平型」指标，不适用「变化法」：
    # 只有一期数据，score.py 会判为「无法算变化」给 0 分，这里按绝对水平覆盖。
    15: (2, "水平型指标，不适用变化法。Q3 政策强度 3/3 分（8·28全国性文件+一线放松），"
            "按「3分→+2」映射（人工判断 2026-10-27）"),
    # 价格只知道定性档位（1/2 档），不适用变化法，按档位直接映射：
    # 0档→-1，1档→+1，2档→+2
    11: (1, "无公开价格指数，按「定性档位法」：1档（涨价函广但落地存疑）→ +1。"
            "依据：8.29-9.5 七产区集中涨价函，但属煤价成本推动且终端难落地（人工判断 2026-10-27）"),
    # 出口：8月单月 -4.32% vs 1—8月累计 -7.72%，降幅收窄，属边际改善，但缺上月单月值，按定性 +1
    8: (1, "8月单月出口额同比 -4.32%，明显好于 1—8月累计的 -7.72%，降幅收窄即边际改善；"
           "因缺 7 月单月值无法算严格变化，按定性 +1 处理（人工判断 2026-10-27）"),
    # 开工率：定性档位 2（历史低位）→ -2
    9: (-2, "定性档位 2/2（历史低位）：江西高安约7成、上高宜丰8-9成，但国内其它产区普遍不足50%，"
            "行业开窑率处于历史低位（人工判断 2026-10-27）"),
    # 环保政策：定性档位 1（温和收紧，负向）→ -1
    16: (-1, "负向指标，定性档位 1/2：建筑卫生陶瓷纳入2026年度工业节能监察重点，"
             "多地环保管控叠加订单不足致停窑减产，但未见主产区大范围停限产（人工判断 2026-10-27）"),
}


def main() -> None:
    wb = load_workbook(XLSX)
    ws = wb["评分卡"]

    for row, (score, reason) in OVERRIDES.items():
        old = ws.cell(row=row, column=8).value
        ws.cell(row=row, column=8, value=score)
        ws.cell(row=row, column=10, value=reason)
        print(f"  第{row}行 {ws.cell(row=row, column=2).value}: {old} → {score}")

    # ---- 独立复算（不依赖 Excel）----
    total, detail = 0.0, []
    for r in range(5, 17):
        name = ws.cell(row=r, column=2).value
        w = ws.cell(row=r, column=4).value or 0
        s = ws.cell(row=r, column=8).value
        if s is None:
            detail.append((name, w, None, None))
            continue
        sub = w * s
        total += sub
        detail.append((name, w, s, sub))

    print("\n=== 评分卡复算 ===")
    for name, w, s, sub in detail:
        if s is None:
            print(f"  {name:<14} 权重{w:>5.0%}  得分 —     （无数据，不计入）")
        else:
            print(f"  {name:<14} 权重{w:>5.0%}  得分{s:>+3d}   加权{sub:>+.3f}")

    level = "景气上行" if total > 0.5 else ("景气下行" if total < -0.5 else "中性震荡")
    print(f"\n  景气分 = {total:+.3f}    分级 = {level}")

    covered = sum(w for _, w, s, _ in detail if s is not None)
    print(f"  权重覆盖率 = {covered:.0%}")

    # ---- 磨具两分 ----
    m = wb["磨具代理"]
    agg = {"磨边/抛光耗材": 0.0, "压机模具": 0.0}
    cover = {"磨边/抛光耗材": 0.0, "压机模具": 0.0}

    def resolve(v):
        """磨具代理的得分列是 =评分卡!H7 这类公式，自行解析取值。"""
        if isinstance(v, str) and v.startswith("="):
            mt = re.search(r"评分卡!H(\d+)", v)
            return ws.cell(row=int(mt.group(1)), column=8).value if mt else None
        return v

    for r in range(5, 12):
        belong = m.cell(row=r, column=1).value
        w = m.cell(row=r, column=3).value or 0
        s = resolve(m.cell(row=r, column=4).value)
        if s is not None and belong in agg:
            agg[belong] += w * s
            cover[belong] += w

    print("\n=== 磨具代理分复算 ===")
    for k in agg:
        flag = "（权重不全，仅供参考）" if cover[k] < 0.99 else ""
        print(f"  {k} 代理分 = {agg[k]:+.3f}   已覆盖权重 {cover[k]:.0%} {flag}")

    backup = XLSX.with_name(f"{XLSX.stem}_备份.xlsx")
    shutil.copy2(XLSX, backup)
    wb.save(XLSX)
    print(f"\n已保存 {XLSX.name}（备份 {backup.name}）")


if __name__ == "__main__":
    main()
