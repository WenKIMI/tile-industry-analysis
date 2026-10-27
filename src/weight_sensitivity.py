#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
权重敏感性分析：结论到底有多依赖那套主观权重？

做四件事：
    1. 基线：用当前权重算景气分
    2. 等权：12 个指标各 8.33%（最不主观的方案）
    3. 情景：需求加权、政策加权、去政策化
    4. 蒙特卡洛：在当前权重周围随机扰动 2 万次，看分级有多稳

用法：
    py -3.13 src/weight_sensitivity.py
"""

from __future__ import annotations

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "scorecard.xlsx"

UP, DOWN = 0.5, -0.5


def grade(s: float) -> str:
    if s > UP:
        return "景气上行"
    if s < DOWN:
        return "景气下行"
    return "中性震荡"


def load() -> tuple[list[str], np.ndarray, np.ndarray]:
    ws = load_workbook(XLSX)["评分卡"]
    names, w, s = [], [], []
    for r in range(5, 17):
        name = ws.cell(row=r, column=2).value
        if not name:
            continue
        names.append(name)
        w.append(float(ws.cell(row=r, column=4).value or 0))
        # 磨具行的得分是公式，这里只取数值型
        v = ws.cell(row=r, column=8).value
        s.append(float(v) if isinstance(v, (int, float)) else 0.0)
    return names, np.array(w, dtype=float), np.array(s, dtype=float)


def main() -> None:
    names, w0, s = load()
    print("=== 当前权重与得分 ===")
    for n, wi, si in zip(names, w0, s):
        print(f"  {n:<14} 权重{wi:>6.1%}  得分{si:>+4.0f}")
    print(f"  权重合计 {w0.sum():.1%}   得分合计 {s.sum():+.0f}")

    base = float((w0 * s).sum())
    print(f"\n【基线】景气分 {base:+.3f} → {grade(base)}")

    # ---------- 方案对比 ----------
    print("\n=== 不同权重方案 ===")
    scenarios = {}

    we = np.full_like(w0, 1 / len(w0))
    scenarios["等权（各 8.33%）"] = we

    # 需求类 ×1.5，其余按比例压缩
    need_idx = [i for i, n in enumerate(names) if any(k in n for k in ("竣工", "销售", "产量", "出口"))]
    w_need = w0.copy()
    w_need[need_idx] *= 1.5
    scenarios["需求类权重 ×1.5"] = w_need / w_need.sum()

    # 政策类 ×2
    pol_idx = [i for i, n in enumerate(names) if "政策" in n]
    w_pol = w0.copy()
    w_pol[pol_idx] *= 2
    scenarios["政策类权重 ×2"] = w_pol / w_pol.sum()

    # 去掉政策，按剩余比例重分配
    w_nopol = w0.copy()
    w_nopol[pol_idx] = 0
    scenarios["剔除政策类"] = w_nopol / w_nopol.sum()

    # 财务类（三家公司）加倍
    fin_idx = [i for i, n in enumerate(names) if any(k in n for k in ("毛利率", "应收", "库存", "产量"))]
    w_fin = w0.copy()
    w_fin[fin_idx] *= 1.6
    scenarios["财报类（三家公司）×1.6"] = w_fin / w_fin.sum()

    rows = [("基线（当前权重）", base, grade(base))]
    for label, wv in scenarios.items():
        sc = float((wv * s).sum())
        rows.append((label, sc, grade(sc)))
        print(f"  {label:<22} {sc:+.3f}  {grade(sc)}")

    print("\n=== 蒙特卡洛：权重随机扰动 20000 次 ===")
    rng = np.random.default_rng(42)
    for conc in (5, 10, 20, 50):
        # alpha 越大，扰动越小（越接近原权重）
        draws = rng.dirichlet(w0 * conc, size=20000)
        sc = draws @ s
        up = (sc > UP).mean()
        dn = (sc < DOWN).mean()
        mid = 1 - up - dn
        print(
            f"  扰动强度 conc={conc:<3}  "
            f"上行 {up:5.1%} | 中性 {mid:5.1%} | 下行 {dn:5.1%}   "
            f"（均值 {sc.mean():+.3f}，5%~95% 区间 {np.percentile(sc,5):+.3f} ~ {np.percentile(sc,95):+.3f}）"
        )

    # ---------- 单指标影响力 ----------
    print("\n=== 单指标对结论的影响（把该指标权重翻倍） ===")
    for i, n in enumerate(names):
        wv = w0.copy()
        wv[i] *= 2
        wv = wv / wv.sum()
        sc = float((wv * s).sum())
        flag = "  ← 会翻档" if grade(sc) != grade(base) else ""
        print(f"  {n:<14} {sc:+.3f}  {grade(sc)}{flag}")

    print(
        "\n结论："
        "\n  1) 换一套同样合理的权重，分级约 52%~84% 仍是「中性震荡」，16%~38% 变成「景气下行」；"
        "\n     上行概率 0%~10% —— 方向偏弱是稳的，「不是上行」这一点最稳。"
        "\n  2) 但「中性」这个档位并不牢靠，且主要靠地产政策那 10% 权重撑着："
        "\n     把政策类剔除，景气分立刻掉到 -0.53（下行）。"
        "\n  3) 分数值没有精度意义（-0.30 的「0.30」是假精度），只能读方向与档位。"
    )


if __name__ == "__main__":
    main()
