#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
把 data/evidence/*.json 合并成一张证据总表，并做基础校验。

用途：
    1. 验证每份 JSON 都能解析、字段齐全（LLM 输出常有缺字段、多逗号）
    2. 生成 CSV，方便排序、筛选、做透视
    3. 打印分环节 / 分方向的统计，快速看出证据是否失衡

用法：
    py -3.13 src/merge_evidence.py
    py -3.13 src/merge_evidence.py --dir data/evidence --out data/证据汇总.csv
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

REQUIRED = ["事实", "类型", "影响环节", "传导到磨具", "方向", "强度", "可信度", "原文摘录"]
VALID = {
    "类型": {"事实", "观点", "推断"},
    "影响环节": {"需求", "供给", "价格", "成本", "盈利", "信用", "政策"},
    "传导到磨具": {"磨边耗材", "压机模具", "无"},
    "方向": {"利多", "利空", "中性"},
    "强度": {"高", "中", "低"},
    "可信度": {"高", "中", "低"},
}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=str(ROOT / "data" / "evidence"))
    ap.add_argument("--out", default=str(ROOT / "data" / "证据汇总.csv"))
    args = ap.parse_args()

    d = Path(args.dir)
    files = sorted(d.glob("*.json"))
    if not files:
        raise SystemExit(f"{d} 里没有 JSON 文件")

    rows, problems = [], []
    for fp in files:
        try:
            obj = json.loads(fp.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            problems.append(f"{fp.name}: JSON 解析失败 — {exc}")
            continue

        for i, ev in enumerate(obj.get("证据", []), start=1):
            miss = [k for k in REQUIRED if not ev.get(k)]
            if miss:
                problems.append(f"{fp.name} 第{i}条: 缺字段 {miss}")

            for field, allowed in VALID.items():
                val = ev.get(field)
                if val and val not in allowed:
                    problems.append(f"{fp.name} 第{i}条: {field}='{val}' 不在 {allowed} 内")

            rows.append(
                {
                    "来源文档": fp.stem,
                    "来源": obj.get("来源", ""),
                    "文档日期": obj.get("日期", ""),
                    "序号": i,
                    "事实": ev.get("事实", ""),
                    "类型": ev.get("类型", ""),
                    "影响环节": ev.get("影响环节", ""),
                    "传导到磨具": ev.get("传导到磨具", ""),
                    "方向": ev.get("方向", ""),
                    "强度": ev.get("强度", ""),
                    "可信度": ev.get("可信度", ""),
                    "原文摘录": ev.get("原文摘录", ""),
                }
            )

    df = pd.DataFrame(rows)
    df.to_csv(args.out, index=False, encoding="utf-8-sig")

    print(f"读取 {len(files)} 份文档，共 {len(df)} 条证据")
    print(f"已写出 {args.out}\n")

    print("=== 按影响环节 ===")
    print(df["影响环节"].value_counts().to_string())
    print("\n=== 按方向 ===")
    print(df["方向"].value_counts().to_string())
    print("\n=== 按类型（事实/观点/推断）===")
    print(df["类型"].value_counts().to_string())
    print("\n=== 传导到磨具 ===")
    print(df["传导到磨具"].value_counts().to_string())

    print(f"\n=== 校验结果 ===")
    if problems:
        print(f"发现 {len(problems)} 个问题：")
        for p in problems:
            print("  -", p)
    else:
        print("字段齐全、枚举值合法、JSON 可解析。")

    # 方向失衡提示：这是研究里最常见的问题
    n_up = (df["方向"] == "利多").sum()
    n_dn = (df["方向"] == "利空").sum()
    if n_up and n_dn:
        ratio = max(n_up, n_dn) / min(n_up, n_dn)
        if ratio >= 2:
            print(
                f"\n⚠️ 方向失衡：利多 {n_up} 条 / 利空 {n_dn} 条，比例 {ratio:.1f}:1。"
                f"\n   多半不是行业真的这么一边倒，而是取材时的确认偏误。"
                f"\n   建议按 prompts/04_red_team.md 补一轮反方证据。"
            )


if __name__ == "__main__":
    main()
