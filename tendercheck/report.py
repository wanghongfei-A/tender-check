# -*- coding: utf-8 -*-
"""报告层：把分级结果写成 Excel（4 个 sheet）。

- 概览：项目基本信息 + 风险统计
- 风险清单：红/黄且高置信度的原文（每类最多 3 条）
- 全部条款：12 类各一行
- 待复核：中低置信度命中 + 未检出的类
"""

import os
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from tendercheck.extract import extract
from tendercheck.grade import grade
from tendercheck.parse import load_blocks


def build_summary(graded, path, pages):
    """概览：键值对。"""
    counts = {"红": 0, "黄": 0, "绿": 0, "待补": 0}
    for g in graded:
        counts[g["level"]] += 1

    def first_text(code):
        for g in graded:
            if g["code"] == code and g["hits"]:
                return g["hits"][0]["text"][:60]
        return "（未检出）"

    rows = [
        ("文件名", os.path.basename(path)),
        ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("总页数", pages),
        ("红（必须处理）", counts["红"]),
        ("黄（需确认）", counts["黄"]),
        ("绿（信息类）", counts["绿"]),
        ("待补（需人工确认）", counts["待补"]),
        ("项目名称", first_text("C01")),
        ("预算/最高限价", first_text("C02")),
        ("投标截止/开标", first_text("C04")),
    ]
    return pd.DataFrame(rows, columns=["项", "内容"])


def build_risk_rows(graded, limit=3):
    """风险清单：只放红/黄，每类最多 limit 条，且同一页只留一条（避免同一要求的正文/表格重复）。"""
    rows = []
    for g in graded:
        if g["level"] not in ("红", "黄"):
            continue
        seen_pages = set()
        for h in g["hits"]:
            if h["page"] in seen_pages:
                continue
            seen_pages.add(h["page"])
            rows.append({
                "等级": g["level"],
                "类别": f'{g["code"]} {g["name"]}',
                "页码": h["page"],
                "原文": h["text"][:150],
                "复核建议": g.get("risk", ""),
            })
            if len(seen_pages) >= limit:
                break
    return pd.DataFrame(rows)


def build_all_sheet(graded):
    """全部条款：12 类各一行。"""
    rows = []
    for g in graded:
        first = g["hits"][0] if g["hits"] else None
        rows.append({
            "代码": g["code"],
            "类别": g["name"],
            "等级": g["level"],
            "置信度": g["confidence"],
            "命中数": len(g["hits"]),
            "页码": first["page"] if first else "",
            "原文": first["text"][:120] if first else "",
            "说明": g["reason"],
        })
    return pd.DataFrame(rows)


def build_review_rows(graded):
    """待复核：只列「整类不确定」的，不列个别命中。

    - 待补：工具完全没抽到，需人工确认文件里有没有
    - 中/低：整类最高置信度都不高，说明这一类整体可疑
    """
    rows = []
    for g in graded:
        if g["level"] == "待补":
            rows.append({
                "类别": f'{g["code"]} {g["name"]}',
                "原因": g["reason"],
                "页码": "",
                "原文": "",
            })
        elif g["confidence"] in ("中", "低"):
            for h in g["hits"][:2]:
                rows.append({
                    "类别": f'{g["code"]} {g["name"]}',
                    "原因": f'整类置信度{h["confidence"]}（章节：{h["section"] or "无"}）',
                    "页码": h["page"],
                    "原文": h["text"][:150],
                })
    return pd.DataFrame(rows)


def fix_widths(path):
    """把每列宽度调到能看清内容，并冻结表头。"""
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for col in ws.columns:
            longest = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(longest * 1.8, 8), 60)
        ws.freeze_panes = "A2"
    wb.save(path)


if __name__ == "__main__":
    path = sys.argv[1]
    blocks = load_blocks(path)
    graded = grade(extract(blocks))
    pages = max(b["page"] for b in blocks)

    df_sum = build_summary(graded, path, pages)
    df_risk = build_risk_rows(graded)
    df_all = build_all_sheet(graded)
    df_rev = build_review_rows(graded)

    os.makedirs("reports", exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join("reports", f"{stem}_风险自查报告.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_sum.to_excel(writer, sheet_name="概览", index=False)
        df_risk.to_excel(writer, sheet_name="风险清单", index=False)
        df_all.to_excel(writer, sheet_name="全部条款", index=False)
        df_rev.to_excel(writer, sheet_name="待复核", index=False)

    fix_widths(out)

    print("已生成:", out)
    print()
    print("概览：")
    print(df_sum.to_string(index=False, header=False))
    print()
    print(f"风险清单 {len(df_risk)} 行 / 全部条款 {len(df_all)} 行 / 待复核 {len(df_rev)} 行")
