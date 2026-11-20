# -*- coding: utf-8 -*-
"""评估层：拿工具输出和人工标注对账，算准确率 / 召回率。"""

import os
import time

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

from tendercheck.extract import extract
from tendercheck.grade import grade
from tendercheck.parse import load_blocks
from tendercheck.rules import RULES

LABELS_PATH = "data/标注表.xlsx"
RAW_DIR = "data/raw"
OUT_PATH = "reports/评估结果.xlsx"

CODES = [f"C{i:02d}" for i in range(1, 13)]
NAME = {r["code"]: r["name"] for r in RULES}


def load_labels(path=LABELS_PATH):
    ws = load_workbook(path)["标注"]
    labels = {}
    for row in ws.iter_rows(min_row=2, max_col=8):
        vals = [c.value for c in row]
        if not vals[0]:
            continue
        fname = str(vals[0]).strip()
        code = str(vals[2]).strip()
        raw = "" if vals[4] is None else str(vals[4]).strip()
        labels.setdefault(fname, {})[code] = int(raw) if raw in ("0", "1") else None
    return labels


def predict(path):
    blocks = load_blocks(path)
    graded = grade(extract(blocks))
    return {g["code"]: (1 if g["hits"] else 0) for g in graded}


def classify(gt, pred):
    if gt == 1 and pred == 1:
        return "TP"
    if gt == 0 and pred == 1:
        return "FP"
    if gt == 1 and pred == 0:
        return "FN"
    return "TN"


def ratio(num, den):
    """安全除法：分母为 0 时返回 None。"""
    return num / den if den else None


def evaluate(labels):
    """返回 (逐类统计, 总耗时秒数)。"""
    per_class = {c: {"gt": 0, "pred": 0, "TP": 0, "FP": 0, "FN": 0, "TN": 0} for c in CODES}
    elapsed = 0.0
    for fname, gt_map in labels.items():
        path = os.path.join(RAW_DIR, fname)
        t0 = time.time()
        preds = predict(path)
        elapsed += time.time() - t0
        for code in CODES:
            gt = gt_map.get(code)
            if gt is None:
                continue
            pt = preds.get(code, 0)
            st = per_class[code]
            st["gt"] += gt
            st["pred"] += pt
            st[classify(gt, pt)] += 1
    return per_class, elapsed


def build_report(per_class, elapsed, n_files):
    wb = Workbook()
    ws = wb.active
    ws.title = "逐类指标"

    head_fill = PatternFill("solid", fgColor="1F4E79")
    headers = ["代码", "类别", "标注存在", "工具命中", "TP", "FP", "FN", "准确率", "召回率"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = head_fill
        c.font = Font(color="FFFFFF", bold=True)

    for row, code in enumerate(CODES, start=2):
        st = per_class[code]
        prec = ratio(st["TP"], st["TP"] + st["FP"])
        rec = ratio(st["TP"], st["TP"] + st["FN"])
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=NAME[code])
        ws.cell(row=row, column=3, value=st["gt"])
        ws.cell(row=row, column=4, value=st["pred"])
        ws.cell(row=row, column=5, value=st["TP"])
        ws.cell(row=row, column=6, value=st["FP"])
        ws.cell(row=row, column=7, value=st["FN"])
        c = ws.cell(row=row, column=8, value=prec)
        c.number_format = "0.0%"
        c = ws.cell(row=row, column=9, value=rec)
        c.number_format = "0.0%"

    for col, width in zip("ABCDEFGHI", [8, 18, 10, 10, 6, 6, 6, 10, 10]):
        ws.column_dimensions[col].width = width

    # 横向条形图：12 类的准确率
    chart = BarChart()
    chart.type = "bar"
    chart.title = "12 类条款准确率"
    chart.y_axis.title = "准确率"
    chart.x_axis.title = "条款"
    data = Reference(ws, min_col=8, min_row=1, max_row=13)
    cats = Reference(ws, min_col=1, min_row=2, max_row=13)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "K2")

    # 汇总页
    ws2 = wb.create_sheet("汇总")
    tp = sum(v["TP"] for v in per_class.values())
    fp = sum(v["FP"] for v in per_class.values())
    fn = sum(v["FN"] for v in per_class.values())
    tn = sum(v["TN"] for v in per_class.values())

    micro_p = ratio(tp, tp + fp)
    micro_r = ratio(tp, tp + fn)
    precs = [ratio(v["TP"], v["TP"] + v["FP"]) for v in per_class.values()]
    precs = [p for p in precs if p is not None]
    macro_p = sum(precs) / len(precs) if precs else None

    rows = [
        ("评估文件数", n_files),
        ("判断总数（文件 × 12 类）", n_files * 12),
        ("TP / FP / FN / TN", f"{tp} / {fp} / {fn} / {tn}"),
        ("微平均准确率 TP/(TP+FP)", micro_p),
        ("宏平均准确率（各类平均）", macro_p),
        ("召回率 TP/(TP+FN)", micro_r),
        ("假阳性率 FP/(TP+FP)", ratio(fp, tp + fp)),
        ("覆盖率（有定义的类）", f"{len(precs)}/12"),
        ("总耗时（秒）", round(elapsed, 2)),
        ("单份平均耗时（秒）", round(elapsed / n_files, 2) if n_files else None),
    ]
    ws2.cell(row=1, column=1, value="指标").font = Font(bold=True)
    ws2.cell(row=1, column=2, value="值").font = Font(bold=True)
    for i, (k, v) in enumerate(rows, start=2):
        ws2.cell(row=i, column=1, value=k)
        if isinstance(v, float) and v <= 1:
            c = ws2.cell(row=i, column=2, value=v)
            c.number_format = "0.0%"
        else:
            ws2.cell(row=i, column=2, value=v)
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 20

    os.makedirs("reports", exist_ok=True)
    wb.save(OUT_PATH)
    return {"tp": tp, "fp": fp, "fn": fn, "tn": tn,
            "micro_p": micro_p, "macro_p": macro_p, "recall": micro_r,
            "fpr": ratio(fp, tp + fp), "elapsed": elapsed}


if __name__ == "__main__":
    labels = load_labels()
    per_class, elapsed = evaluate(labels)
    summary = build_report(per_class, elapsed, len(labels))

    print(f"{'代码':<5}{'类别':<18}{'标注':<5}{'工具':<5}{'TP':<4}{'FP':<4}{'FN':<4}{'准确率':<8}{'召回率'}")
    for code in CODES:
        st = per_class[code]
        p = ratio(st["TP"], st["TP"] + st["FP"])
        r = ratio(st["TP"], st["TP"] + st["FN"])
        ps = f"{p:.1%}" if p is not None else "-"
        rs = f"{r:.1%}" if r is not None else "-"
        print(f"{code:<5}{NAME[code]:<18}{st['gt']:<5}{st['pred']:<5}{st['TP']:<4}{st['FP']:<4}{st['FN']:<4}{ps:<8}{rs}")

    print()
    print(f"微平均准确率: {summary['micro_p']:.1%}")
    print(f"宏平均准确率: {summary['macro_p']:.1%}")
    print(f"召回率:       {summary['recall']:.1%}")
    print(f"假阳性率:     {summary['fpr']:.1%}")
    print(f"总耗时:       {summary['elapsed']:.2f} 秒（{len(labels)} 份）")
    print()
    print("已生成:", OUT_PATH)