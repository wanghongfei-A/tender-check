# -*- coding: utf-8 -*-
"""分级层：把抽取结果变成红/黄/绿。

第 3 步：加文本信号（废标类升级 / 免收类降级）+ 未检出按风险分类。
"""

import sys

from tendercheck.extract import extract
from tendercheck.parse import load_blocks

# 字典：键是条款代码，值是基础等级
BASE_LEVEL = {
    "C04": "红", "C09": "红", "C10": "红",
    "C02": "黄", "C03": "黄", "C05": "黄",
    "C06": "黄", "C07": "黄", "C11": "黄", "C12": "黄",
    "C01": "绿", "C08": "绿",
}

# 文本信号：命中句里出现这些词，说明条款的"后果"很重
WASTE_WORDS = ["废标", "无效", "不予受理", "拒绝", "否决"]
# 文本信号：明确写了"不用管"
FREE_WORDS = ["免收", "不收取", "不设"]


def adjust(level, text):
    """按命中文本微调等级。返回 (新等级, 说明或 None)。"""
    if any(w in text for w in WASTE_WORDS):
        return "红", "命中文本含废标类表述，升级为红"
    if any(w in text for w in FREE_WORDS):
        if level == "红":
            return "黄", "命中文本明确免收/不设，降为黄"
        if level == "黄":
            return "绿", "命中文本明确免收/不设，降为绿"
    return level, None


def grade(results):
    """给每一类结果打等级，返回新的列表。"""
    graded = []
    for r in results:
        if not r["hits"]:
            level = "待补"
            if BASE_LEVEL[r["code"]] == "红":
                reason = "未检出，但属高风险条款，必须人工确认"
            else:
                reason = "全文未出现该类关键词，倾向文件无此要求，建议抽查"
        elif r["confidence"] in ("中", "低"):
            level = "黄"
            reason = "命中但章节不对或疑似目录行，需复核"
        else:
            level = BASE_LEVEL[r["code"]]
            reason = "命中且章节吻合"
            level, note = adjust(level, r["hits"][0]["text"])
            if note:
                reason = note

        item = dict(r)
        item["level"] = level
        item["reason"] = reason
        graded.append(item)
    return graded


if __name__ == "__main__":
    path = sys.argv[1]
    blocks = load_blocks(path)
    results = extract(blocks)
    graded = grade(results)

    for g in graded:
        print(f"{g['code']}  {g['name']}  {g['level']}  {g['reason']}")

    counts = {"红": 0, "黄": 0, "绿": 0, "待补": 0}
    for g in graded:
        counts[g["level"]] = counts[g["level"]] + 1
    print("-" * 60)
    print(f"统计：红 {counts['红']}  黄 {counts['黄']}  绿 {counts['绿']}  待补 {counts['待补']}")