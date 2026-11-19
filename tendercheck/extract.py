# -*- coding: utf-8 -*-
"""抽取层：在文本块里找 12 类条款。

思路是「章节定位 + 关键词窗口 + 置信度分级」：
- 高：命中关键词，且该块所处章节正是这类条款该出现的章节
- 中：命中关键词，但章节对不上（可能误伤）
- 低：命中在目录行（标题 + 省略号 + 页码），可信度存疑
"""

from tendercheck.rules import RULES, SECTION_ALIASES


def section_tags(section_text):
    """把章节标题映射成代号集合。

    '第二章 投标人须知' -> {'须知'}
    '第一章 招标公告'   -> {'公告'}
    """
    tags = set()
    for tag, aliases in SECTION_ALIASES.items():
        for a in aliases:
            if a and a in section_text:
                tags.add(tag)
    return tags


def attach_sections(blocks):
    """给每个块打上「所属章节」：遇到标题就切换当前章节。"""
    current = ""
    for b in blocks:
        if b.get("is_heading"):
            current = b["text"]
        b["section"] = current
    return blocks


def judge_confidence(rule, block):
    """单条命中的置信度：高 / 中 / 低。"""
    text = block["text"]
    if ".." in text:
        return "低"
    tags = section_tags(block.get("section", ""))
    if tags & set(rule["sections"]):
        return "高"
    return "中"


def extract(blocks):
    """对每个块跑 12 类规则，返回按类聚合的结果。"""
    attach_sections(blocks)
    order = {"高": 0, "中": 1, "低": 2}
    results = []
    for rule in RULES:
        hits = []
        for b in blocks:
            matched = [k for k in rule["keywords"] if k in b["text"]]
            if not matched:
                continue
            hits.append({
                "page": b["page"],
                "text": b["text"],
                "section": b.get("section", ""),
                "keyword": matched[0],
                "confidence": judge_confidence(rule, b),
            })
        hits.sort(key=lambda h: (order[h["confidence"]], h["page"]))
        results.append({
            "code": rule["code"],
            "name": rule["name"],
            "risk": rule["risk"],
            "confidence": hits[0]["confidence"] if hits else "无",
            "hits": hits,
        })
    return results


if __name__ == "__main__":
    import sys

    from tendercheck.parse import load_blocks

    path = sys.argv[1]
    blocks = load_blocks(path)
    results = extract(blocks)
    print("%-5s %-16s %-4s %-5s %s" % ("代码", "条款", "置信", "命中数", "首条（页 + 原文）"))
    for r in results:
        first = r["hits"][0] if r["hits"] else None
        brief = ("p%d %s" % (first["page"], first["text"][:30])) if first else "-"
        print("%-5s %-16s %-4s %-5d %s"
              % (r["code"], r["name"], r["confidence"], len(r["hits"]), brief))
