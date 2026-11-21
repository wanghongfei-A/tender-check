# -*- coding: utf-8 -*-
"""抽取层：在文本块里找 12 类条款。

思路是「章节定位 + 关键词窗口 + 置信度分级」：
- 高：命中关键词，且该块所处章节正是这类条款该出现的章节
- 中：命中关键词，但章节对不上（可能误伤）
- 低：命中在目录行（标题 + 省略号 + 页码），可信度存疑
"""

import re

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


# 「公告类」章节的常见叫法：招标公告 / 投标邀请 / 磋商邀请函 / 询价公告 …
ANNOUNCE_WORDS = ["招标公告", "投标邀请", "采购公告", "招标邀请",
                  "磋商邀请", "谈判邀请", "询价公告", "比选公告"]


def detect_announcement(blocks, threshold=100):
    """检查「公告类」章节是否存在、内容是否过少。

    部分电子招标文件（尤其建设工程类）不在文件内重复招标公告，只在
    公告章节留一行「公告见 XX 平台」。此时该章节的正文极少（几十个字），
    而关键信息（截止时间等）只在平台上——工具应明确提示，而不是猜。

    返回 None（没找到公告章节）或 dict：
        {page, heading, chars, empty, sample}
    """
    for idx, b in enumerate(blocks):
        if b.get("is_heading") and any(w in b["text"] for w in ANNOUNCE_WORDS):
            seg = []
            for nxt in blocks[idx + 1:]:
                if nxt.get("is_heading"):
                    break
                seg.append(nxt)
            chars = sum(len(x["text"]) for x in seg)
            return {
                "page": b["page"],
                "heading": b["text"][:30],
                "chars": chars,
                "empty": chars < threshold,
                "sample": " ".join(x["text"] for x in seg)[:60],
            }
    return None


# 截止时间必须带「时刻」：2025年09月19日 08时30分 / 2026-11-17 09:00
# 只有日期（如 2024年6月1日）的多半是业绩区间或有效期，不能当截止时间
STRICT_DT_RE = re.compile(
    r"\d{4}\s*[-年/]\s*\d{1,2}\s*[-月/]\s*\d{1,2}\s*日?"
    r"[^\d\n]{0,10}?"
    r"\d{1,2}\s*[:：时点]\s*\d{1,2}\s*分?"
)
# 出现这些词的块，附近很可能有截止/开标时间
DEADLINE_KEY_RE = re.compile(
    r"截止时间|开标时间|投标截止|递交截止|响应文件截止|递交投标文件|投标文件递交|开标时间及地点")


def find_deadline(blocks, span=3):
    """在「截止/开标」附近的块里找带时刻的截止时间（时间值常与关键词不在同一块）。

    返回 {"page", "when", "source"} 或 None。
    """
    for i, b in enumerate(blocks):
        if not DEADLINE_KEY_RE.search(b["text"]):
            continue
        m = STRICT_DT_RE.search(b["text"])
        if m:
            return {"page": b["page"], "when": m.group(0).strip(), "source": b["text"][:60]}
        for nxt in blocks[i + 1:i + 1 + span]:
            m = STRICT_DT_RE.search(nxt["text"])
            if m:
                return {"page": nxt["page"], "when": m.group(0).strip(),
                        "source": nxt["text"][:60]}
    return None


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
