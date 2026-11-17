# -*- coding: utf-8 -*-
"""解析层：把 PDF 变成带页码的文本块。"""

import re

import pdfplumber

# 章节标题的正则：以「第」开头 + 中文数字 + 章/部/分/节
HEAD_RE = re.compile(r"^第[一二三四五六七八九十百]+[章部分节]")

def load_blocks(pdf_path):
    """读 PDF，返回文本块列表。

    每个块是一个 dict：
        {"page": 页码（从 1 开始）, "text": 这一行的文字, "type": "text"}
    """
    blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""

            # 通道 1：正文，按行切（先取出本页所有非空行）
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

            # 第一遍：本页有"几个像标题"的行（只看正则 + 前两条过滤）
            cands = [ln for ln in lines
                     if HEAD_RE.match(ln) and ".." not in ln and len(ln) >= 6]
            # 一页出现 3 个以上 → 判定这页是目录，整页标题都不算
            is_toc_page = len(cands) >= 3

            # 第二遍：逐行判断并收集
            for line in lines:
                is_head = (
                        not is_toc_page
                        and HEAD_RE.match(line) is not None
                        and ".." not in line
                        and len(line) >= 6
                )
                blocks.append({
                    "page": pno,
                    "text": line,
                    "type": "text",
                    "is_heading": is_head,
                })
            # 通道 2：表格，按行拼成 "a | b | c"
            for tb in page.extract_tables():
                for row in tb:
                    cells = [(c or "").replace("\n", " ").strip() for c in row]
                    line = " | ".join(cells)
                    if line.strip(" |"):
                        blocks.append({"page": pno, "text": line, "type": "table","is_heading": False,})
    return blocks


if __name__ == "__main__":
    import sys

    path = sys.argv[1]
    blocks = load_blocks(path)
    print("总块数:", len(blocks))
    print("覆盖页数:", max(b["page"] for b in blocks))
    print("表格块:", len([b for b in blocks if b["type"] == "table"]))
    print("标题块:", len([b for b in blocks if b.get("is_heading")]))
    total_chars = sum(len(b["text"]) for b in blocks)
    pages = max(b["page"] for b in blocks)
    print("平均每页字符:", total_chars // pages)
    print("需OCR:", total_chars / pages < 50)
    for b in blocks[:5]:
        print(b)