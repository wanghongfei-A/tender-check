# -*- coding: utf-8 -*-
"""解析层：把 PDF / DOCX 变成「带页码的文本块」。

两种格式产出**同一种块结构**，所以上层（extract / grade / report）完全不用区分：
    {"page": 页码（从 1 开始）, "text": 一段文字, "type": "text"|"table",
     "is_heading": 是否章节标题, "page_src": "pdf"|"break"|"estimate"}

PDF 和 DOCX 的本质差异：
- PDF 每一页真实存在，逐页抽取即可；
- DOCX 没有"页"的概念（python-docx 只提供段落/表格流），
  页码要从 XML 里的分页标记还原：
    1) w:lastRenderedPageBreak —— Word 保存时写入的"上次渲染分页位置"
    2) <w:br w:type="page"/>   —— 手动分页符
  一个标记都找不到时，按字符数估算（每页约 3000 字），并标记 page_src="estimate"。
"""

import os
import re

import pdfplumber
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# 章节标题的正则：以「第」开头 + 中文数字 + 章/部/分/节
HEAD_RE = re.compile(r"^第[一二三四五六七八九十百]+[章部分节]")

# DOCX 无分页标记时的兜底：按每页多少字符估算
DOCX_CHARS_PER_PAGE = 1000

# 目录行：标题 + Tab + 页码，如「第五章 评标方法及评标标准\t58」
TOC_LINE_RE = re.compile(r"^第[一二三四五六七八九十百]+[章部分节].*\t\s*\d+\s*$")


def _is_heading(line, is_toc_page=False):
    """判断一行是不是章节标题（含三条上下文过滤）。"""
    return (
            not is_toc_page
            and HEAD_RE.match(line) is not None
            and ".." not in line      # 目录行：标题 + 省略号 + 页码
            and not TOC_LINE_RE.match(line)   # 目录行：标题 + Tab + 页码（DOCX 常见）
            and len(line) >= 6        # 太短的多半是页眉残留
    )


def _pdf_blocks(pdf_path):
    """读 PDF：正文通道 + 表格通道，逐页产出块。"""
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
                blocks.append({
                    "page": pno,
                    "text": line,
                    "type": "text",
                    "is_heading": _is_heading(line, is_toc_page),
                    "page_src": "pdf",
                })

            # 通道 2：表格，按行拼成 "a | b | c"
            for tb in page.extract_tables():
                for row in tb:
                    cells = [(c or "").replace("\n", " ").strip() for c in row]
                    line = " | ".join(cells)
                    if line.strip(" |"):
                        blocks.append({"page": pno, "text": line, "type": "table",
                                       "is_heading": False, "page_src": "pdf"})
    return blocks


def _assign_pages_by_chars(blocks, chars_per_page):
    """按累计字符数给块分配页码（没有目录可校准时用）。"""
    cum = 0
    for b in blocks:
        b["page"] = 1 + int(cum // chars_per_page)
        b["page_src"] = "estimate"
        cum += len(b["text"])


def _assign_pages_by_toc(blocks, anchors, tail_chars_per_page):
    """用目录页的「章标题 → 真实页码」做分段线性校准。

    anchors = [(块下标, 目录里该章的真实页码), ...]，按块下标升序。
    同一章内部按字符数在「上一章页码 → 下一章页码」之间线性插值。
    """
    n = len(blocks)
    pref = [0] * (n + 1)
    for i, b in enumerate(blocks):
        pref[i + 1] = pref[i] + len(b["text"])

    for i, b in enumerate(blocks):
        # 找 i 之前最近的锚点
        k = -1
        for j, (idx, _pg) in enumerate(anchors):
            if idx <= i:
                k = j
            else:
                break

        if k < 0:
            # 第一个章标题之前（封面、目录页本身）
            b["page"] = 1
        elif k == len(anchors) - 1:
            # 最后一章之后：按平均字符数往后推
            idx0, page0 = anchors[k]
            b["page"] = page0 + int((pref[i] - pref[idx0]) // tail_chars_per_page)
        else:
            idx0, page0 = anchors[k]
            idx1, page1 = anchors[k + 1]
            span = max(pref[idx1] - pref[idx0], 1)
            b["page"] = page0 + int((pref[i] - pref[idx0]) * (page1 - page0) / span)

        b["page_src"] = "toc"


def _docx_blocks(docx_path):
    """读 DOCX：按文档真实顺序遍历「段落 / 表格」，从分页标记还原页码。

    页码策略（三级降级）：
    1. 分页标记（lastRenderedPageBreak / 手动分页符）→ page_src="break"
    2. 分页标记太少 → 用目录页的「章 → 页码」做分段线性校准 → page_src="toc"
    3. 连目录都没有 → 按 DOCX_CHARS_PER_PAGE 粗估 → page_src="estimate"
    """
    doc = Document(docx_path)
    blocks = []
    page = 1
    breaks = 0

    for child in doc.element.body.iterchildren():
        tag = child.tag
        if tag.endswith("}p"):
            # ---- 段落 ----
            para = Paragraph(child, doc)
            xml = child.xml
            # 本段内部有几个分页标记（渲染分页 + 手动分页）
            n_break = xml.count("lastRenderedPageBreak") + xml.count('w:type="page"')
            text = para.text.strip()
            if text:
                blocks.append({
                    "page": page,
                    "text": text,
                    "type": "text",
                    "is_heading": _is_heading(text),
                    "page_src": "break",
                })
            page += n_break
            breaks += n_break

        elif tag.endswith("}tbl"):
            # ---- 表格：逐行拼成 "a | b | c"，一个表格算一个块 ----
            table = Table(child, doc)
            lines = []
            for row in table.rows:
                cells = [(c.text or "").replace("\n", " ").strip() for c in row.cells]
                line = " | ".join(cells)
                if line.strip(" |"):
                    lines.append(line)
            if lines:
                blocks.append({"page": page, "text": "\n".join(lines), "type": "table",
                               "is_heading": False, "page_src": "break"})

    if not blocks:
        return blocks

    total_chars = sum(len(b["text"]) for b in blocks)
    marker_pages = breaks + 1

    # 分页标记够多（平均每页字符数合理）→ 直接信它
    if breaks and total_chars / marker_pages <= 3000:
        return blocks

    # 收集目录锚点：目录行「第X章 ... \t页码」
    # 注意：目录行的位置在文首，但页码指的是**正文里那一章**的位置，
    # 所以要在正文中重新找到同名标题，用它当锚点。
    toc_entries = []       # [(章标题, 目录页码)]
    last_toc_idx = -1
    for i, b in enumerate(blocks):
        if TOC_LINE_RE.match(b["text"]):
            title = b["text"].rsplit("\t", 1)[0].strip()
            toc_entries.append((title, int(b["text"].rsplit("\t", 1)[-1].strip())))
            b["is_heading"] = False        # 目录行不是正文标题
            last_toc_idx = i

    def norm(s):
        return re.sub(r"\s+", "", s)

    anchors = []
    search_from = last_toc_idx + 1
    for title, page_no in toc_entries:
        for i in range(search_from, len(blocks)):
            if norm(blocks[i]["text"]) == norm(title):
                anchors.append((i, page_no))
                search_from = i + 1
                break

    if len(anchors) >= 2 and anchors[-1][1] > 1:
        tail = total_chars / max(anchors[-1][1], 1)
        _assign_pages_by_toc(blocks, anchors, tail)
    else:
        _assign_pages_by_chars(blocks, DOCX_CHARS_PER_PAGE)

    return blocks


def load_blocks(path):
    """按扩展名分派：.pdf → PDF 通道，.docx → DOCX 通道。

    返回的块结构完全相同，上层无需区分格式。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return _pdf_blocks(path)
    if ext == ".docx":
        return _docx_blocks(path)
    raise ValueError("不支持的格式：%s（当前支持 .pdf / .docx）" % ext)


if __name__ == "__main__":
    import sys

    path = sys.argv[1]
    blocks = load_blocks(path)
    pages = max(b["page"] for b in blocks)
    total_chars = sum(len(b["text"]) for b in blocks)
    src = blocks[0].get("page_src", "?")

    print("文件:", os.path.basename(path))
    print("总块数:", len(blocks))
    print("覆盖页数:", pages, "（页码来源：%s）" % src)
    print("表格块:", len([b for b in blocks if b["type"] == "table"]))
    print("标题块:", len([b for b in blocks if b.get("is_heading")]))
    print("平均每页字符:", total_chars // pages)
    print("需OCR:", total_chars / pages < 50)
    print()
    print("前 3 个标题：")
    for b in [x for x in blocks if x.get("is_heading")][:3]:
        print("   p%-4d %s" % (b["page"], b["text"][:50]))
