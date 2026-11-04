"""从期刊规则集渲染一份版式样板 docx。

样板的作用有三个：
1. 用户可在 Word 里直接核对字号、行距、表注层次
2. 作为验收基准：样式包变更后重新渲染并比对，即可发现版式漂移
3. 文字内容为格式说明，样板本身可读，不依赖文档

用法：python tools/render_sample.py rules/journal/gjjryj.json -o rules/journal/gjjryj.sample.docx
"""

import argparse
import json
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

import sys
from pathlib import Path

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_meta import sanitize_docx

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_run(run, cn=None, latin=None, size=None, bold=None, letter_spacing=None):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    if cn:
        rf.set(qn("w:eastAsia"), cn)
    if latin:
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if letter_spacing:
        # 字间距（Word 的「加宽」）：单位 twips，1pt=20twips；写进 w:spacing/@w:val
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(round(letter_spacing * 20))))
        rpr.append(sp)


def set_chars_indent(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(int(chars * 210)))


def set_hanging_indent(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:leftChars"), str(int(chars * 100)))
    ind.set(qn("w:left"), str(int(chars * 210)))
    ind.set(qn("w:hangingChars"), str(int(chars * 100)))
    ind.set(qn("w:hanging"), str(int(chars * 210)))


def set_line_unit_space(p, before_lines=0, after_lines=0):
    ppr = p._p.get_or_add_pPr()
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        ppr.append(sp)
    sp.set(qn("w:beforeLines"), str(int(before_lines * 100)))
    sp.set(qn("w:afterLines"), str(int(after_lines * 100)))


def add_para(doc, text, *, cn=None, latin=None, size=None, bold=None,
             align=None, line=None, indent_chars=None,
             before_lines=0, after_lines=0, keep=False, letter_spacing=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = ALIGN[align]
    if line:
        if isinstance(line, dict):
            p.paragraph_format.line_spacing = line["mult"]      # 倍数 → lineRule=auto
        else:
            p.paragraph_format.line_spacing = Pt(line)
    if indent_chars:
        set_chars_indent(p, indent_chars)
    if before_lines or after_lines:
        set_line_unit_space(p, before_lines, after_lines)
    if keep:
        p.paragraph_format.keep_with_next = True
    if text:
        # 双语参考文献条目要在一段里换行，用 \n 分段
        segs = str(text).split("\n")
        run = p.add_run(segs[0])
        set_run(run, cn=cn, latin=latin, size=size, bold=bold,
                letter_spacing=letter_spacing)
        for seg in segs[1:]:
            run.add_break()
            r2 = p.add_run(seg)
            set_run(r2, cn=cn, latin=latin, size=size, bold=bold)
    return p


def three_line_table(doc, headers, rows, bp, header_size, body_size):
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(int(bp[edge] * 8)))
        else:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)

    for c, h in enumerate(headers):
        cell = table.rows[0].cells[c]
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(h)
        set_run(run, cn="宋体", latin="Times New Roman", size=header_size)
        tcpr = cell._tc.get_or_add_tcPr()
        tcb = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), str(int(bp["header_bottom"] * 8)))
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        tcb.append(bottom)
        tcpr.append(tcb)

    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(val))
            set_run(run, cn="宋体", latin="Times New Roman", size=body_size)
    return table


def build(rules, out_path):
    j = rules["journal"]
    tp = rules["typography"]
    hd = rules["headings"]
    tb = rules["tables"]
    fig = rules["figures"]
    body = tp["body"]
    # 倍数行距（mode=multiple）要真的按倍数写进 docx（lineRule=auto）。
    # 折成固定磅值的话，作者在 Word 里看到的是「固定值 18.9 磅」，与规范页写的「1.5 倍」对不上。
    ls = body["line_spacing"]
    line_mult = ls.get("value") if ls.get("mode") == "multiple" else None
    line_pt = ls.get("value_pt") or ls.get("equivalent_exact_pt")
    if line_pt is None:
        line_pt = round(body["size_pt"] * 1.2 * (line_mult or 1), 1)
    # LINE 给正文用（倍数或固定值），line_pt 只用于算标题／行首这类需要数值的地方
    LINE = {"pt": line_pt, "mult": line_mult} if line_mult else line_pt
    line_desc = f"{line_mult} 倍" if line_mult else f"{line_pt} 磅固定值"
    indent = body["first_line_indent_chars"]

    doc = Document()

    # 页面设置遍历全部节（F32）：只设 sections[0] 的话，分节符造出的其余节
    # 保持 Word 默认版面。样板虽从空白文档生成（正常单节），遍历姿势与
    # render_from_ir 的页面设置保持同一，多节也不漏。
    for sec in doc.sections:
        sec.page_width = Mm(rules["page"]["size_mm"]["width"])
        sec.page_height = Mm(rules["page"]["size_mm"]["height"])
        for side, key in (("top", "top"), ("bottom", "bottom"), ("left", "left"), ("right", "right")):
            setattr(sec, f"{side}_margin", Mm(rules["page"]["margin_mm"][key]))

    normal = doc.styles["Normal"]
    normal.font.name = body["font_latin"]
    normal.font.size = Pt(body["size_pt"])
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), body["font_cn"])

    title_tp = tp["title"]
    # 样板占位文字可由规则集提供：英文刊需要英文样例，否则出现「张　三」这类中文占位
    smp = rules.get("sample") or {}

    # ---------- 作者信息页（仅当该刊要求独立成页时渲染） ----------
    if rules["submission"]["author_page"].get("separate_page"):
        add_para(doc, smp.get("author_page_title") or (j["name_zh"] + " 投稿样板（作者信息页）"),
                 cn=title_tp["font_cn"],
                 size=title_tp["size_pt"], align="center", line=20)
        add_para(doc, "", line=LINE)
        for fid in rules["submission"]["author_page"]["fields"]:
            label = next((f["label_zh"] for f in rules["personal_fields"] if f["id"] == fid), fid)
            add_para(doc, f"{label}：", cn=body["font_cn"], latin=body["font_latin"],
                     size=body["size_pt"], line=LINE)
        add_para(doc, "", line=LINE)
        add_para(doc, smp.get("author_page_note") or
                 "※ 本页独立成页，置于正文之前，请勿单独提交；正文部分须为匿名形式。",
                 cn=body["font_cn"], size=9, line=LINE)

        doc.add_page_break()

    # ---------- 栏目名（如「·构建新发展格局·」） ----------
    # 由编辑部加，不是作者必填项；只有 blocks 声明了才渲染，避免凭空多一行
    col_blk = next((b for b in rules["blocks"] if b["type"] == "column_name"), None)
    if col_blk:
        ctp = tp.get("column_name") or {}
        add_para(doc, smp.get("column_name") or col_blk.get("sample_value")
                 or col_blk.get("label") or "·构建新发展格局·",
                 cn=ctp.get("font_cn", "黑体"), size=ctp.get("size_pt", body["size_pt"]),
                 align=ctp.get("align", "left"), line=LINE,
                 indent_chars=ctp.get("indent_chars") or None, after_lines=0.3)

    # ---------- 正文标题 ----------
    # 标题行距必须大于字号，否则多行标题会重叠
    title_line = max(line_pt, round(title_tp["size_pt"] * 1.35))
    add_para(doc, smp.get("title") or "美国宽松货币政策对金砖国家的溢出效应研究",
             letter_spacing=title_tp.get("letter_spacing_pt"),
             cn=title_tp["font_cn"],
             size=title_tp["size_pt"], align=title_tp["align"], line=title_line)

    # ---------- 副标题与作者 ----------
    # 只有规则集明确规定了才渲染，避免给没有该构件的期刊凭空加行
    sub_tp = tp.get("subtitle")
    if sub_tp:
        sub_line = max(line_pt, round(sub_tp["size_pt"] * 1.35))
        add_para(doc, smp.get("subtitle") or "——基于金砖五国宏观月度数据的证据", cn=sub_tp["font_cn"],
                 latin=body.get("font_latin"), size=sub_tp["size_pt"],
                 align=sub_tp.get("align", "center"), line=sub_line)
    au_tp = tp.get("authors")
    if au_tp:
        au_line = max(line_pt, round(au_tp["size_pt"] * 1.35))
        add_para(doc, smp.get("authors") or "张　三　　李　四", cn=au_tp["font_cn"],
                 latin=body.get("font_latin"),
                 size=au_tp["size_pt"], align=au_tp.get("align", "center"), line=au_line)

    # ---------- 作者单位 ----------
    # 此前样板只渲染英文单位，中文刊的 typography.affiliations 一直画不出来
    # （§16.6 同类问题：规则集声明了、渲染器不消费，样板与规则集漂移）
    aff_blk_zh = next((b for b in rules["blocks"]
                       if b["type"] == "affiliations" and b.get("lang") != "en"), None)
    aff_tp_zh = tp.get("affiliations")
    if aff_blk_zh and aff_tp_zh:
        aff_line_zh = max(line_pt, round(aff_tp_zh["size_pt"] * 1.35))
        add_para(doc, smp.get("affiliations") or
                 "（1. 样本大学经济学院，北京　100000；2. 样本研究院，上海　200000）",
                 cn=aff_tp_zh["font_cn"], latin=aff_tp_zh.get("font_latin", body.get("font_latin")),
                 size=aff_tp_zh["size_pt"], align=aff_tp_zh.get("align", "center"),
                 line=aff_line_zh)

    # ---------- 内容摘要 ----------
    ab_tp = tp.get("abstract") or body      # 摘要常单独定字体（楷体/仿宋）
    abs_block = next((b for b in rules["blocks"] if b["type"] == "abstract" and b.get("lang") == "zh"),
                     next(b for b in rules["blocks"] if b["type"] == "abstract"))
    abs_text = smp.get("abstract") or (
        "本文以美联储及金砖五国宏观月度数据为样本，考察美国宽松货币政策对金砖国家的溢出效应。"
        f"此处文字仅用于核验版式：{ab_tp['font_cn']} {ab_tp['size_pt']}pt、"
        f"行距 {line_desc}、首行缩进 {indent} 字符。")
    add_para(doc, f"{abs_block.get('label') or ''}{abs_text}",
             cn=ab_tp["font_cn"], latin=ab_tp.get("font_latin", body["font_latin"]),
             size=ab_tp["size_pt"], align="both", line=LINE, indent_chars=indent)

    kw_block = next((b for b in rules["blocks"] if b["type"] == "keywords" and b.get("lang") == "zh"),
                    next(b for b in rules["blocks"] if b["type"] == "keywords"))
    kw_items = smp.get("keywords") or ["宽松货币政策", "美联储", "金砖国家", "溢出效应"]
    add_para(doc, (kw_block.get("label") or "") + kw_block.get("separator", "  ").join(kw_items),
             cn=ab_tp["font_cn"], latin=ab_tp.get("font_latin", body["font_latin"]),
             size=ab_tp["size_pt"], align="both", line=LINE, indent_chars=indent)

    # 号段（DOI、中图分类号、文献标识码、文章编号）在多数刊里是小字号、黑体，
    # 与正文不同。此前样板统一按正文样式画，凡规则集写了 typography.classification
    # 的刊都会与规则集漂移（§16.6）。这里改为读该键，缺省才退回正文样式。
    code_tp = tp.get("classification") or {}
    code_cn = code_tp.get("font_cn") or body["font_cn"]
    code_latin = code_tp.get("font_latin") or body.get("font_latin")
    code_size = code_tp.get("size_pt") or body["size_pt"]
    # 小字号行按 1.45 倍折算行距，避免 9pt 字撑在 17pt 行里显得过散
    code_line = code_tp.get("line_spacing", {}).get("value_pt") or round(code_size * 1.45, 1)

    def code_para(text):
        add_para(doc, text, cn=code_cn, latin=code_latin, size=code_size,
                 line=code_line, indent_chars=indent)

    doi_block = next((b for b in rules["blocks"] if b["type"] == "doi"), None)
    if doi_block:
        code_para((doi_block.get("label") or "DOI: ") +
                  str(doi_block.get("sample_value") or "10.19343/j.cnki.11-1302/c.2026.10.001"))

    cls_blocks = [b for b in rules["blocks"]
                  if b["type"] in ("classification_code", "document_code", "article_number")
                  and b.get("lang") != "en"]
    if cls_blocks:
        codes_map = {"CLC": "F821", "JEL": "E52  F41  G15", "CLC-DOC": "A"}
        segs = []
        for b in cls_blocks:
            # sample_value 让规则集直接给出样例值；缺省按 scheme 取通用值
            val = b.get("sample_value")
            if val is None:
                if b["type"] == "classification_code":
                    val = codes_map.get(b.get("scheme", ""), "XX")
                else:
                    val = b.get("fixed_value", "")
            segs.append(b.get("label", "") + str(val))
        code_para("        ".join(segs))

    body_pos = next((i for i, b in enumerate(rules["blocks"]) if b["type"] == "body"), 10**6)

    def en_cls_line(blk):
        add_para(doc, (blk.get("label") or "JEL: ") + ", ".join(smp.get("jel") or ["F13", "F14"]),
                 cn=body["font_cn"], latin=body["font_latin"], size=body["size_pt"],
                 align="both", line=LINE, indent_chars=indent)

    for i, blk in enumerate(rules["blocks"]):
        if blk["type"] == "classification_code" and blk.get("lang") == "en" and i < body_pos:
            en_cls_line(blk)

    # 英文四件套的位置由 blocks 顺序决定：排在正文之前（中英文摘要同置首页，如统计研究、
    # 财经研究）就紧跟中文摘要；排在正文之后（世界经济、财贸经济、经济管理）就落到文末。
    # 此前一律追加在末尾，凡是首页并排的刊，样板都会把英文部分错放到参考文献之后。
    en_idx = [i for i, b in enumerate(rules["blocks"])
              if b.get("lang") == "en" and b["type"] in ("title", "authors", "abstract", "keywords")]
    en_before_body = bool(en_idx) and min(en_idx) < body_pos
    # 纯英文刊（没有中文摘要块）主体已是英文，不再追加一份「英文四件套」
    has_zh = any(b["type"] in ("abstract", "keywords") and b.get("lang") != "en"
                 for b in rules["blocks"])

    def render_english():
        en_title = tp.get("title_en") or tp["title"]
        en_title_ls = (en_title.get("line_spacing") or {}).get("value_pt")
        add_para(doc, smp.get("title_en") or
                 "The Spillover Effects of US Monetary Policy on BRICS Countries",
                 cn=en_title.get("font_cn"), latin=en_title.get("font_latin", body.get("font_latin")),
                 size=en_title.get("size_pt"), bold=en_title.get("bold", False),
                 align=en_title.get("align", "center"), line=en_title_ls or LINE)
        # 英文作者：只有规则集声明了英文作者块才渲染，避免给不需要的期刊凭空加行
        if any(b["type"] == "authors" and b.get("lang") == "en" for b in rules["blocks"]):
            en_au = tp.get("authors_en") or tp.get("authors") or body
            au_size = en_au.get("size_pt", body["size_pt"])
            add_para(doc, smp.get("authors_en") or "ZHANG San　LI Si",
                     cn=en_au.get("font_cn", body["font_cn"]), latin=body["font_latin"],
                     size=au_size, align=en_au.get("align", "center"),
                     line=max(line_pt, round(au_size * 1.5)))
        # 英文单位：只有规则集声明了才渲染（经济研究的 a/b/c 分行）
        if any(b["type"] == "affiliations" and b.get("lang") == "en" for b in rules["blocks"]):
            aff_tp = tp.get("affiliations_en") or body
            add_para(doc, smp.get("affiliations_en") or
                     "(a: School of Economics, Sample University; b: Sample Research Institute)",
                     cn=aff_tp.get("font_cn", body["font_cn"]),
                     latin=aff_tp.get("font_latin", body["font_latin"]),
                     size=aff_tp.get("size_pt", body["size_pt"]),
                     align=aff_tp.get("align", "center"), line=LINE)
        en_abs = next((b for b in rules["blocks"]
                       if b["type"] == "abstract" and b.get("lang") == "en"), {})
        en_kw = next((b for b in rules["blocks"]
                      if b["type"] == "keywords" and b.get("lang") == "en"), {})
        abs_tp = tp.get("abstract_en") or body
        # 行距取值必须落回数值：LINE 在倍数行距下是 dict（{"pt","mult"}），
        # 直接拼进占位文字会把字典原样印到样板上（财经问题研究样板曾出现）
        _en_abs_ls = abs_tp.get("line_spacing") or {}
        abs_ls = (_en_abs_ls.get("value_pt") or _en_abs_ls.get("equivalent_exact_pt")
                  or line_pt)
        add_para(doc, (en_abs.get("label") or "Abstract: ") +
                      (smp.get("abstract_en") or
                       "This sample is generated for layout verification. "
                       f"{body.get('font_latin', 'Times New Roman')}, {abs_tp['size_pt']} pt, "
                       f"exact {abs_ls} pt line spacing."),
                 cn=abs_tp.get("font_cn", body["font_cn"]),
                 latin=abs_tp.get("font_latin", body["font_latin"]),
                 size=abs_tp["size_pt"], align="both", line=abs_ls, indent_chars=indent)
        kw_tp = tp.get("keywords_en") or abs_tp
        kw_ls = ((kw_tp.get("line_spacing") or {}).get("value_pt")
                 or (kw_tp.get("line_spacing") or {}).get("equivalent_exact_pt")
                 or abs_ls)
        add_para(doc, (en_kw.get("label") or "Key words: ") +
                      en_kw.get("separator", "; ").join(
                          smp.get("keywords_en") or
                          ["Quantitative Easing", "the Fed", "BRICS Countries", "Spillover Effect"]),
                 cn=kw_tp.get("font_cn", body["font_cn"]),
                 latin=kw_tp.get("font_latin", body["font_latin"]),
                 size=kw_tp["size_pt"], align="both", line=kw_ls, indent_chars=indent)

    if has_zh and en_before_body:
        render_english()

    # ---------- 引言与层级标题 ----------
    lv1 = hd["level1"]
    cn_seq = "一二三四五六七八九十"
    numeric = str(hd.get("numbering", "")).startswith("numeric")
    # 一级标题序号分隔符取 levels[0] 的末字符：可能是「、」，也可能是全角空格（如「一　引言」）
    # 序号分隔符＝levels[0] 去掉首个序号字符后的剩余部分：
    # 「一　」→全角空格、「一、」→顿号、「一、 」→顿号加空格（经济研究）
    sep = " " if numeric else hd["levels"][0][1:]

    def lv1_line():
        """一级标题行距：规则集可用 line_multiple 指定倍数（如两倍行距）。"""
        mult = lv1.get("line_multiple")
        return max(line_pt, round(lv1["size_pt"] * mult)) if mult else line_pt

    def head_no(level, n):
        if numeric:
            # 模板取自 levels[0]：写「1」得「1 引言」（中国管理科学），写「1.」得「1. Introduction」（CER）
            return {1: hd["levels"][0].replace("1", str(n), 1),
                    2: hd["levels"][1], 3: hd["levels"][2]}[level]
        return cn_seq[n - 1] if level == 1 else hd["levels"][level - 1]

    def heading(level, n, text):
        lv = hd.get(f"level{level}", lv1)
        num = head_no(level, n)
        # 数字体系每级都要空格；中文体系只有一级用分隔符（「（一）标题」不带顿号）
        joiner = sep if (numeric or level == 1) else ""
        return add_para(doc, f"{num}{joiner}{text}",
                        cn=lv.get("font_cn", lv1["font_cn"]),
                        size=lv.get("size_pt", lv1["size_pt"]),
                        bold=lv.get("bold", False),
                        align=lv1["align"] if level == 1 else lv.get("align", "left"),
                        line=lv1_line() if level == 1 else line_pt,
                        before_lines=lv.get("space_before_lines", 0) if level == 1 else 0,
                        after_lines=lv.get("space_after_lines", 0) if level == 1 else 0)

    # 引言编号：多数刊从 1 开始，管理科学学报写成「0 引言」，后文从 1 起
    intro_no = hd.get("intro_number", 1)
    sec_no = 0
    intro_unnumbered = (smp.get("intro") or "引言") in hd.get("unnumbered_sections", [])
    if intro_unnumbered:
        add_para(doc, hd.get("intro_title_text") or smp.get("intro") or "引言",
                 cn=lv1["font_cn"], size=lv1["size_pt"], bold=lv1.get("bold", False),
                 align=lv1["align"], line=lv1_line(),
                 before_lines=lv1.get("space_before_lines", 0),
                 after_lines=lv1.get("space_after_lines", 0))
    else:
        sec_no = intro_no
        heading(1, sec_no, smp.get("intro") or "引言")
    add_para(doc, smp.get("body") or
             (f"次贷危机爆发之后，美国经济陷入严重衰退。此段用于核验正文版式："
              f"{body['font_cn']} {body['size_pt']} pt，英文与数字为 {body.get('font_latin','Times New Roman')}，"
              f"行距 {line_desc}，首行缩进 {indent} 字符，两端对齐。"),
             cn=body["font_cn"], latin=body["font_latin"], size=body["size_pt"],
             align="both", line=LINE, indent_chars=indent or None)

    # ---------- 一级标题 ----------
    sec_no += 1
    heading(1, sec_no, smp.get("section2") or "文献综述")
    add_para(doc, (smp.get("l1_note") or
                   f"一级标题：{lv1.get('font_cn', '')} {lv1.get('size_pt', '')}pt，"
                   f"{'加粗' if lv1.get('bold') else '不加粗'}，"
                   f"{'居中' if lv1['align'] == 'center' else '左对齐'}，单独占行。"),
             cn=body["font_cn"], size=body["size_pt"], align="both", line=LINE, indent_chars=indent or None)

    # ---------- 二级标题 ----------
    lv2 = hd.get("level2", {})
    heading(2, sec_no, smp.get("section2_1") or "理论基础")
    add_para(doc, (smp.get("l2_note") or
                   f"二级标题：{hd['levels'][1]} 式编号，"
                   f"{lv2.get('font_cn', body['font_cn'])} {lv2.get('size_pt', body['size_pt'])}pt，"
                   f"缩进 {lv2.get('indent_chars', 2)} 字，单独占行。"),
             cn=body["font_cn"], size=body["size_pt"], align="both", line=LINE, indent_chars=indent or None)

    # ---------- 三级标题 ----------
    lv3 = hd.get("level3", {})
    l3_num = head_no(3, sec_no)
    l3_text = smp.get("section2_1_1") or "货币政策传导机制"
    l3_note = (smp.get("l3_note") or
               f"三级标题：{hd['levels'][2]} 式编号，"
               f"{lv3.get('font_cn', body['font_cn'])} {lv3.get('size_pt', body['size_pt'])}pt，"
               f"缩进 {lv3.get('indent_chars', 2)} 字，"
               f"{'与正文接排' if lv3.get('inline_with_body') else '单独占行'}。")
    if lv3.get("inline_with_body"):
        # 接排体例：「1.制造业行业差异。接下来的正文……」编号行与正文同一段
        add_para(doc, f"{l3_num}{l3_text}。{l3_note}",
                 cn=lv3.get("font_cn") or body["font_cn"],
                 size=lv3.get("size_pt") or body["size_pt"],
                 align="both", line=LINE, indent_chars=indent or None)
    else:
        heading(3, sec_no, l3_text)
        add_para(doc, l3_note,
                 cn=body["font_cn"], size=body["size_pt"], align="both", line=LINE, indent_chars=indent or None)

    # ---------- 表格 ----------
    t_cap = smp.get("table_caption") or "表1  主要变量描述性统计"
    t_head = smp.get("table_headers") or ["变量", "样本量", "均值", "标准差", "最小值", "最大值"]
    t_rows = smp.get("table_rows") or [["M2", "60", "8.421①", "0.317", "7.902", "9.115"],
                                       ["GDP", "60", "12.884", "0.642", "11.933", "14.012"]]
    add_para(doc, t_cap, cn=tb["caption_font_cn"], size=tb["caption_size_pt"],
             align=tb["caption_align"], line=LINE, keep=True)
    bp = tb.get("border_pt") or {"top": 1.5, "header_bottom": 0.5, "bottom": 1.5}
    three_line_table(
        doc,
        t_head, t_rows,
        bp,
        header_size=tb.get("font_size_pt", 10.5),
        body_size=tb.get("font_size_pt", 10.5),
    )
    for layer in tb["note_layers"]:
        prefix = layer.get("prefix") or ""
        sample = (smp.get("table_notes") or {}).get(layer["kind"]) or {
            "source": "Wind数据库",
            "general": "括号内为标准差；*、**、***分别表示在10%、5%、1%的水平下显著。",
            "data_source": "国家统计局年度数据。",
            "cell": "①该数值保留四位小数，原值为8.42061。",
        }[layer["kind"]]
        add_para(doc, f"{prefix}{sample}", cn=layer.get("font_cn"), size=layer.get("size_pt"),
                 line=None if tb.get("note_line_spacing", {}).get("mode") == "multiple" else line_pt,
                 indent_chars=layer.get("first_line_indent_chars"))

    # ---------- 图 ----------
    add_para(doc, smp.get("figure_placeholder") or "［插图位置：矢量图，无边框］", cn=None,
             size=body["size_pt"], align="center", line=LINE)
    add_para(doc, smp.get("figure_caption") or "图1  美联储资产负债表规模变化",
             cn=fig["caption_font_cn"], size=fig["caption_size_pt"],
             align=fig["caption_align"], line=LINE, keep=True)
    fig_samples = (smp.get("figure_notes") or {}) or {"general": "曲线为 HP 滤波后的趋势项。",
                   "source": "Wind数据库", "data_source": "国家统计局年度数据。", "cell": ""}
    for layer in fig["note_layers"]:
        prefix = layer.get("prefix") or ""
        add_para(doc, prefix + fig_samples.get(layer["kind"], ""),
                 cn=layer.get("font_cn"), size=layer.get("size_pt"),
                 align=layer.get("align"), line=LINE)

    # ---------- 生成式 AI 声明（官方要求时排在参考文献之前，如 Elsevier 系英文刊） ----------
    ai_decl = (rules.get("submission") or {}).get("ai_declaration") or {}
    if ai_decl.get("required"):
        add_para(doc, ai_decl.get("section_title") or
                 "Declaration of generative AI and AI-assisted technologies in the manuscript preparation process",
                 cn=body.get("font_latin"), latin=body.get("font_latin"),
                 size=body["size_pt"], bold=True, align="left", line=line_pt, keep=True)
        add_para(doc, ai_decl.get("sample_text") or
                 "During the preparation of this work the author(s) used [NAME OF TOOL / SERVICE] "
                 "in order to [REASON]. After using this tool/service, the author(s) reviewed and "
                 "edited the content as needed and take(s) full responsibility for the content of "
                 "the published article.",
                 cn=body.get("font_latin"), latin=body.get("font_latin"),
                 size=body["size_pt"], align="both", line=line_pt)

    # ---------- 参考文献 ----------
    ref = next(b for b in rules["blocks"] if b["type"] == "references")
    rh = tp.get("references_heading") or {}
    add_para(doc, ref.get("label", "参考文献："), cn=rh.get("font_cn", "黑体"),
             size=rh.get("size_pt", body["size_pt"]),
             align=rh.get("align", "left"), line=LINE,
             indent_chars=rh.get("indent_chars") or None, keep=True)
    # 规则集给了示例就用示例——文献格式逐刊不同，不该由代码兜底
    examples = rules["citation"].get("reference_list", {}).get("examples")
    if examples:
        ref_items = list(examples)
    elif rules["citation"]["system"] == "author-year":
        ref_items = [
            "何国华和彭意，2011，《美、日货币政策对中国产出的溢出效应研究》，《国际金融研究》第6期，第23~30页。",
            "谭小芬、熊爱宗和陈思翀，2018，《美国量化宽松的退出机制、溢出效应与中国的对策》，《世界经济》第4期，第3~25页。",
            "Eickmeier, S., L. Gambacorta and B. Hofmann, 2013, \"Understanding Global Liquidity,\" BIS Working Papers, No. 402.",
        ]
    else:
        ref_items = [
            "[1] 何国华，彭意. 美、日货币政策对中国产出的溢出效应研究[J]. 国际金融研究，2011（6）：23-30",
            "[2] Eickmeier S，Gambacorta L，Hofmann B. Understanding Global Liquidity[R]. BIS Working Papers，2013",
        ]
    hang = rules["citation"].get("reference_list", {}).get("hanging_indent_chars")
    # 文献条目版式逐刊不同（六号单倍 / 小五固定行距 / 首行缩进 2 字），必须查规则集
    ref_tp = tp.get("references") or {}
    ref_size = ref_tp.get("size_pt", body["size_pt"])
    ref_ls = ref_tp.get("line_spacing") or {}
    ref_line = ref_ls.get("value_pt") or ref_ls.get("equivalent_exact_pt") or line_pt
    ref_indent = ref_tp.get("indent_chars",
                            rules["citation"].get("reference_list", {}).get("indent_chars", 0))
    # 编号制文献表：规则集只写 number_format，编号由渲染器按条数生成，
    # 否则每本刊的示例都要手写 [1][2][3]，漏一条就对不上（jjgl/jqte/tjyj 都踩过）
    ref_cfg = rules["citation"].get("reference_list", {})
    auto_no = bool(ref_cfg.get("numbered")) and not any(
        re.match(r"^[\[\(（［〔]\s*\d", x) for x in ref_items)

    def ref_no(n, fmt):
        fmt = fmt or "[1]"
        # 六角括号〔1〕是国内部分刊（管理科学学报）的写法，与方括号要分开
        if fmt.startswith("〔"):
            return f"〔{n}〕"
        if fmt.startswith("［"):
            return f"［{n}］ "
        if fmt.startswith("["):
            return f"[{n}] "
        if fmt.startswith("（"):
            return f"（{n}）"
        if fmt.startswith("("):
            return f"({n})"
        return f"{n}．"

    for k, item in enumerate(ref_items, 1):
        text = ref_no(k, ref_cfg.get("number_format")) + item if auto_no else item
        p = add_para(doc, text, cn=ref_tp.get("font_cn", body["font_cn"]),
                     latin=ref_tp.get("font_latin", body["font_latin"]),
                     size=ref_size, align="both", line=ref_line,
                     indent_chars=ref_indent or None)
        if hang:
            set_hanging_indent(p, hang)

    # ---------- 作者简介（参考文献之后，按 blocks 顺序判断） ----------
    ref_pos = next(i for i, b in enumerate(rules["blocks"]) if b["type"] == "references")
    note_blk = next(((i, b) for i, b in enumerate(rules["blocks"])
                     if b["type"] == "author_note" and i > ref_pos), (None, None))[1]
    if note_blk:
        an = tp.get("author_note") or {}
        an_size = an.get("size_pt", body["size_pt"])
        an_ls = an.get("line_spacing") or {}
        an_line = an_ls.get("value_pt") or an_ls.get("equivalent_exact_pt") or line_pt
        add_para(doc, note_blk.get("label") or "作者简介",
                 cn=an.get("label_font_cn", "黑体"), size=an_size,
                 align=an.get("heading_align", "left"),
                 line=an_line, indent_chars=an.get("heading_indent_chars") or None, keep=True)
        for line in (smp.get("author_note") or
                     ["张　三，男，经济学博士，某某大学某某学院教授。研究方向为货币理论与政策。"
                      "电子邮箱：zhangsan@example.edu.cn。"]):
            add_para(doc, line, cn=an.get("font_cn", body["font_cn"]),
                     latin=an.get("font_latin", body["font_latin"]),
                     size=an_size, align="both", line=an_line,
                     indent_chars=an.get("heading_indent_chars") or None)

    # 排在正文之后的英文四件套（世界经济、财贸经济这类把英文摘要放文末的刊）
    # 只有规则集声明了英文块才追加文末英文区——不声明的刊（如财经问题研究，
    # 投稿模板不含英文摘要）不应凭空多出一段（§9：未声明该构件则不渲染）
    if has_zh and en_idx and not en_before_body:
        render_english()

    # 排在正文之后的英文分类号（如《世界经济》把 JEL 放在全文末尾）
    for i, blk in enumerate(rules["blocks"]):
        if blk["type"] == "classification_code" and blk.get("lang") == "en" and i > body_pos:
            en_cls_line(blk)

    # ---------- 生成说明 ----------
    add_para(doc, "", line=LINE)
    if smp.get("sample_note"):
        add_para(doc, smp["sample_note"], cn=body["font_cn"], size=9, line=LINE)
        doc.save(out_path)
        sanitize_docx(out_path)
        return out_path
    src_parts = [f"本样板由规则集 {j['id']}.json 自动生成，依据 {rules['verification']['source_url']}"]
    if rules["verification"].get("template_url"):
        src_parts.append(f"与官方模板（{rules['verification'].get('template_date', '')}）")
    if rules.get("assumptions"):
        src_parts.append("标注「推定」的项为官方未规定、由本工具假定的值，不属于期刊要求")
    src_parts.append("文字内容为版式核验用途，非真实论文")
    add_para(doc, "；".join(src_parts) + "。", cn=body["font_cn"], size=9, line=LINE)

    doc.save(out_path)
    sanitize_docx(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()
    with open(args.ruleset, encoding="utf-8") as f:
        rules = json.load(f)
    path = build(rules, args.out)
    print("已生成:", path)


if __name__ == "__main__":
    main()
