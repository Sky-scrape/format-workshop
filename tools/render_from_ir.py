"""把 Paper IR 渲染成目标期刊格式，并输出审计报告。

实现选择：就地重排（打开原稿 → 逐段套目标格式 → 另存），而不是新建文档重新拼装。
原因：原稿里的图片、公式对象、脚注都带关系 ID，跨文档复制必然断链，
重建一个文档会静默丢失这些内容。

红线：正文与表格内的文字零改动。只改版式、编号、标签，且全部进审计报告。

用法：python tools/render_from_ir.py <ir.json> <ruleset.json> -o out.docx
"""

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.table import Table
from lxml import etree
from docx.text.paragraph import Paragraph

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 规则集用中文描述题注编号与标题的间隔，这里映射成实际空格
CAPTION_GAP = {"一字空": "  ", "半字空": " ", "两字空": "    ", "无间隔": ""}

CN_SEQ = "一二三四五六七八九十"
NUM_PREFIX = {
    1: re.compile(r"^[一二三四五六七八九十]{1,3}、\s*"),
    2: re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]\s*"),
    3: re.compile(r"^\d{1,2}\s*[.．]\s*"),
    4: re.compile(r"^[（(]\d{1,2}[）)]\s*"),
}
# 数字编号体系（1 / 1.1 / 1.1.1）的前缀，用于剥掉稿件原有的编号
NUM_PREFIX_NUM = {
    1: re.compile(r"^\d{1,2}\.\s*"),
    2: re.compile(r"^\d{1,2}\.\d{1,2}\s*"),
    3: re.compile(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\s*"),
}


# 用户补录后注入的语义块 → (规则集 typography 键, 对齐, 加粗)
# 只做"类型到语义"的映射，不含任何期刊判断；字号字体一律查规则集
SEMANTIC = {
    "affiliations": ("authors", "center", False),
    "english_authors": ("authors", "center", False),
    "english_affiliations": ("abstract", "center", False),
    "classification": ("abstract", "left", False),
    "english_title": ("title", "center", False),
    "english_abstract": ("abstract", "both", False),
    "english_keywords": ("abstract", "both", False),
}


class Audit:
    def __init__(self, rejects=None):
        self.items = []
        self.sec = "前置"
        # 驳回记录用「块序号|动作」定位：同一稿件、同一规则集重跑时该键恒定，
        # 因此驳回可以跨次复用，不需要依赖易变的数组下标
        self.rejects = set(rejects or [])

    def blocked(self, idx, action):
        return f"{idx}|{action}" in self.rejects

    def add(self, idx, btype, action, before="", after="", note="", rejected=False, rule=""):
        """rule 是本条改动的依据，指向规则集里的具体字段（见 rules_ref.py）。

        结尾带 `~` 表示该键在部分规则集里可缺省，缺省时按代码内约定取值。
        """
        self.items.append({"id": f"{idx}|{action}", "idx": idx, "sec": self.sec, "block": btype,
                           "action": action, "before": before, "after": after, "note": note,
                           "rule": rule, "rejected": bool(rejected)})


def fmt_snapshot(p):
    """取段落当前的字体、字号与行距，用于判断这次套版到底改没改东西。"""
    sizes, cns, latins = set(), set(), set()
    for r in p.runs:
        if r.font.size is not None:
            sizes.add(round(r.font.size.pt, 2))
        rpr = r._element.rPr
        if rpr is not None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                if rf.get(qn("w:eastAsia")):
                    cns.add(rf.get(qn("w:eastAsia")))
                if rf.get(qn("w:ascii")):
                    latins.add(rf.get(qn("w:ascii")))
    ls = p.paragraph_format.line_spacing
    line = round(ls.pt, 2) if hasattr(ls, "pt") else (ls if isinstance(ls, (int, float)) else None)
    return {"sizes": sizes, "cn": cns, "latin": latins, "line": line}


def fmt_delta(before, after, target):
    """把「改了什么」写成一句话。没变化就只报目标值。"""
    names = {("sizes", "字号"), ("cn", "中文字体"), ("latin", "西文字体"), ("line", "行距")}
    changed = [label for key, label in sorted(names) if before[key] != after[key]]
    return (f"改{'、'.join(changed)} → {target}" if changed else target)


def assumed(tp):
    """规则集标了 measured: false 的属于推定值，报告里必须说明。

    用户有权知道这个字号是期刊明文写的，还是我们照着别的刊推出来的。
    """
    return "，推定值（样本未实测）" if isinstance(tp, dict) and tp.get("measured") is False else ""


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
        # 字间距按 twips 写入（1pt = 20twips）；《经济研究》标题每字加宽约 1.3pt
        old = rpr.find(qn("w:spacing"))
        if old is not None:
            rpr.remove(old)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(round(letter_spacing * 20))))
        rpr.append(sp)


def clear_para_formatting(p):
    ppr = p._p.get_or_add_pPr()
    for tag in ("w:spacing", "w:ind", "w:jc", "w:numPr"):
        el = ppr.find(qn(tag))
        if el is not None:
            ppr.remove(el)


def set_chars_indent(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(int(chars * 210)))


def set_hanging(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:leftChars"), str(int(chars * 100)))
    ind.set(qn("w:left"), str(int(chars * 210)))
    ind.set(qn("w:hangingChars"), str(int(chars * 100)))
    ind.set(qn("w:hanging"), str(int(chars * 210)))


def head_line_pt(lv, size, body_line):
    """标题行距：规则集可用 line_multiple 指定倍数（如一级标题两倍行距）。

    倍数只在大于正文字号时生效——否则「20 磅固定值 + 22pt 标题」会重叠。
    """
    mult = lv.get("line_multiple")
    if mult:
        return max(body_line, round(size * mult))
    return body_line


def restyle(p, *, cn, latin, size, align, line_pt, indent=None, bold=None,
            before_pt=None, after_pt=None, letter_spacing=None, line_rule="exact"):
    clear_para_formatting(p)
    if align:
        p.alignment = ALIGN[align]
    if line_pt:
        p.paragraph_format.line_spacing = Pt(line_pt)
        if line_rule == "at_least":
            # 「固定值」会裁掉超出行高的内容（分式、积分、上下标、内嵌图），
            # 「最小值」只保底不封顶，视觉上一样，但不会被裁
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    # 前置区的段间距：标题、作者、摘要之间靠它撑开，否则 26pt 标题会与下一行贴死
    if before_pt is not None:
        p.paragraph_format.space_before = Pt(before_pt)
    if after_pt is not None:
        p.paragraph_format.space_after = Pt(after_pt)
    if indent:
        set_chars_indent(p, indent)
    for run in p.runs:
        set_run(run, cn=cn, latin=latin, size=size, bold=bold,
                letter_spacing=letter_spacing)


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def replace_text(p, new_text):
    """替换段落文字，但保留公式、脚注引用这类非文字对象。

    不能用「删光 run 再塞一个新 run」的做法：段落里若有 OMML 公式或脚注引用，
    会被连对象一起删掉，而且删得悄无声息——文字看着对，公式没了。
    这里只改 w:t 文字节点，其余节点原样留着。

    范围取 p.runs，与 p.text 的口径一致：调用方看到什么文字，就改什么文字。
    """
    nodes = [t for r in p.runs for t in r._r.findall(qn("w:t"))]
    if not nodes:
        return p.add_run(new_text)

    nodes[0].text = new_text
    if new_text[:1].isspace() or new_text[-1:].isspace():
        nodes[0].set(XML_SPACE, "preserve")
    for node in nodes[1:]:
        node.text = ""
    return None


def set_run_xml(r, cn=None, latin=None, size=None):
    """在原始 XML 层给一个 w:r 套字体字号。

    脚注部件不在 python-docx 的对象模型里，拿不到 Run 对象，只能直接改元素。
    """
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    if cn:
        rf.set(qn("w:eastAsia"), cn)
    if latin:
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)
    if size:
        for tag in ("w:sz", "w:szCs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rpr.append(el)
            el.set(qn("w:val"), str(int(round(size * 2))))


def run_style(r):
    """读一个 w:r 当前的中文/西文字体与字号（半磅），用于判断是否真的需要改。"""
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return (None, None, None)
    rf = rpr.find(qn("w:rFonts"))
    sz = rpr.find(qn("w:sz"))
    return (rf.get(qn("w:eastAsia")) if rf is not None else None,
            rf.get(qn("w:ascii")) if rf is not None else None,
            int(sz.get(qn("w:val"))) if sz is not None and sz.get(qn("w:val")) else None)


def style_footnotes(doc, rules):
    """按规则集给脚注正文套字体字号，返回 (实际改动的条数, 目标描述)。

    期刊对脚注字号要求差异很大（这本 7.5pt、那本 9pt），而脚注部件独立于正文，
    正文套版碰不到它。不处理的话，成品正文合规、脚注字号还是稿件原来的——
    编辑一眼就看出来，报告却写着"已按规则集处理"。

    只统计**真的变了**的脚注：本来就合规的不进报告，否则"改为 7.5pt"这种话
    会出现在什么都没变的稿子上。
    """
    tp = (rules.get("typography") or {}).get("footnote") or {}
    size, cn = tp.get("size_pt"), tp.get("font_cn")
    latin = tp.get("font_latin") or "Times New Roman"
    if not (size or cn):
        return 0, ""

    changed = 0
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/footnotes"):
            continue
        try:
            root = etree.fromstring(rel.target_part.blob)
        except etree.XMLSyntaxError:
            continue

        touched = False
        for fn in root.findall(qn("w:footnote")):
            # id < 1 是分隔符与续行符，没有正文，不必动
            if int(fn.get(qn("w:id")) or "0") < 1:
                continue

            before = [run_style(r) for r in fn.iter(qn("w:r"))]
            for r in fn.iter(qn("w:r")):
                set_run_xml(r, cn=cn, latin=latin, size=size)
            after = [run_style(r) for r in fn.iter(qn("w:r"))]

            if before != after:
                changed += 1
                touched = True

        if touched:
            rel.target_part._blob = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True)

    desc = f"{cn or ''} {size:g}pt".strip() if size else (cn or "")
    return changed, desc


def replace_prefix(p, old_len, new_prefix):
    """只改段落开头的若干字符，其余文字与格式原样保留。

    参考文献条目常带斜体刊名，整段重写会把这些格式一起抹平，所以这里
    逐个 w:t 节点扣减，不重建段落。
    """
    remaining = old_len
    for r in p.runs:
        for t in r._r.findall(qn("w:t")):
            txt = t.text or ""
            if remaining <= 0:
                return
            take = min(remaining, len(txt))
            if take == len(txt) and remaining - take > 0:
                t.text = ""
            elif take > 0 and remaining == old_len:
                t.text = new_prefix + txt[take:]
            else:
                t.text = txt[take:]
            remaining -= take


REF_NUM = re.compile(
    r"^(?P<open>[\[［〔【(（])?\s*(?P<num>\d{1,3})\s*(?P<close>[\]］〕】)）])?(?P<tail>[.．、]?)")


def ref_number_prefix(text):
    """识别参考文献条目开头的编号前缀，返回 (原文前缀, 序号) 或 (None, None)。

    只认真正像编号的开头：带括号，或数字后跟句点/顿号/空白。否则
    「2016 年……」这种以年份开头的条目会被当成 201 号文献。
    """
    m = REF_NUM.match(text)
    if not m or not m.group("num"):
        return None, None
    bracketed = bool(m.group("open") and m.group("close"))
    tail = m.group("tail")
    after = text[m.end():m.end() + 1]
    if not (bracketed or tail or after.isspace()):
        return None, None
    return m.group(0), int(m.group("num"))


def format_ref_number(fmt, num, space_after=None):
    """按模板把序号渲染成目标刊的编号写法，如 [1] / 〔1〕 / 1．"""
    if not fmt:
        return None
    open_ = close = tail = ""
    m = re.search(r"([\[［〔【(（])\s*1\s*([\]］〕】)）])", fmt)
    if m:
        open_, close = m.group(1), m.group(2)
    elif fmt.strip() in ("1．", "1.", "1、"):
        tail = fmt.strip()[1:]
    out = f"{open_}{num}{close}{tail}"
    if space_after:
        out += " "
    return out


def apply_table(table, bp, size_pt, cn, latin):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for el in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(el)
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
    tblPr.append(borders)

    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(el)
            if ri == 0:
                tcb = OxmlElement("w:tcBorders")
                b = OxmlElement("w:bottom")
                b.set(qn("w:val"), "single")
                b.set(qn("w:sz"), str(int(bp["header_bottom"] * 8)))
                b.set(qn("w:space"), "0")
                b.set(qn("w:color"), "000000")
                tcb.append(b)
                tcPr.append(tcb)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run(run, cn=cn, latin=latin, size=size_pt)


def reorder_notes(elems, ir, rules, audit):
    """按目标刊的层序重排表注／图注。

    真实差异：国际金融研究是「资料来源 → 注」，金融研究是「注 → 资料来源」。
    稿件通常沿用其中一种，投另一本时必须换过来，否则读者看到的解释顺序是错的。
    """
    layers = {
        "table": [l["kind"] for l in rules["tables"]["note_layers"]],
        "figure": [l["kind"] for l in rules["figures"]["note_layers"]],
    }
    blocks = ir["blocks"]

    def kind_of(txt):
        if txt.startswith("数据来源"):
            return "data_source"
        if txt.startswith("资料来源"):
            return "source"
        if txt.startswith("注"):
            return "general"
        return "cell"

    for i, b in enumerate(blocks):
        if b["type"] not in ("table", "figure"):
            continue
        j = i + 1
        cap = None
        # 图题夹在插图和图注之间，收集图注时要先跳过它
        if b["type"] == "figure" and j < len(blocks) and blocks[j]["type"] == "figure_caption":
            cap = blocks[j]
            j += 1
        notes = []
        while j < len(blocks) and blocks[j]["type"] == "table_note":
            notes.append(blocks[j])
            j += 1
        if len(notes) < 2:
            continue
        rank = {k: n for n, k in enumerate(layers[b["type"]])}
        seq = [rank.get(kind_of(n["text"]), 99) for n in notes]
        if seq == sorted(seq):
            continue
        if audit.blocked(b["idx"], "表注／图注层序调整"):
            audit.add(b["idx"], b["type"], "表注／图注层序调整",
                      " / ".join(kind_of(n["text"]) for n in notes),
                      " / ".join(kind_of(n["text"]) for n in notes),
                      "已驳回：保持原层序", rejected=True,
                      rule=f"{b['type']}s.note_layers")
            continue
        if b["type"] == "table":
            anchor = elems[b["idx"]]._tbl
        else:
            anchor = elems[cap["idx"]]._p if cap else elems[b["idx"]]._p
        ordered = [n for _, n in sorted(zip(seq, notes), key=lambda x: x[0])]
        for n in notes:
            el = elems[n["idx"]]._p
            el.getparent().remove(el)
        cur = anchor
        for n in ordered:
            el = elems[n["idx"]]._p
            cur.addnext(el)
            cur = el
        audit.add(b["idx"], b["type"], "表注／图注层序调整",
                  " / ".join(kind_of(n["text"]) for n in notes),
                  " / ".join(kind_of(n["text"]) for n in ordered),
                  rule=f"{b['type']}s.note_layers")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ir")
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--reject", help="驳回清单 JSON：{\"ids\": [...]} 或 [...]，键为「块序号|动作」")
    a = ap.parse_args()

    ir = json.loads(Path(a.ir).read_text(encoding="utf-8"))
    rules = json.loads(Path(a.ruleset).read_text(encoding="utf-8"))
    src = Path(a.ir).with_suffix("")
    src_docx = Path(a.ir.replace(".ir.json", ".docx"))
    doc = Document(str(src_docx))

    rejects = set()
    if a.reject and Path(a.reject).is_file():
        raw = json.loads(Path(a.reject).read_text(encoding="utf-8"))
        rejects = set(raw.get("ids", []) if isinstance(raw, dict) else raw)

    audit = Audit(rejects)
    body = rules["typography"]["body"]
    cn, latin = body["font_cn"], body.get("font_latin", "Times New Roman")
    size = body["size_pt"]
    # 行距：固定值用 value_pt；倍数行距（如「1.2 倍」）用 equivalent_exact_pt 折算成固定值，
    # 由规则集给出折算依据，代码不猜系数
    ls = body["line_spacing"]
    line_pt = ls.get("value_pt") or ls.get("equivalent_exact_pt") or 20
    indent = body["first_line_indent_chars"]
    hd = rules["headings"]
    lv1 = hd["level1"]

    # 页面设置
    sec = doc.sections[0]
    sec.page_width = Mm(rules["page"]["size_mm"]["width"])
    sec.page_height = Mm(rules["page"]["size_mm"]["height"])
    for side in ("top", "bottom", "left", "right"):
        setattr(sec, f"{side}_margin", Mm(rules["page"]["margin_mm"][side]))
    mg = rules["page"]["margin_mm"]
    mg_txt = "，".join(f"{cn}{mg[side]:g}" for side, cn in
                       (("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右")))
    audit.add(-1, "page", "页面设置",
              note=f"纸张 {rules['page']['size']}，页边距 {mg_txt} mm",
              rule="page.margin_mm")

    # 建立 idx → 元素映射
    elems = {}
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            elems[i] = Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            elems[i] = Table(child, doc)

    # 一级标题起始号：introduction 编 1（多数刊）或 0（管理科学学报）
    sec_no = {1: hd.get("intro_number", 1) - 1, 2: 0, 3: 0}
    table_no = 0
    figure_no = 0
    fm_index = 0
    for blk in ir["blocks"]:
        p = elems.get(blk.get("idx"))
        if p is None:
            continue
        t = blk["type"]

        if t == "front_matter":
            fm_index += 1
            txt = p.text.strip()
            title_tp = rules["typography"]["title"]
            sub_tp = rules["typography"].get("subtitle") or {"font_cn": "黑体", "size_pt": 16}
            if fm_index == 1:
                sz = title_tp.get("size_pt", 22)
                restyle(p, cn=title_tp.get("font_cn", "黑体"), latin=latin, size=sz,
                        align="center", line_pt=round(sz * 1.5), bold=False,
                        before_pt=round(sz * 0.2), after_pt=round(sz * 0.25),
                        letter_spacing=title_tp.get("letter_spacing_pt"))
                audit.add(blk["idx"], t, "文章标题套版",
                          note=f"{sz}pt，行距 {round(sz * 1.5)}pt" + assumed(title_tp),
                          rule="typography.title")
            elif fm_index == 2 and txt.startswith(("——", "—", "-")):
                sz = sub_tp.get("size_pt", 16)
                restyle(p, cn=sub_tp.get("font_cn", "黑体"), latin=latin, size=sz,
                        align="center", line_pt=round(sz * 1.5), bold=False,
                        after_pt=round(sz * 0.2))
                audit.add(blk["idx"], t, "副标题套版", note=f"{sz}pt" + assumed(sub_tp),
                          rule="typography.subtitle~")
            else:
                # 作者署名：期刊可单独定字体字号（如四号楷体），未定则随正文
                au = rules["typography"].get("authors") or {}
                asz = au.get("size_pt", size)
                restyle(p, cn=au.get("font_cn", cn), latin=latin, size=asz,
                        align=au.get("align", "center"),
                        line_pt=max(line_pt, round(asz * 1.5)), bold=False,
                        before_pt=round(asz * 0.15), after_pt=round(asz * 0.9))
                audit.add(blk["idx"], t, "作者署名套版",
                          note=f"{au.get('font_cn', cn)} {asz}pt 居中",
                          # 该刊没单列 authors 样式时随正文，依据就是正文样式
                          rule="typography.authors" if au else "typography.body")
            continue

        if t == "table":
            table_no += 1
            apply_table(p, rules["tables"]["border_pt"], rules["tables"].get("font_size_pt", 9), cn, latin)
            audit.add(blk["idx"], t, "表格重设为三线表",
                      note=f"清除原有 {len(blk.get('borders') or {})} 处单元格级边框后重设",
                      rule="tables.border_pt")
            continue

        if isinstance(p, Table):
            # 公式常被排成「1 行 2 列」的伪表格：左边公式、右边编号。这种结构一改
            # 就散（拆表会把公式对象与编号断开），所以原样保留——但必须在报告里
            # 说明"这里没动"，否则用户看到公式没变，会以为工具漏了。
            audit.add(blk["idx"], t, "公式（伪表格）原样保留",
                      note="1 行表格排版的公式与编号不做改动，避免拆散公式对象与编号的对应关系")
            continue

        if t == "figure":
            # 图片段落必须解除固定行距：固定 18 磅的段落会把 12cm 高的内嵌图裁成一条。
            # 末尾的正文套版会给所有剩余块设固定行距，所以这里必须提前拦下、直接跳过。
            clear_para_formatting(p)
            p.alignment = ALIGN["center"]
            p.paragraph_format.line_spacing = 1.0
            audit.add(blk["idx"], t, "图片段解除固定行距",
                      note="固定行距会裁切内嵌图片，改为单倍行距并居中",
                      rule="figures")
            continue

        if t == "references_heading":
            # 文献表标题的写法逐刊不同（「参考文献」/「参 考 文 献」/「﹝参考文献﹞」），
            # 必须按规则集改写；版式默认同一级标题，可用 typography.references_heading 覆盖
            want = next((b.get("label") for b in rules["blocks"]
                         if b["type"] == "references" and b.get("label")), None)
            old = p.text
            if want and old.strip() != want:
                if audit.blocked(blk["idx"], "文献表标题归一"):
                    audit.add(blk["idx"], t, "文献表标题归一", old, old,
                              "已驳回：保留原写法", rejected=True,
                              rule="blocks:references/label")
                else:
                    replace_text(p, want)
                    audit.add(blk["idx"], t, "文献表标题归一", old, want, "仅改标题写法，条目未动",
                              rule="blocks:references/label")
            rh = rules["typography"].get("references_heading") or {}
            rh_size = rh.get("size_pt", lv1.get("size_pt") or size)
            restyle(p, cn=rh.get("font_cn", lv1.get("font_cn", cn)), latin=latin,
                    size=rh_size, align=rh.get("align", lv1.get("align", "center")),
                    line_pt=head_line_pt(lv1, rh_size, line_pt), bold=False)
            continue
        # 作者信息与缺失要件是用户补录后注入的块（apply_details 产出）。
        # 这里只把"语义块"映射到规则集里的样式键，字号字体一律查表，
        # 因此新增期刊仍然不需要改这段代码。
        if t in SEMANTIC:
            key, align, bold = SEMANTIC[t]
            tp = rules["typography"].get(key) or rules["typography"]["body"]
            sz = tp.get("size_pt", size)
            use_latin = latin
            if t == "english_title":
                ten = rules["typography"].get("title_en")
                if ten:            # 期刊若明确规定英文标题字号（如三号 Times），以它为准
                    tp = ten
                    sz = ten.get("size_pt", sz)
                    align = ten.get("align", align)
                    bold = ten.get("bold", bold)
                    use_latin = ten.get("font_latin", latin)
                else:              # 未声明时取标题与副标题中较小的字号，避免英文标题压过中文标题
                    sub = rules["typography"].get("subtitle") or {}
                    sz = min(sz, sub.get("size_pt", sz))
            # 英文标题与英文作者单独给段前距，否则会与中文关键词贴在一起
            before = after = None
            if t == "english_title":
                before, after = round(sz * 0.9), round(sz * 0.15)
            elif t == "english_authors":
                after = round(sz * 0.6)
            restyle(p, cn=tp.get("font_cn", cn), latin=use_latin, size=sz, align=align,
                    line_pt=line_pt, bold=bold, before_pt=before, after_pt=after)
            continue

        if t in ("heading1", "heading2", "heading3"):
            level = blk.get("level", 1)
            old = p.text
            numeric = str(hd.get("numbering", "")).startswith("numeric")
            # 剥前缀要认"稿子原来是什么编号"，不能按目标刊的体系去剥：
            # 中文稿投英文刊时，源码是「一、」，目标是「1.」，按目标剥会剥不掉
            core = old.strip()
            for rx in (NUM_PREFIX.get(level), NUM_PREFIX_NUM.get(level)):
                if rx is None:
                    continue
                stripped = rx.sub("", core).strip()
                if stripped != core:
                    core = stripped
                    break
            core_norm = re.sub(r"\s", "", core)

            if numeric:
                # 编号模板取自 levels：写「1」得「1 引言」（中国管理科学），
                # 写「1.」得「1. Introduction」（英文刊）。数字按层级依次替换。
                if level == 1:
                    sec_no[1] += 1
                    sec_no[2] = sec_no[3] = 0
                    nums = [sec_no[1]]
                elif level == 2:
                    sec_no[2] += 1
                    sec_no[3] = 0
                    nums = [sec_no[1], sec_no[2]]
                else:
                    sec_no[3] += 1
                    nums = [sec_no[1], sec_no[2], sec_no[3]]
                it = iter(nums)
                head = re.sub(r"\d+", lambda m: str(next(it)), hd["levels"][level - 1], count=len(nums))
                new = f"{head} {core}"
            elif level == 1:
                if core_norm in [re.sub(r"\s", "", s) for s in hd.get("unnumbered_sections", [])]:
                    new = hd.get("intro_title_text", core)
                else:
                    sec_no[1] += 1
                    sec_no[2] = sec_no[3] = 0          # 二级编号按一级标题重新起算
                    new = f"{CN_SEQ[sec_no[1] - 1]}{hd['levels'][0][1:]}{core}"
            elif level == 2:
                sec_no[2] += 1
                sec_no[3] = 0
                new = re.sub(r"[一二三四五六七八九十]+", CN_SEQ[sec_no[2] - 1], hd["levels"][1]) + core
            else:
                sec_no[3] += 1
                new = re.sub(r"\d+", str(sec_no[3]), hd["levels"][2]) + core

            if new != old:
                if audit.blocked(blk["idx"], "标题重新编号"):
                    audit.add(blk["idx"], t, "标题重新编号", old, new,
                              "已驳回：保留原编号", rejected=True,
                              rule="headings.numbering")
                else:
                    replace_text(p, new)
                    audit.add(blk["idx"], t, "标题重新编号", old, new, "仅改编号前缀，标题文字未动",
                              rule="headings.numbering")
            if level == 1:
                audit.sec = core  # 后续改动都归到这个一级标题下，审计里才能定位
            restyle(p, cn=lv1["font_cn"] if level == 1 else cn,
                    latin=latin, size=lv1.get("size_pt") if level == 1 else size,
                    align=lv1["align"] if level == 1 else "left",
                    line_pt=head_line_pt(lv1, (lv1.get("size_pt") or size), line_pt)
                    if level == 1 else line_pt,
                    bold=lv1.get("bold", False) if level == 1 else False)
            continue

        if t == "title":
            tp_t = rules["typography"]["title"]
            sz = tp_t.get("size_pt", 22)
            restyle(p, cn=tp_t.get("font_cn", "黑体"), latin=latin,
                    size=sz, align="center", line_pt=round(sz * 1.5), bold=False,
                    before_pt=round(sz * 0.2), after_pt=round(sz * 0.25),
                    letter_spacing=tp_t.get("letter_spacing_pt"))
            continue

        if t in ("abstract", "keywords"):
            want = next((b.get("label") for b in rules["blocks"]
                         if b["type"] == t and b.get("lang") == "zh" and b.get("label")), None)
            old = p.text
            if want:
                m = re.match(r"^(内容摘要|内容提要|摘\s*要|提\s*要|关\s*键\s*词)\s*[:：]", old)
                if m and not old.startswith(want):
                    new = want + old[m.end():]
                    if audit.blocked(blk["idx"], "标签归一"):
                        audit.add(blk["idx"], t, "标签归一", old, new,
                                  "已驳回：保留原标签", rejected=True,
                                  rule=f"blocks:{t}/zh/label")
                    else:
                        replace_text(p, new)
                        audit.add(blk["idx"], t, "标签归一", old, new, "仅改标签前缀，正文文字未动",
                                  rule=f"blocks:{t}/zh/label")
            # 关键词分隔符归一：各刊要求「；」「，」或空格，稿件里常是混的
            sep = next((b.get("separator") for b in rules["blocks"]
                        if b["type"] == "keywords" and b.get("lang") == "zh"), None)
            if t == "keywords" and sep and want and p.text.startswith(want):
                before_kw = p.text
                items = [x for x in re.split(r"[\s,，;；、]+", before_kw[len(want):]) if x]
                new_kw = want + sep.join(items)
                if new_kw != before_kw and len(items) >= 2:
                    if audit.blocked(blk["idx"], "关键词分隔符归一"):
                        audit.add(blk["idx"], t, "关键词分隔符归一", before_kw, new_kw,
                                  "已驳回：保留原分隔符", rejected=True,
                                  rule="blocks:keywords/zh/separator")
                    else:
                        replace_text(p, new_kw)
                        audit.add(blk["idx"], t, "关键词分隔符归一", before_kw, new_kw,
                                  f"改为「{sep}」，关键词数量与文字未动",
                                  rule="blocks:keywords/zh/separator")
            # 摘要与关键词常各自定字体（统计研究：提要宋体、关键词黑体），必须分别取
            key = ("abstract_en" if t == "abstract" else "keywords_en") if blk.get("lang") == "en" \
                else ("abstract" if t == "abstract" else "keywords")
            absp = rules["typography"].get(key) or rules["typography"].get("abstract") or {}
            abs_ls = (absp.get("line_spacing") or {}).get("value_pt") \
                or (absp.get("line_spacing") or {}).get("equivalent_exact_pt") or line_pt
            restyle(p, cn=absp.get("font_cn", cn), latin=latin,
                    size=absp.get("size_pt", size), align="both", line_pt=abs_ls, indent=indent)
            continue

        if t == "author_note":
            # 文末作者简介：标签一段、每人一段，字体字号单独声明
            an = rules["typography"].get("author_note") or {}
            sz = an.get("size_pt", size)
            an_ls = (an.get("line_spacing") or {}).get("value_pt") or round(sz * 1.6, 1)
            lab = (next((b.get("label") for b in rules["blocks"]
                         if b["type"] == "author_note"), "") or "作者简介").strip()
            if p.text.strip() == lab:
                restyle(p, cn=an.get("label_font_cn", "黑体"), latin=latin, size=sz,
                        align=an.get("heading_align", "left"), line_pt=an_ls,
                        indent=an.get("heading_indent_chars") or indent)
            else:
                restyle(p, cn=an.get("font_cn", cn), latin=latin, size=sz,
                        align="both", line_pt=an_ls, indent=indent)
            continue

        if t in ("reference_item",):
            rl = rules["citation"].get("reference_list", {})
            hang = rl.get("hanging_indent_chars")

            # 编号前缀按目标刊改写法（[1] / 〔1〕 / 1．）。只动前缀——条目正文里的
            # 作者、年份、标点一律不动，那属于改写内容，不归这个产品管。
            old_prefix, num = ref_number_prefix(p.text)
            if old_prefix is not None:
                new_prefix = format_ref_number(rl.get("number_format"), num,
                                               rl.get("space_after_number"))
                if new_prefix and new_prefix != old_prefix:
                    before_txt = p.text
                    if audit.blocked(blk["idx"], "文献条目编号归一"):
                        audit.add(blk["idx"], t, "文献条目编号归一", before_txt, before_txt,
                                  "已驳回：保留原编号写法", rejected=True,
                                  rule="citation.reference_list.number_format")
                    else:
                        replace_prefix(p, len(old_prefix), new_prefix)
                        audit.add(blk["idx"], t, "文献条目编号归一", before_txt, p.text,
                                  f"编号写法改为 {new_prefix.strip()}，条目内容未动",
                                  rule="citation.reference_list.number_format")

            ref_tp = rules["typography"].get("references") or {}
            r_size = ref_tp.get("size_pt", size)
            # 参考文献行距逐刊不同（经济研究 14.2、统计研究 12.3），不能沿用正文行距
            ref_ls = ref_tp.get("line_spacing") or {}
            ref_line = (ref_ls.get("value_pt") or ref_ls.get("equivalent_exact_pt")
                        or line_pt)
            restyle(p, cn=ref_tp.get("font_cn", cn), latin=latin, size=r_size,
                    align="both", line_pt=ref_line, indent=None, bold=False)
            if ref_tp.get("line_single"):
                # 规范要求「单倍行距」时不能沿用正文的固定行距，否则六号字会被撑开
                p.paragraph_format.line_spacing = 1.0
            if hang:
                set_hanging(p, hang)
                audit.add(blk["idx"], t, "参考文献悬挂缩进", after=f"{hang} 字符",
                          rule="citation.reference_list.hanging_indent_chars")
            elif rl.get("indent_chars"):
                # 部分期刊要求首行缩进而非悬挂缩进
                set_chars_indent(p, rl["indent_chars"])
            continue

        if t in ("table_caption", "figure_caption"):
            figure_no += t == "figure_caption"
            size_cap = rules["tables"].get("caption_size_pt", size)
            font_cap = rules["tables"].get("caption_font_cn", "黑体")
            # 「表1」与标题之间的间隔由规则集给出（一字空/半字空…），不再写死两空格
            gap_kind = (rules["tables"].get("caption_gap")
                        or (rules.get("figures") or {}).get("caption_gap") or "一字空")
            pad = CAPTION_GAP.get(gap_kind, "  ")
            gap_rule = ("tables.caption_gap" if "caption_gap" in rules["tables"]
                        else "figures.caption_gap" if "caption_gap" in (rules.get("figures") or {})
                        else "tables.caption_gap~")
            # 稿件里的表题常用制表符把「表1」和标题撑开，需按目标刊体例归一
            m = re.match(r"^(表|图)\s*(\d+)\s*(.*)$", p.text.strip())
            if m:
                old_cap = p.text
                new_cap = f"{m.group(1)}{m.group(2)}{pad}{m.group(3)}".rstrip()
                if new_cap != old_cap:
                    if audit.blocked(blk["idx"], "题注归一"):
                        audit.add(blk["idx"], t, "题注归一", old_cap, new_cap,
                                  "已驳回：保留原题注", rejected=True, rule=gap_rule)
                    else:
                        replace_text(p, new_cap)
                        audit.add(blk["idx"], t, "题注归一", old_cap, new_cap,
                                  f"编号与标题间隔改为「{gap_kind}」，题注文字未动", rule=gap_rule)
            restyle(p, cn=font_cap, latin=latin, size=size_cap, align="center",
                    line_pt=line_pt, bold=False)
            continue

        if t == "table_note":
            _nl = rules["tables"].get("note_layers") or [{}]
            note_sz = _nl[0].get("size_pt", 9)
            restyle(p, cn=cn, latin=latin, size=note_sz, align="left", line_pt=line_pt)
            continue

        # 独立公式段落：行距用「最小值」，不能用「固定值」。固定值会把分式、
        # 积分、上下标这类向上向下伸展的公式裁掉——和图片被裁是同一类事故。
        # 对齐方式保持原样，不强制居中：短段落里含行内公式时，解析器也会判为
        # equation，强行居中会把普通正文段落挪位置。
        if t == "equation":
            restyle(p, cn=cn, latin=latin, size=size, align=None, line_pt=line_pt,
                    line_rule="at_least")
            audit.add(blk["idx"], t, "公式段行距",
                      note=f"行距改为「最小值 {line_pt:g}pt」：固定值会裁掉分式与上下标",
                      rule="typography.body")
            continue

        # 正文、公式、伪表格、英文区等一律按正文字体处理。
        # 套版本身属 A 档（不动文字），但仍是「改动」，必须进报告——
        # 否则报告只记标签、编号这些小事，最显眼的正文重排反倒没有记录。
        before_fmt = fmt_snapshot(p)
        restyle(p, cn=cn, latin=latin, size=size, align="both", line_pt=line_pt, indent=indent)
        after_fmt = fmt_snapshot(p)
        if before_fmt != after_fmt:
            audit.add(blk["idx"], t, "正文套版",
                      note=fmt_delta(before_fmt, after_fmt,
                                     f"{cn} {size:g}pt，行距 {line_pt:g}pt"),
                      rule="typography.body")

    out = Path(a.out)
    reorder_notes(elems, ir, rules, audit)

    fn_count, fn_desc = style_footnotes(doc, rules)
    if fn_count:
        audit.add(-2, "footnote", "脚注正文套版",
                  note=f"{fn_count} 条脚注改为 {fn_desc}",
                  rule="typography.footnote")

    doc.save(str(out))

    audit_path = out.with_suffix(".audit.json")
    audit_path.write_text(json.dumps({"source": ir["source"], "target": rules["journal"]["name_zh"],
                                      "items": audit.items,
                                      "assumptions": rules.get("assumptions") or [],
                                      "rejected": sorted(rejects)},
                                     ensure_ascii=False, indent=2),
                          encoding="utf-8")

    print(f"已生成: {out}")
    print(f"审计条目: {len(audit.items)} 条 → {audit_path.name}")
    if rejects:
        print(f"已驳回: {len(rejects)} 条改动（保留原文）")
    kinds = {}
    for it in audit.items:
        kinds[it["action"]] = kinds.get(it["action"], 0) + 1
    for k, v in sorted(kinds.items(), key=lambda x: -x[1]):
        print(f"   {v:4d}  {k}")


if __name__ == "__main__":
    main()
