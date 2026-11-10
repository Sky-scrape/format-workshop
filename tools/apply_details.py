"""把用户补录的作者信息与缺失要件注入稿件，产出"信息填充完成后的快照"。

DESIGN §10.5 规定：审计的比对基线是个人信息填充完成后的文档快照，不是原始上传文件。
所以这一步必须发生在 render_from_ir 之前，注入的内容不计入格式变更。

只做两件事：
  1. 正文前置区插入缺失要件（单位、分类号、英文标题/作者/单位/摘要/关键词）
  2. 首页脚注写入基金项目、作者简介、通讯作者

一段文字都不改写、不删减；插入的段落只带文本，版式全部交给渲染器按规则集套。

用法：
  python tools/apply_details.py in.docx in.ir.json details.json rules.json \
      --docx-out filled.docx --ir-out filled.ir.json
"""

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_meta import sanitize_docx

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 展示顺序即插入顺序。字段 → (块类型, 语言, 锚点候选)
# 锚点写成 (类型, 语言)，因为同一类型可能同时存在中英两套（如分类号）
FIELD_SPEC = [
    ("affiliation_zh", "affiliations", "zh", [("front_matter", "zh")]),
    ("classification", "classification", "zh", [("keywords", "zh"), ("abstract", "zh")]),
    ("english_title", "english_title", "en",
     [("classification", "zh"), ("keywords", "zh"), ("abstract", "zh")]),
    ("english_authors", "english_authors", "en",
     [("english_title", "en"), ("front_matter", "zh")]),
    ("english_affiliation", "english_affiliations", "en",
     [("english_authors", "en"), ("english_title", "en")]),
    ("english_abstract", "english_abstract", "en",
     [("english_affiliations", "en"), ("english_title", "en")]),
    ("english_keywords", "english_keywords", "en",
     [("english_abstract", "en"), ("english_title", "en")]),
    ("english_classification", "classification", "en",
     [("english_keywords", "en"), ("english_abstract", "en"), ("english_title", "en")]),
]

FOOTER_FIELDS = [("funding", "基金项目："), ("author_bio", "作者简介："),
                 ("corresponding", "通讯作者：")]
# 首页脚注字段 → personal_fields 里对应的字段 id，用来判断该刊是否要求改放别处
FOOTER_PLACEMENT = {"funding": "funding_info", "author_bio": "author_bio",
                    "corresponding": "corresponding_author"}

# 注入块 → 规则集 blocks 里的声明，用来取该刊规定的标签（Abstract：/ Key words：/ JEL分类号：）
LABEL_KEY = {
    ("affiliations", "zh"): ("affiliations", None),
    ("classification", "zh"): ("classification_code", None),
    ("classification", "en"): ("classification_code", "en"),
    ("english_title", "en"): ("title", "en"),
    ("english_authors", "en"): ("authors", "en"),
    ("english_affiliations", "en"): ("affiliations", "en"),
    ("english_abstract", "en"): ("abstract", "en"),
    ("english_keywords", "en"): ("keywords", "en"),
}

def blang(b):
    """IR 块的语言：english_* 一律视为英文，其余看 lang 字段，默认中文。"""
    return "en" if b["type"].startswith("english_") or b.get("lang") == "en" else "zh"


def label_for(rules, btype, lang):
    """标签属于期刊规则集，不在代码里硬写；没声明就不加前缀。"""
    key = LABEL_KEY.get((btype, lang))
    if not key:
        return ""
    for blk in rules.get("blocks", []):
        if blk.get("type") == key[0] and (blk.get("lang") or None) == key[1]:
            return blk.get("label") or ""
    return ""


def with_label(text, label):
    """补上期刊要求的标签，同时避免用户已经自己写了标签造成重复。"""
    if not label:
        return text
    core = re.sub(r"[:：]\s*$", "", label).strip()
    if text.lower().startswith(core.lower()):
        text = text[len(core):].lstrip(":： ")
    if "JEL" in label:
        text = re.sub(r"^JEL\s*[:：]?\s*", "", text, flags=re.I)
    return label + text


def add_paragraph_after(doc, anchor_el, text):
    """在 anchor 之后插入一个新段落并返回它。

    python-docx 只有 insert_paragraph_before，这里先造一个干净段落再搬位置；
    版式留给渲染器统一处理，此处不设任何格式。
    """
    p = doc.add_paragraph()
    anchor_el.addnext(p._element)
    p.add_run(text)
    return p


AUTHOR_PAGE_STYLE = "作者信息页"

# 「作者：张三　　单位：××大学」——模板首页常把两个字段塞在同一行
AUTHOR_SPLIT = re.compile(
    r"^作\s*者\s*[:：]?\s*(?P<name>.*?)"
    r"(?:(?:\u3000|\s{2,})+\s*单\s*位\s*[:：]\s*(?P<aff>.*))?$")

# 前置区模板里的占位符：「单位：[单位全称，城市，邮编]」
PLACEHOLDER = {
    "affiliation": re.compile(r"(单\s*位\s*[:：]\s*)[\[【［][^\]】］]*[\]】］]"),
    "author_name": re.compile(r"(作\s*者\s*[:：]\s*)[\[【［][^\]】］]*[\]】］]"),
}


def split_author_line(text):
    """把「作者：张三　　单位：××大学」拆成 (姓名, 单位)。

    整行塞进「作者姓名：」会变成双标签的「作者姓名：作者：张三……」，
    首页第一眼就露馅。
    """
    t = (text or "").strip()
    m = AUTHOR_SPLIT.match(t)
    if not m:
        return t, ""
    return (m.group("name") or "").strip(), (m.group("aff") or "").strip()


def fill_placeholders(doc, blocks, block_el, details):
    """把前置区模板里的占位符就地换成补录值，返回替换成功的字段名。

    不换的话，成品会一边留着「单位：[单位全称，城市，邮编]」、一边在下一行
    印上真单位——同一条信息出现两次，比完全不填更糟。
    只替换用户明确填过的字段，没填的占位符原样保留，不擅自删字。
    """
    vals = {
        "affiliation": (details.get("affiliation_zh") or details.get("affiliation") or "").strip(),
        "author_name": (details.get("author_name") or "").strip(),
    }
    if not any(vals.values()):
        return []

    done = set()
    for b in blocks:
        if b["type"] != "front_matter":
            continue
        el = block_el.get(id(b))
        if el is None:
            continue
        p = Paragraph(el, doc)
        txt = p.text
        new = txt
        candidates = set()
        for key, val in vals.items():
            if val:
                new, n = PLACEHOLDER[key].subn(lambda m, v=val: m.group(1) + v, new)
                if n:
                    candidates.add(key)
        if new == txt:
            continue
        # 物理文字节点按段落全部后代 run 收集（不下钻文本框/删除内容）：
        # p.runs 只取直接子级 w:r，占位符整段在 w:hyperlink 或域结果里时
        # 它一个都看不见——只按它写 docx 会漏写，而 IR 照更新，两者就此
        # 分叉（体检按 IR 判「已具备」，成品却还留着占位符）。此处口径与
        # render_from_ir.para_runs 一致。
        nodes = [t for r in p._p.xpath(".//w:r[not(ancestor::w:drawing) and "
                                       "not(ancestor::w:pict) and not(ancestor::w:object)"
                                       " and not(ancestor::w:del)]")
                 for t in r.findall(qn("w:t"))]
        if not nodes:
            # 物理节点一个都没有：不写 docx，也不更新 IR、不算填好——
            # 兜底注入（existing 不再添加）会补上这条信息，IR 与成品一致
            continue
        nodes[0].text = new
        nodes[0].set(XML_SPACE, "preserve")
        for node in nodes[1:]:
            node.text = ""
        b["text"] = new
        done |= candidates
    return done


def build_author_page(doc, ir, rules, details):
    """按规则集要求的 8 类要素生成「作者信息页」，插在正文之前并分页。

    只放字段值，不做任何版式：字体字号交给渲染器（该页段落带专用样式名，
    匿名稿据此整页删除，不必猜哪些段落属于这一页）。
    """
    cfg = (rules.get("submission") or {}).get("author_page") or {}
    if not cfg.get("separate_page"):
        return {"n": 0, "keys": []}

    fm = [b.get("text", "").strip() for b in ir["blocks"] if b["type"] == "front_matter"]
    title = fm[0] if fm else ""
    author_line = next((t for t in fm[1:] if not t.startswith(("——", "—", "-"))), "")
    author_name, author_aff = split_author_line(author_line)
    abstract = next((b.get("text", "").strip() for b in ir["blocks"] if b["type"] == "abstract"), "")
    keywords = next((b.get("text", "").strip() for b in ir["blocks"] if b["type"] == "keywords"), "")

    values = {
        "paper_title": title,
        "paper_title_en": (details.get("english_title") or "").strip(),
        "author_name": author_name,
        "authors": author_name,
        "affiliation": (details.get("affiliation_zh") or author_aff or "").strip(),
        "affiliations": (details.get("affiliation_zh") or author_aff or "").strip(),
        "abstract": abstract,
        "keywords": keywords,
        "author_note": (details.get("author_bio") or "").strip(),
        "author_bio": (details.get("author_bio") or "").strip(),
        "contact": (details.get("corresponding") or "").strip(),
        "corresponding_author": (details.get("corresponding") or "").strip(),
        "funding_info": (details.get("funding") or "").strip(),
        "funding": (details.get("funding") or "").strip(),
    }
    labels = {}
    for f in rules.get("personal_fields") or []:
        labels[f["id"]] = f.get("label_zh") or f["id"]
    # 表单里按字段 id 存下来的值（性别、职称、身份证号等）优先用；上面那几张是兜底
    for k, v in details.items():
        if isinstance(v, str) and v.strip():
            values.setdefault(k, v.strip())

    order = [x for x in (cfg.get("fields") or []) if values.get(x)]
    if not order:                      # 规则集没给字段清单时，退回到有值的几项
        order = [k for k in ("paper_title", "author_name", "affiliation",
                             "abstract", "keywords", "author_note", "contact", "funding_info")
                 if values.get(k)]
    if not order:
        return {"n": 0, "keys": []}

    try:
        style = doc.styles[AUTHOR_PAGE_STYLE]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        style = doc.styles.add_style(AUTHOR_PAGE_STYLE, WD_STYLE_TYPE.PARAGRAPH)

    first = doc.element.body[0]
    made = 0
    fp = rules.get("typography", {}).get("body", {})
    size_pt = fp.get("size_pt", 10.5)
    font_cn = fp.get("font_cn", "宋体")
    for i, key in enumerate(order):
        p = doc.add_paragraph(style=style)
        label = labels.get(key, "")
        r1 = p.add_run(f"{label}：" if label else "")
        r2 = p.add_run(values[key])
        for r in (r1, r2):
            r.font.size = Pt(size_pt)
            r._element.get_or_add_rPr()
            rf = r._element.rPr.get_or_add_rFonts()
            rf.set(qn("w:eastAsia"), font_cn)
        first.addprevious(p._element)
        made += 1
        if i == len(order) - 1:        # 分页符放在本页最后一段，删页时一并消失
            from docx.enum.text import WD_BREAK
            p.add_run().add_break(WD_BREAK.PAGE)
    return {"n": made, "keys": order}



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("ir")
    ap.add_argument("details")
    ap.add_argument("rules")
    ap.add_argument("--docx-out", required=True)
    ap.add_argument("--ir-out", required=True)
    a = ap.parse_args()

    doc = Document(a.docx)
    ir = json.loads(Path(a.ir).read_text(encoding="utf-8"))
    details = json.loads(Path(a.details).read_text(encoding="utf-8"))
    rules = json.loads(Path(a.rules).read_text(encoding="utf-8"))

    blocks = ir["blocks"]
    children0 = list(doc.element.body)

    # 先把"老块 → 元素对象"固定下来：插入会移动序号，元素对象不会变
    block_el = {}
    for b in blocks:
        i = b["idx"]
        if 0 <= i < len(children0):
            block_el[id(b)] = children0[i]
    type_last_el = {}
    for b in blocks:
        if id(b) in block_el:
            type_last_el[(b["type"], blang(b))] = block_el[id(b)]

    existing = {(b["type"], blang(b)) for b in blocks}
    inserted = []
    tail = {}  # 同一锚点连续插入时，后一个接在前一个之后，保持字段顺序

    # 前置区模板占位符能就地填的，就不再另起一行重复注入
    filled = fill_placeholders(doc, blocks, block_el, details)
    if "affiliation" in filled:
        existing.add(("affiliations", "zh"))
    if "author_name" in filled:
        existing.add(("authors", "zh"))

    # 英文区排在正文之后（经济研究、世界经济）时，注入点应落在参考文献之后，
    # 而不是紧跟中文关键词——否则成品布局就与该刊不符。位置由规则集决定，不靠猜。
    def _blk_idx(t, lang=None):
        for i, b in enumerate(rules.get("blocks") or []):
            if b["type"] == t and (lang is None or (b.get("lang") or None) == lang):
                return i
        return 10 ** 6

    en_first = min([_blk_idx(t, "en") for t in ("title", "authors", "abstract", "keywords")]
                   + [10 ** 6])
    en_at_tail = en_first > _blk_idx("body")
    spec = []
    for field, btype, lang, anchors in FIELD_SPEC:
        # 只有英文标题需要换锚点，其余英文块靠原有链式锚点（标题→作者→单位→摘要→关键词）
        # 依次接续；若把所有英文锚点都改成文末，链就断了，Summary 会跑到单位前面
        if en_at_tail and field == "english_title":
            anchors = [("reference_item", "zh"), ("references_heading", "zh"),
                       ("author_note", "zh")] + list(anchors)
        spec.append((field, btype, lang, anchors))

    for field, btype, lang, anchors in spec:
        text = (details.get(field) or "").strip()
        if not text or (btype, lang) in existing:
            continue
        anchor_el = None
        for at in anchors:
            if at in type_last_el:
                anchor_el = type_last_el[at]
                break
        if anchor_el is None:
            print(f"  [skip] {field}: 找不到可插入的位置", file=sys.stderr)
            continue
        anchor_el = tail.get(id(anchor_el), anchor_el)

        text = with_label(text, label_for(rules, btype, lang))
        p = add_paragraph_after(doc, anchor_el, text)
        tail[id(anchor_el)] = p._element
        type_last_el[(btype, lang)] = p._element   # 后面的字段可以以它为新锚点
        inserted.append({"field": field, "type": btype, "text": p.text,
                         "lang": lang, "el": p._element})

    # ---- 作者信息页：该刊要求单独成页时，在正文之前生成 ----
    page = build_author_page(doc, ir, rules, details)
    page_items, page_keys = page["n"], set(page["keys"])

    # ---- 文末作者简介：该刊把简介放在参考文献之后（如统计研究），不进首页脚注 ----
    # placement 由 personal_fields 声明，代码不猜
    pf = {f["id"]: f for f in rules.get("personal_fields") or []}
    back_note = (pf.get("author_bio") or {}).get("placement") == "back_matter"
    if back_note and (details.get("author_bio") or "").strip() and ("author_note", "zh") not in existing:
        note_blk = next((b for b in rules.get("blocks", []) if b["type"] == "author_note"), None)
        if note_blk:
            anchor = type_last_el.get(("reference_item", "zh"))
            if anchor is None:
                anchor = type_last_el.get(("references_heading", "zh"))
            if anchor is None:
                # 兜底取最后一个块级元素：body 末尾的 sectPr 是节属性不是正文，
                # addnext 到它后面会产出「节属性之后还有正文」的非法结构
                # （ECMA-376 CT_Body 里 sectPr 必须居末），Word/WPS 重存时要
                # 归一化搬动段落。IR 类型名是 references_heading（parse_docx
                # 产出），不存在 "references" 类型——旧锚点键永远落空。
                block_els = [el for el in doc.element.body
                             if el.tag in (qn("w:p"), qn("w:tbl"))]
                anchor = block_els[-1] if block_els else None
                if anchor is None:
                    anchor = doc.element.body[-1] if len(doc.element.body) else None
            p1 = add_paragraph_after(doc, anchor, note_blk.get("label") or "作者简介")
            p2 = add_paragraph_after(doc, p1._element, details["author_bio"].strip())
            inserted.append({"field": "author_note", "type": "author_note",
                             "text": p1.text, "lang": "zh", "el": p1._element})
            inserted.append({"field": "author_bio", "type": "author_note",
                             "text": p2.text, "lang": "zh", "el": p2._element})
            type_last_el[("author_note", "zh")] = p2._element

    # ---- 首页脚注：基金 / 作者简介 / 通讯作者 ----
    # 作者页上已经写过的项，首页脚注不再重复一遍
    ON_PAGE = {"funding": {"funding_info", "funding"},
               "author_bio": {"author_note", "author_bio"},
               "corresponding": {"contact", "corresponding_author"}}
    footer_lines = []
    for key, prefix in FOOTER_FIELDS:
        v = (details.get(key) or "").strip()
        if not v or (page_keys & ON_PAGE.get(key, set())):
            continue
        # 规则集把该项标为文末（作者页/文末简介）时，不再往首页脚注塞一遍
        if FOOTER_PLACEMENT.get(key) and \
                (pf.get(FOOTER_PLACEMENT[key]) or {}).get("placement") == "back_matter":
            continue
        footer_lines.append(f"{prefix}{v}")
    if footer_lines:
        sec = doc.sections[0]
        sec.different_first_page_header_footer = True
        foot = sec.first_page_footer
        # 原稿首页脚注可能已有内容（收稿日期等），只在空段落上写，不动原有文字
        if foot.paragraphs and not foot.paragraphs[0].text.strip():
            para = foot.paragraphs[0]
        else:
            para = foot.add_paragraph()
        for r in list(para.runs):
            r._element.getparent().remove(r._element)
        size = rules.get("typography", {}).get("footnote", {}).get("size_pt", 9)
        for i, line in enumerate(footer_lines):
            if i:
                para.add_run().add_break()
            para.add_run(line).font.size = Pt(size)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.save(a.docx_out)
    # 中间快照也会被后续工序打开和另存；先把原稿的 Word/WPS 属性清掉，
    # 避免它们一路带到交付文件里。
    sanitize_docx(a.docx_out)

    # ---- 重排 IR：idx 是 body 子元素序号，插入后必须整体重算 ----
    children1 = list(doc.element.body)
    pos = {id(el): i for i, el in enumerate(children1)}
    new_blocks = []
    dropped = []
    for b in blocks:
        el = block_el.get(id(b))
        if el is None:
            # idx 越出 body 子元素范围（IR 与 docx 失配）的块会整块从
            # filled.ir.json 里消失：致谢、基金这类块丢了，匿名工具和
            # 合规报告就永久看不见它们——匿名稿会留着致谢段且审计零痕迹
            # （F4-8）。注入失败有 [skip] 告警，这里同样不能静默。
            dropped.append(f"{b['type']}：「{(b.get('text') or '')[:20]}」(idx={b['idx']})")
            continue
        nb = dict(b)
        nb["idx"] = pos[id(el)]
        new_blocks.append(nb)
    for item in inserted:
        new_blocks.append({"idx": pos[id(item["el"])], "type": item["type"],
                           "text": item["text"], "lang": item["lang"],
                           "injected": True})
    new_blocks.sort(key=lambda x: x["idx"])

    ir2 = dict(ir)
    ir2["blocks"] = new_blocks
    Path(a.ir_out).write_text(json.dumps(ir2, ensure_ascii=False, indent=2), encoding="utf-8")

    names = "、".join(i["field"] for i in inserted) or "无"
    print(f"已注入 {len(inserted)} 项正文要件：{names}")
    if page_items:
        print(f"作者信息页 {page_items} 行（独立成页，匿名稿会整页删除）")
    if footer_lines:
        print(f"首页脚注 {len(footer_lines)} 行：" + "；".join(l[:14] for l in footer_lines))
    if dropped:
        print(f"[warn] {len(dropped)} 个 IR 块的 idx 越出 body 子元素范围，"
              "已从输出 IR 丢弃——下游（匿名/合规报告）看不到这些块，"
              "请检查 IR 与 docx 是否同一次解析产出：" + "；".join(dropped),
              file=sys.stderr)


if __name__ == "__main__":
    main()
