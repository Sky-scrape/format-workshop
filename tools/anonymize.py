"""由成品稿派生一份「投稿用匿名稿」（DESIGN §11.4 的双版本交付）。

署名稿保持完整、供留档与校样；匿名稿只删身份信息，不动任何正文文字。
两份文件同源同版式，差别仅在身份信息，且删了什么逐条写进审计报告。

只删三类东西：
  1. 前置区的作者署名行（标题与副标题保留）
  2. 首页脚注 / 脚注里的作者简介、基金项目、通讯作者
  3. 「致谢」小节（若有）

用法：
    python tools/anonymize.py 成品.docx 稿件.ir.json 规则集.json -o 匿名稿.docx
"""

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from docx.table import Table
from docx.text.paragraph import Paragraph

# 一条脚注同时命中「单位类」与「身份类」词才判为作者信息：
# 只写「数据来自北京大学」的注不会被误删
UNIT_RE = re.compile(r"(大学|学院|研究院|研究所|商学院|经济学院|管理学院|系\b)")
IDENT_RE = re.compile(r"(邮政编码|邮编|通讯地址|通信地址|电子信箱|邮箱|E-?mail|@|"
                      r"教授|副教授|讲师|博士生|硕士生|研究方向|通讯作者|基金|项目编号|批准号)")
ACK_RE = re.compile(r"^(致\s*谢|鸣\s*谢|Acknowledg(e)?ment)")


def _iter_children(doc):
    """按 body 子元素顺序产出 (序号, 元素)，序号与 IR 的 idx 对齐。"""
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            yield i, Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield i, Table(child, doc)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("ir")
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    rules = json.loads(Path(a.ruleset).read_text(encoding="utf-8"))
    ir = json.loads(Path(a.ir).read_text(encoding="utf-8"))
    sub = rules.get("submission") or {}
    if not sub.get("anonymous"):
        raise SystemExit("该刊不要求匿名，无需生成匿名稿")

    doc = Document(a.docx)
    by_idx = dict(_iter_children(doc))
    removed = []

    # ---- 0. 作者信息页：整页删除（该页段落带专用样式，不必猜范围）----
    page_paras = [p for p in doc.paragraphs if p.style is not None and p.style.name == "作者信息页"]
    if page_paras:
        removed.append(f"作者信息页（{len(page_paras)} 行）")
        for p in page_paras:
            p._p.getparent().remove(p._p)

    # ---- 1. 前置区：保留标题与副标题，其余（作者行、单位行）删掉 ----
    fm_seen = 0
    for b in ir["blocks"]:
        if b["type"] != "front_matter":
            continue
        fm_seen += 1
        el = by_idx.get(b["idx"])
        if el is None or not isinstance(el, Paragraph):
            continue
        txt = el.text.strip()
        is_subtitle = txt.startswith(("——", "—", "-"))
        if fm_seen == 1 or is_subtitle:
            continue
        if txt:
            removed.append(f"作者署名：{txt[:40]}")
        el._p.getparent().remove(el._p)

    # ---- 2. 补录注入的单位/作者简介/基金块 ----
    # 作者、单位、简介、基金四类本身就是身份信息，凡要求匿名一律删除；
    # 致谢是否删除由规则集声明（有的刊允许保留，只要不含致谢人姓名）
    #
    # 安全优先：规则集没声明"作者""单位"时也删。漏删是身份泄露（不可挽回），
    # 多删是基金信息缺失（可从署名稿拿回来，而且署名稿本来就完整保留）。
    # 但多删这件事必须在报告里说清，不能悄悄多删。
    listed = set(sub.get("anonymize_blocks") or [])
    SAFETY_TYPES = {"authors", "affiliations", "author_note", "funding_statement",
                    "corresponding_author"}
    drop_types = SAFETY_TYPES | listed

    # 规则集声明了、工具却不认识的类型：记下来让用户人工检查，
    # 不能静默跳过——静默跳过等于身份信息留在稿子里
    unknown = sorted(listed - SAFETY_TYPES - {"acknowledgement"})
    for t in unknown:
        removed.append(f"未识别：规则集要求匿名化「{t}」，本工具不认识该类型，请人工检查")

    # 致谢小节：删标题还要删标题下面的正文，否则会留下一段孤立的「感谢……」——
    # 既泄露了致谢内容，又让成品多出一段没头没尾的文字
    ack_idx = set()
    for pos, b in enumerate(ir["blocks"]):
        head = (b["type"] == "acknowledgement"
                or ACK_RE.match((b.get("text") or "").strip()))
        if not head or b["type"] not in drop_types:
            continue
        ack_idx.add(b["idx"])
        for nxt in ir["blocks"][pos + 1:]:
            if (nxt["type"].startswith("heading")
                    or nxt["type"] in ("references_heading", "appendix_heading",
                                       "acknowledgement")):
                break
            if nxt["type"] in ("body", "table", "table_note", "table_caption",
                               "figure", "figure_caption", "equation"):
                ack_idx.add(nxt["idx"])

    for b in ir["blocks"]:
        el = by_idx.get(b["idx"])
        if b["idx"] in ack_idx:
            if isinstance(el, Paragraph):
                removed.append(f"致谢小节：{el.text.strip()[:30]}")
                el._p.getparent().remove(el._p)
            continue
        if b["type"] not in drop_types or b["type"] == "front_matter":
            continue
        if el is None or not isinstance(el, Paragraph):
            continue
        if b["type"] == "acknowledgement" or ACK_RE.match(el.text.strip()):
            removed.append(f"致谢小节：{el.text.strip()[:30]}")
        else:
            # 规则集没要求、但为确保匿名一并删的，标注清楚
            mark = "" if b["type"] in listed else "（规则集未声明，为确保匿名一并移除）"
            removed.append(f"{b['type']}{mark}：{el.text.strip()[:40]}")
        el._p.getparent().remove(el._p)

    # ---- 3. 首页脚注（补录写入的基金 / 作者简介 / 通讯作者） ----
    sec = doc.sections[0]
    foot = sec.first_page_footer
    for p in list(foot.paragraphs):
        if p.text.strip():
            removed.append(f"首页脚注：{p.text.strip()[:40]}")
            for r in list(p.runs):
                r._element.getparent().remove(r._element)

    # ---- 4. 正文脚注里的作者信息（原稿自带的作者简介常在这里）----
    # 脚注部件不在 python-docx 的对象模型里，只能按关系取到原始 XML 再改
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/footnotes"):
            continue
        from lxml import etree
        root = etree.fromstring(rel.target_part.blob)
        killed = set()
        for fn in root.findall(qn("w:footnote")):
            text = "".join(t.text or "" for t in fn.iter(qn("w:t")))
            if not text.strip():
                continue
            if UNIT_RE.search(text) and IDENT_RE.search(text):
                removed.append(f"脚注：{text.strip()[:40]}")
                killed.add(fn.get(qn("w:id")))
                for t in fn.iter(qn("w:t")):
                    t.text = ""
        if killed:
            rel.target_part._blob = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True)
            # 正文里的脚注引用标记一并摘掉，否则成品会留下一个空脚注号
            for ref in list(doc.element.body.iter(qn("w:footnoteReference"))):
                if ref.get(qn("w:id")) in killed:
                    run = ref.getparent()
                    run.getparent().remove(run)

    doc.save(a.out)
    print(json.dumps({"removed": removed}, ensure_ascii=False))


if __name__ == "__main__":
    main()
