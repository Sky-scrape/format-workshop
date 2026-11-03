"""排序说明与嵌套内容控件的端到端回归钉。

- cn-first-by-alphabet-then-en-by-alphabet 的中文键使用 GB 近似拼音序，
  张三（zhang）排在周一（zhou）前；SORT_NOTE 必须说明实际使用的 GB 序，
  不得再沿用过时的 Unicode 码位序描述。
- parse_docx 与 render_from_ir 递归展开块级 sdt：外层与内嵌段落都进 IR、
  按文档序套版，文字原样保留；只有真正未识别或取空的控件才计数。
  段落/表格内部的行内 sdt 文字已随 para_text 处理，不重复计数。

夹具姿势：python-docx 造稿、跑 parse/render 用 conftest.run_tool（参数
列表直传、不经 shell，与 test_citation_convert.py 相同的子进程姿势）。

注：工单1 点名的 jryj（金融研究）是唯一声明该 scheme 的规则集，但
rules/journal 里入库的只有 cjyj/gjjryj/tjyj（conftest 守卫），所以这里
用 tjyj 副本把 sort 改成同一个声明来钉渲染器行为——note 映射按 scheme
字符串取值，与来自哪本刊无关。
"""

import json
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from conftest import RULES, TOOLS, run_tool

# 名字取自证据复现：拼音序 张三(zhang) 在 周一(zhou) 前；码位序恰好相反
# （周一 U+5468 < 张三 U+5F20）。
REF_ZHANG = "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10."
REF_ZHOU = "[2] 周一. 融资约束的度量[J]. 金融研究, 2021, 76(2): 20-40."

# 与 test_fix_render_core.py 同一套可渲染模板（jryj 自带的 templates 是
# 无占位符的著录样例，条目重排不会发生、排序分支走不到——复现需可渲染模板）
TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
    "en_journal": "{authors}，{year}，“{title}”，{journal}{volume_part}{page_part}.",
    "en_book": "{authors}，{year}，“{title}”，{publisher}.",
    "en_other": "{authors}，{year}，“{title}”，{source}.",
}

SORT_SCHEME = "cn-first-by-alphabet-then-en-by-alphabet"


def alphabet_rules(tmp_path):
    """作者-年份制规则集，sort 覆盖为字母类声明（工单1 的 scheme）。"""
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["citation"]["system"] = "author-year"
    rl = rules["citation"].setdefault("reference_list", {})
    rl["templates"] = TEMPLATES
    rl["sort"] = SORT_SCHEME
    p = tmp_path / "alphabet_rules.json"
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return p


def make_src(tmp_path, name="src"):
    """文献表故意按 周一[1]、张三[2] 排列，以触发拼音序重排与审计。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHOU.replace("[2]", "[1]", 1))
    doc.add_paragraph(REF_ZHANG.replace("[1]", "[2]", 1))
    p = tmp_path / f"{name}.docx"
    doc.save(str(p))
    return p


def parse(tmp_path, docx):
    ir = docx.with_suffix(".ir.json")
    run_tool("parse_docx.py", docx, "-o", ir)
    return ir


def load_ir(ir):
    return json.loads(Path(ir).read_text(encoding="utf-8"))


def render(tmp_path, ir, rules_path, citation=False):
    out = Path(ir).with_suffix(".out.docx")
    args = [ir, rules_path, "-o", out]
    if citation:
        args.append("--citation")
    r = run_tool("render_from_ir.py", *args)
    return r, out


def audit_items(out):
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    return [i for i in audit["items"] if i["action"] == "文献表排序"]


def docx_xml_text(out):
    """成品 word/document.xml 的全部文字（含 sdt 内段落，.paragraphs 取不到）。

    解析的是渲染器自己刚产出的 XML，仍按仓库口径先拒 DOCTYPE/ENTITY
    （同 conftest.layout_signature / render_from_ir.footnote_texts 的姿势）。
    """
    with zipfile.ZipFile(str(out)) as z:
        raw = z.read("word/document.xml")
    if b"<!DOCTYPE" in raw or b"<!ENTITY" in raw:
        raise ValueError("word/document.xml 带有 DOCTYPE/ENTITY 声明，拒绝解析")
    root = ET.fromstring(raw)
    return "".join(t.text or "" for t in root.iter() if t.tag.endswith("}t"))


# ---------------------------------------------------------------- 工单1

def test_alphabet_scheme_note_matches_gb_order(tmp_path):
    """端到端：实际重排顺序、排序键与审计 note 必须都指向 GB 近似拼音序。"""
    src = make_src(tmp_path)
    ir = parse(tmp_path, src)
    rp = alphabet_rules(tmp_path)
    _stdout, out = render(tmp_path, ir, rp, citation=True)

    items = audit_items(out)
    assert items, "顺序有变化就该有文献表排序条目"
    assert items[0]["rule"] == "citation.reference_list.sort"
    note = items[0]["note"]
    assert "GB 序" in note, f"note 必须写明实际按 GB 序排：{note}"
    assert "字母序" in note, "note 要与规则集的字母序声明对应"
    assert "码位序" not in note, "实现已不是码位序，旧描述必须移除"

    texts = [p.text for p in Document(str(out)).paragraphs]
    pos = {t[:2]: i for i, t in enumerate(texts) if t.startswith(("张三", "周一"))}
    assert pos["张三"] < pos["周一"], "张三（zhang）应排在周一（zhou）前"


# ---------------------------------------------------------------- 工单2

def sdt_with(paras):
    """w:sdt，sdtContent 下依次挂给定文字的段落。"""
    sdt = parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>")
    content = sdt.find(qn("w:sdtContent"))
    for text in paras:
        content.append(parse_xml(
            f'<w:p {nsdecls("w")}><w:r><w:t>{text}</w:t></w:r></w:p>'))
    return sdt


def make_nested_src(tmp_path):
    """外层 sdt（1 直接段落 + 1 嵌套 sdt）+ 文献表的源稿。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    outer = sdt_with(["外层内容控件里的段落"])
    outer.find(qn("w:sdtContent")).append(sdt_with(["嵌套内容控件里的段落"]))
    doc.element.body.append(outer)
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHANG)
    p = tmp_path / "nested.docx"
    doc.save(str(p))
    return p


def test_nested_sdt_is_processed_and_content_kept(tmp_path):
    """外层与嵌套 sdt 都进 IR 并套版；全部可处理时不应虚报未处理块。"""
    src = make_nested_src(tmp_path)
    ir = load_ir(parse(tmp_path, src))

    assert not ir.get("unsupported_content_blocks"), \
        "可递归处理的嵌套内容控件不该计为未处理"
    texts = [b.get("text", "") for b in ir["blocks"]]
    assert any("外层内容控件里的段落" in t for t in texts)
    assert any("嵌套内容控件里的段落" in t for t in texts), \
        "嵌套 sdt 的段落必须递归进 IR"

    r, out = render(tmp_path, src.with_suffix(".ir.json"), RULES / "tjyj.json")
    all_text = docx_xml_text(out)
    assert "外层内容控件里的段落" in all_text
    assert "嵌套内容控件里的段落" in all_text, "套版后嵌套文字必须原样保留"
    assert "未处理内容块:" not in r, "全部处理成功时不应出现未处理块汇总"


def test_inline_sdt_inside_paragraph_is_not_counted(tmp_path):
    """边界钉：段落内部的行内 sdt 文字已随段落处理，不算未处理块。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    p = parse_xml(
        f'<w:p {nsdecls("w")}><w:r><w:t>行内控件前</w:t></w:r>'
        f'<w:sdt {nsdecls("w")}><w:sdtPr/><w:sdtContent>'
        f'<w:r><w:t>行内控件里</w:t></w:r>'
        f'</w:sdtContent></w:sdt>'
        f'<w:r><w:t>行内控件后</w:t></w:r></w:p>')
    doc.element.body.append(p)
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHANG)
    src = tmp_path / "inline.docx"
    doc.save(str(src))

    ir = load_ir(parse(tmp_path, src))
    assert "unsupported_content_blocks" not in ir, \
        "行内 sdt 的文字随段落文本处理了，不该记成未处理块"
    texts = [b.get("text", "") for b in ir["blocks"]]
    assert any("行内控件前" in t and "行内控件里" in t and "行内控件后" in t
               for t in texts), "行内控件里的文字要在段落聚合文本里"
