"""审读工单两处修复的回归钉：排序 note 如实 + 嵌套 sdt 计数。

- 工单1  cn-first-by-alphabet-then-en-by-alphabet 的中文键是名字裸比较＝
         Unicode 码位序（sort_key，F22 有意钉住），SORT_NOTE 却写「中文按
         字母序在前」——措辞冒用了拼音字母序（「周一」U+5468 排「张三」
         U+5F20 前，与 zhang<zhou 恰好相反）。现在 note 照实写「码位序
         （非拼音序）」；排序行为不变（test_fix_render_core.py 的
         test_cn_alphabet_scheme_uses_name_codepoint_order 钉着）。
- 工单2  parse_docx 的 sdt 下钻只收 sdtContent 的直接 w:p/w:tbl：外层 sdt
         下钻了、内嵌套的 w:sdt 没有下钻，其内段落既不进 IR 也不计数——
         「未处理内容块 N 个」漏报。现在嵌套的块级 sdt 逐个记数进 IR
         （内容仍原样保留，不套版）；段落/表格内部的行内 sdt 文字已随
         para_text 处理，不数。

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
    """文献表已按拼音字母序（张三[1] 在 周一[2] 前）的源稿。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHANG)
    doc.add_paragraph(REF_ZHOU)
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

def test_alphabet_scheme_note_describes_codepoint_order(tmp_path):
    """工单1 端到端：输入已按拼音序，码位序重排后 note 必须照实说码位序。"""
    src = make_src(tmp_path)
    ir = parse(tmp_path, src)
    rp = alphabet_rules(tmp_path)
    _stdout, out = render(tmp_path, ir, rp, citation=True)

    items = audit_items(out)
    assert items, "顺序有变化就该有文献表排序条目"
    assert items[0]["rule"] == "citation.reference_list.sort"
    note = items[0]["note"]
    assert "码位序" in note, f"note 必须写明实际按码位序排：{note}"
    assert "非拼音序" in note, f"必须点明码位序不是拼音序：{note}"
    assert "中文按字母序" not in note, f"不得冒用拼音字母序：{note}"

    # 排序行为不变（F22 有意钉住码位序）：周一(U+5468) 排到 张三(U+5F20) 前
    texts = [p.text for p in Document(str(out)).paragraphs]
    pos = {t[:2]: i for i, t in enumerate(texts) if t.startswith(("张三", "周一"))}
    assert pos["周一"] < pos["张三"], \
        "码位序行为被改动了——先改 test_cn_alphabet_scheme_uses_name_codepoint_order"


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


def test_nested_sdt_is_counted_and_content_kept(tmp_path):
    """工单2：外层下钻、嵌套 sdt 计数进 IR；嵌套内容在成品里原样保留。"""
    src = make_nested_src(tmp_path)
    ir = load_ir(parse(tmp_path, src))

    assert ir.get("unsupported_content_blocks") == 1, \
        "嵌套内容控件要逐个记数进 IR，不能静默"
    texts = [b.get("text", "") for b in ir["blocks"]]
    assert any("外层内容控件里的段落" in t for t in texts), \
        "外层 sdt 的直接段落照旧下钻进 IR"
    assert not any("嵌套内容控件" in t for t in texts), \
        "下钻口径不变：嵌套 sdt 的段落仍不进 IR（只把没处理报出来）"

    r, out = render(tmp_path, src.with_suffix(".ir.json"), RULES / "tjyj.json")
    assert "嵌套内容控件里的段落" in docx_xml_text(out), \
        "计数只是可见化，内容控件对象必须原样保留在成品里"
    assert "未处理内容块: 1 个" in r, "渲染汇总要报出嵌套内容控件"


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
