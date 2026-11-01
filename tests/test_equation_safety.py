"""公式：既不能被删，也不能被裁。

Word 里两句古话在这里都成立：
- 「删光 run 再塞新 run」会把 OMML 公式连同对象一起删掉，而且不报错；
- 「固定值」行距会把分式、积分、上下标裁掉，同样不报错。

两类事故的共同点：文字看着都在，只是公式没了、或者断了一半。
"""

import json
import subprocess
import sys

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls

from conftest import NS_W, ROOT, render_into, run_tool

M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def add_math(paragraph, content="y=α+βx"):
    """往段落尾部追加一个 OMML 公式对象。"""
    paragraph._p.append(parse_xml(
        f'<m:oMath {nsdecls("m", "w")}><m:r><m:t>{content}</m:t></m:r></m:oMath>'))
    return paragraph


def count_math(docx_path):
    doc = Document(str(docx_path))
    return len(list(doc.element.body.iter(M_NS + "oMath")))


def build_equation_manuscript(path):
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")

    # 正文中的行内公式
    doc.add_paragraph("本文的基本设定如下，其中系数由下式给出：" * 3)
    inline = add_math(doc.add_paragraph("其中 β 为待估系数，" * 4))
    inline.paragraph_format.line_spacing = 18

    # 独立公式段落：带分式，最能体现「被裁」
    eq = doc.add_paragraph()
    add_math(eq, "β=(X'X)^{-1}X'y")
    eq.paragraph_format.line_spacing = 24

    doc.add_paragraph("上式为最小二乘估计量。")
    doc.save(str(path))
    return path


def test_replace_text_keeps_math_objects():
    """replace_text 只该改文字，不该把公式对象一起删掉。"""
    sys.path.insert(0, str(ROOT / "tools"))
    from render_from_ir import replace_text

    doc = Document()
    p = doc.add_paragraph("一、模型设定与")
    add_math(p, "y")
    assert count_math_paragraph(p) == 1

    replace_text(p, "二、模型设定与")

    assert count_math_paragraph(p) == 1, "改写文字时把公式对象删掉了"
    assert p.text.startswith("二、"), f"文字没替换成功：{p.text!r}"


def count_math_paragraph(p):
    return len(list(p._p.iter(M_NS + "oMath")))


def test_heading_with_math_is_still_a_heading(tmp_path):
    """带公式的标题必须仍是标题，否则整篇不再编号。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融的影响。")
    add_math(doc.add_paragraph("一、模型设定与"), "y")
    doc.add_paragraph("本节的模型设定如下，包含若干待估参数与识别假设。" * 3)

    src = tmp_path / "heading_math.docx"
    doc.save(str(src))

    run_tool("parse_docx.py", src, "-o", tmp_path / "h.ir.json")
    ir = json.loads((tmp_path / "h.ir.json").read_text(encoding="utf-8"))
    kinds = {b["type"] for b in ir["blocks"]}
    assert "heading1" in kinds, f"带公式的标题没有被识别成标题：{sorted(kinds)}"
    assert "equation" not in kinds, "标题被误判成公式段落"


def test_math_survives_rendering(tmp_path):
    """整条流水线跑下来，公式对象的数量不能少。"""
    src = build_equation_manuscript(tmp_path / "eq.docx")
    before = count_math(src)
    assert before >= 2, "构造的稿件公式太少，测不出问题"

    result = render_into(tmp_path / "run", src, target="tjyj")
    assert count_math(result.docx) == before, "渲染后公式数量变了"


def test_equation_paragraph_uses_at_least_spacing(tmp_path):
    """公式段落的行距必须是「最小值」，不能是「固定值」。"""
    src = build_equation_manuscript(tmp_path / "eq2.docx")
    result = render_into(tmp_path / "run2", src, target="tjyj")

    doc = Document(str(result.docx))
    exact = []
    for i, child in enumerate(doc.element.body):
        if child.tag != NS_W + "p":
            continue
        if not list(child.iter(M_NS + "oMath")):
            continue
        spacing = child.find(f".//{NS_W}pPr/{NS_W}spacing")
        if spacing is not None and spacing.get(NS_W + "lineRule") == "exact":
            exact.append(i)
    assert not exact, f"这些公式段仍被设成固定行距，会被裁掉：{exact}"


def test_equation_restyling_is_audited(tmp_path):
    """公式段的行距处理要进报告。"""
    src = build_equation_manuscript(tmp_path / "eq3.docx")
    result = render_into(tmp_path / "run3", src, target="tjyj")
    actions = {it["action"] for it in result.items}
    assert "公式段行距" in actions, f"没有记录公式段处理，实际只有：{sorted(actions)}"


def test_pure_math_paragraph_is_not_dropped_by_parser(tmp_path):
    """纯公式段落没有文字，但不能被当成空行丢掉。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融的影响。")
    eq = doc.add_paragraph()
    add_math(eq, "σ²=Σ(x-μ)²/n")
    doc.add_paragraph("上式为企业规模的定义，其中 n 为样本量。" * 3)

    src = tmp_path / "pure_math.docx"
    doc.save(str(src))

    run_tool("parse_docx.py", src, "-o", tmp_path / "p.ir.json")
    ir = json.loads((tmp_path / "p.ir.json").read_text(encoding="utf-8"))
    kinds = [b["type"] for b in ir["blocks"]]
    assert "equation" in kinds, f"纯公式段落被丢了，实际块类型：{kinds}"


def test_pseudo_table_formula_is_reported_as_untouched(tmp_path):
    """用 1 行表格排版的公式：不改结构，但报告里要说清「没动」。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融的影响。")
    doc.add_paragraph("估计量为：" * 6)

    table = doc.add_table(rows=1, cols=2)
    add_math(table.cell(0, 0).paragraphs[0], "β=(X'X)^{-1}X'y")
    table.cell(0, 1).paragraphs[0].text = "（1）"

    doc.add_paragraph("其中各项含义如上。" * 6)
    src = tmp_path / "pseudo.docx"
    doc.save(str(src))

    result = render_into(tmp_path / "run4", src, target="tjyj")
    actions = {it["action"] for it in result.items}
    assert "公式（伪表格）原样保留" in actions, f"公式伪表格没有说明，实际：{sorted(actions)}"

    # 表格结构必须原样：1 行 2 列，公式还在
    kept = [t for t in Document(str(result.docx)).element.body.iter(NS_W + "tbl")]
    assert len(kept) == 1, "公式表格被改动或删除了"
    assert count_math(result.docx) >= 1, "公式对象在渲染后丢失"
