"""字数上限：规则集声明了几种，报告就该覆盖几种。

规则集里字数限制有两种写法——单向的 `value`（国际金融研究：350 字）
和区间 `min`/`max`（统计研究：300–500 字）。原先代码只认 `value`，
于是统计研究的摘要限制在报告里等于不存在。关键词条数与标题字数则完全没查。

这些都属于 C 档：只报不改。删摘要、删关键词、改标题都是作者的决定。
"""

import re

import pytest
from docx import Document

from conftest import RULES, run_tool


def build_doc(path, title="数字金融发展与企业融资约束",
              abstract="内容提要：本文考察数字金融对企业融资约束的影响。",
              keywords="关键词：数字金融；融资约束；企业创新"):
    doc = Document()
    doc.add_paragraph(title)
    doc.add_paragraph(abstract)
    doc.add_paragraph(keywords)
    doc.add_paragraph("正文段落。")
    doc.save(str(path))
    return path


def report_text(tmp_path, docx_path, journal):
    """跑一遍体检，返回去掉标签后的纯文本。"""
    ir = tmp_path / "m.ir.json"
    run_tool("parse_docx.py", docx_path, "-o", ir)
    out = tmp_path / "triage.html"
    run_tool("compliance_report.py", ir, RULES / f"{journal}.json", "-o", out)
    return re.sub(r"<[^>]+>", " ", out.read_text(encoding="utf-8"))


def row_value(text, label):
    """取出「字数」表里某一行的内容。"""
    m = re.search(label + r"\s+([^|<]{0,80})", text)
    assert m, f"报告里没有「{label}」这一行"
    return re.sub(r"\s+", " ", m.group(1)).strip()


def test_abstract_range_is_reported(tmp_path):
    """区间型上限（300–500）必须原样呈现，不能当成没写。"""
    doc = build_doc(tmp_path / "a.docx")
    text = report_text(tmp_path, doc, "tjyj")
    row = row_value(text, "摘要")
    assert "300" in row and "500" in row, f"区间没显示出来：{row}"
    assert "低于下限" in row, f"49 字远低于 300，应提示低于下限：{row}"


def test_abstract_single_value_is_reported(tmp_path):
    """单向上限（350）仍按原样呈现。"""
    doc = build_doc(tmp_path / "b.docx")
    row = row_value(report_text(tmp_path, doc, "gjjryj"), "摘要")
    assert "/ 350" in row, f"摘要行没显示上限 350：{row}"
    assert "余量" in row, f"未超限时应显示余量：{row}"


def test_abstract_over_limit_is_flagged(tmp_path):
    """超上限要提示，且不替作者删字。"""
    long_abs = "内容提要：" + "本文考察数字金融对企业融资约束的影响。" * 30
    doc = build_doc(tmp_path / "c.docx", abstract=long_abs)
    text = report_text(tmp_path, doc, "gjjryj")
    row = row_value(text, "摘要")
    assert "超出" in row or "高于建议值" in row, f"超上限没提示：{row}"
    assert "需作者自行精简" in text, "没有把决定权交回作者"


def test_keywords_count_is_checked(tmp_path):
    """关键词条数按规则集上下限判断。"""
    six = "关键词：数字金融；融资约束；企业创新；公司治理；货币政策；信息披露"
    doc = build_doc(tmp_path / "d.docx", keywords=six)
    row = row_value(report_text(tmp_path, doc, "tjyj"), "关键词")
    assert "6" in row, f"没数出 6 个关键词：{row}"
    assert "不在范围内" in row, f"6 个超出 3–5 的上限，应提示：{row}"


def test_keywords_in_range_passes(tmp_path):
    doc = build_doc(tmp_path / "e.docx")
    row = row_value(report_text(tmp_path, doc, "tjyj"), "关键词")
    assert "在范围内" in row, f"3 个在上限内，不该报警：{row}"


def test_title_char_limit_is_checked(tmp_path):
    """标题字数上限（国际金融研究 20 字）。"""
    long_title = "数字金融发展对企业融资约束与创新产出的影响机制研究"
    doc = build_doc(tmp_path / "f.docx", title=long_title)
    row = row_value(report_text(tmp_path, doc, "gjjryj"), "标题字数")
    assert "20" in row, f"没显示上限：{row}"
    assert "超出" in row, f"标题 {len(long_title)} 字超过 20，应提示：{row}"


def test_undeclared_limit_says_so(tmp_path):
    """规则集没声明时要说"未声明"，不能默认通过。"""
    doc = build_doc(tmp_path / "g.docx")
    row = row_value(report_text(tmp_path, doc, "tjyj"), "标题字数")
    assert "未声明" in row, f"没声明却没说明：{row}"


def test_report_does_not_modify_the_manuscript(tmp_path):
    """体检只读：跑完稿子一个字不能变。"""
    doc = build_doc(tmp_path / "h.docx")
    before = doc.read_bytes()
    report_text(tmp_path, doc, "tjyj")
    assert doc.read_bytes() == before, "体检报告改动了原稿"


@pytest.mark.parametrize("label", ["内容摘要：", "内容提要：", "摘　要：", "提要："])
def test_abstract_label_variants_are_recognized(tmp_path, label):
    """摘要的四种常见写法都要认出来。

    解析器原先只认「内容摘要／摘要」，而渲染器的标签正则里早就写了
    「内容提要／提要」。两边不一致的后果：《经济研究》那种用「内容提要」的稿子
    解析不成 abstract，于是字号不套、字数不查、标签也不归一——全都静默失效。
    """
    doc = build_doc(tmp_path / "v.docx", abstract=f"{label}本文考察数字金融的影响。")
    ir = tmp_path / "v.ir.json"
    run_tool("parse_docx.py", doc, "-o", ir)

    import json
    kinds = {b["type"] for b in json.loads(ir.read_text(encoding="utf-8"))["blocks"]}
    assert "abstract" in kinds, f"「{label}」没被识别成摘要，实际类型：{sorted(kinds)}"
