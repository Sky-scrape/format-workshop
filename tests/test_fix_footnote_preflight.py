"""版式体检：脚注/尾注使用方式的只读检测（DESIGN §12.1 遗留缺口）。

工具仍然只调脚注字体字号、不搬位置、不重编号；但「稿件用了尾注」「编号
样式/范围与刊要求不一致」「手敲的①②③不是真脚注」必须出现在体检报告里，
归入「需你处理」，不能假装没看见。
"""

import json
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from conftest import RULES, TOOLS

sys.path.insert(0, str(TOOLS))

import preflight  # noqa: E402
from rules_ref import resolvable  # noqa: E402


def _rules():
    return json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))


def _doc_with_refs(path, footnote=0, endnote=0, fake=None, num_fmt=None, restart=None):
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    p = doc.add_paragraph("正文引用了注释")
    for i in range(footnote):
        r = OxmlElement("w:r")
        ref = OxmlElement("w:footnoteReference")
        ref.set(qn("w:id"), str(i + 1))
        r.append(ref)
        p._p.append(r)
    for i in range(endnote):
        r = OxmlElement("w:r")
        ref = OxmlElement("w:endnoteReference")
        ref.set(qn("w:id"), str(i + 1))
        r.append(ref)
        p._p.append(r)
    for text in fake or []:
        doc.add_paragraph(text)
    if num_fmt or restart:
        pr = OxmlElement("w:footnotePr")
        if num_fmt:
            nf = OxmlElement("w:numFmt")
            nf.set(qn("w:val"), num_fmt)
            pr.append(nf)
        if restart:
            nr = OxmlElement("w:numRestart")
            nr.set(qn("w:val"), restart)
            pr.append(nr)
        doc.settings.element.append(pr)
    doc.save(str(path))
    return path


def _items(result, prop):
    return [it for it in result["items"] if it["property"] == prop]


def test_endnotes_reported_as_position_mismatch(tmp_path):
    src = _doc_with_refs(tmp_path / "end.docx", endnote=2)
    result = preflight.analyze(src, _rules())
    items = _items(result, "位置")
    assert items and items[0]["fix"] == preflight.FIX_MANUAL
    assert "尾注" in items[0]["actual"]
    assert resolvable(_rules(), items[0]["rule"])


def test_footnote_number_format_and_restart_reported(tmp_path):
    src = _doc_with_refs(tmp_path / "fn.docx", footnote=2,
                         num_fmt="decimal", restart="continuous")
    result = preflight.analyze(src, _rules())
    fmt = _items(result, "编号样式")
    scope = _items(result, "编号范围")
    assert fmt and fmt[0]["fix"] == preflight.FIX_MANUAL, "①②③ 与阿拉伯数字的差异没报"
    assert "①" in fmt[0]["expected"] and "阿拉伯" in fmt[0]["actual"]
    assert scope and scope[0]["fix"] == preflight.FIX_MANUAL
    assert "全文连续" in scope[0]["actual"] and "每页" in scope[0]["expected"]
    rules = _rules()
    assert resolvable(rules, fmt[0]["rule"]) and resolvable(rules, scope[0]["rule"])


def test_matching_footnote_setup_is_not_reported(tmp_path):
    src = _doc_with_refs(tmp_path / "ok.docx", footnote=1,
                         num_fmt="decimalEnclosedCircle", restart="eachPage")
    result = preflight.analyze(src, _rules())
    assert not _items(result, "编号样式"), "编号本来就合规却报了偏差"
    assert not _items(result, "编号范围"), "重编号范围本来就合规却报了偏差"


def test_typed_circled_numbers_are_manual_fake_footnotes(tmp_path):
    src = _doc_with_refs(tmp_path / "fake.docx",
                         fake=["① 数据来自 Wind 数据库。", "注：② 同前注。"])
    result = preflight.analyze(src, _rules())
    items = _items(result, "手工脚注")
    assert items and items[0]["fix"] == preflight.FIX_MANUAL
    assert "手敲" in items[0]["note"]
