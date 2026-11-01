"""必备声明：23 本刊都声明了，报告里却从来没出现过。

`statements` 声明各刊要哪些声明——版权转让、利益冲突、数据可得性、
AI 使用声明（6 本必填）。作者不会自己知道「这本刊要 AI 使用声明」，
报告不提，他就漏了。

只做 C 档：列出要求，判断稿件里提没提到，不代写、不代补。
"""

import json
import re

import pytest
from docx import Document

from conftest import RULES, run_tool


def build_doc(path, extra=None):
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("正文段落。")
    for line in (extra or []):
        doc.add_paragraph(line)
    doc.save(str(path))
    return path


def alt_rules(tmp_path, name, statements):
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["statements"] = statements
    path = tmp_path / name
    path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return path


def report(tmp_path, docx_path, rules_path, out_name="triage.html"):
    ir = tmp_path / "m.ir.json"
    run_tool("parse_docx.py", docx_path, "-o", ir)
    out = tmp_path / out_name
    run_tool("compliance_report.py", ir, rules_path, "-o", out)
    raw = out.read_text(encoding="utf-8")
    return raw, re.sub(r"<[^>]+>", " ", raw)


def test_statements_section_is_rendered(tmp_path):
    """规则集声明了几条，页面就列几条。"""
    doc = build_doc(tmp_path / "a.docx")
    raw, text = report(tmp_path, doc, RULES / "tjyj.json")

    assert "该刊要求的声明" in text, "报告里没有声明区块"
    for label in ("版权转让协议", "利益冲突声明", "数据可得性声明"):
        assert label in text, f"缺少声明项：{label}"


def test_required_and_missing_goes_to_todo(tmp_path):
    """必填且稿件里没提到，要进待办提醒。"""
    doc = build_doc(tmp_path / "b.docx")
    _, text = report(tmp_path, doc, RULES / "tjyj.json")
    assert re.search(r"该刊要求版权转让协议", text), "必填声明缺失却没有提醒"


def test_optional_missing_is_not_a_todo(tmp_path):
    """可选声明缺失不该派活给作者。"""
    doc = build_doc(tmp_path / "c.docx")
    _, text = report(tmp_path, doc, RULES / "tjyj.json")
    assert not re.search(r"该刊要求利益冲突声明", text), "可选声明被派成了待办"


def test_detected_statement_is_not_a_todo(tmp_path):
    """稿件里已经写了这份声明，就不该再提醒。"""
    rules = alt_rules(tmp_path, "ai.json",
                      {"ai_declaration": {"required": True, "note": "须知：须声明 AI 使用情况"}})
    doc = build_doc(tmp_path / "d.docx", extra=["AI 使用声明：本文使用 ChatGPT 辅助语言润色。"])
    _, text = report(tmp_path, doc, rules)

    assert "AI 使用声明" in text, "没列出这条声明"
    assert "已提到" in text, "没识别出稿件里已有该声明"
    assert not re.search(r"该刊要求AI 使用声明", text), "已写了却还在待办里"


def test_missing_statement_is_listed_as_not_found(tmp_path):
    rules = alt_rules(tmp_path, "ai2.json",
                      {"ai_declaration": {"required": True, "note": "须声明 AI 使用情况"}})
    doc = build_doc(tmp_path / "e.docx")
    _, text = report(tmp_path, doc, rules)

    assert "未检出" in text, "该声明缺失却没标未检出"
    assert re.search(r"该刊要求AI 使用声明", text), "必填缺失没有提醒"


def test_unknown_statement_key_still_shows(tmp_path):
    """规则集出现工具不认识的声明键，也要显示出来，不能静默丢弃。"""
    rules = alt_rules(tmp_path, "unknown.json",
                      {"ethics_review": {"required": True, "note": "须附伦理审查批件"}})
    doc = build_doc(tmp_path / "f.docx")
    _, text = report(tmp_path, doc, rules)
    assert "ethics_review" in text, "未知声明键被吞掉了"
    assert "须附伦理审查批件" in text, "声明的说明没显示"


def test_statement_detection_is_explained(tmp_path):
    """「未检出」不等于「缺失」，报告要把这个区别讲清楚。"""
    doc = build_doc(tmp_path / "g.docx")
    _, text = report(tmp_path, doc, RULES / "tjyj.json")
    assert "不代表一定缺失" in text, "没有解释未检出的含义"
