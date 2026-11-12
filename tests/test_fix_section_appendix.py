"""投稿体检的定位辅助：分节字数条与可附录化候选（DESIGN §5.2）。

两项都只读：给出「哪一节最占篇幅」「如果把这张表移入附录，正文大约少多少字」，
不搬表、不删字、不生成任何缩减版本（§5.3）。门槛按 B 口径 500 字。
"""

import json
import sys

from conftest import RULES, TOOLS

sys.path.insert(0, str(TOOLS))

import compliance_report as report  # noqa: E402


def _rules():
    return json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))


def _ir(tables=()):
    blocks = [
        {"type": "front_matter", "text": "数字金融与融资约束", "idx": 0},
        {"type": "abstract", "text": "内容提要：本文考察数字金融的影响。", "idx": 1},
        {"type": "keywords", "text": "关键词：数字金融；融资约束", "idx": 2},
        {"type": "heading1", "text": "一、引言", "idx": 3},
        {"type": "body", "text": "本文研究数字金融。" * 10, "idx": 4},
    ]
    for i, cells in enumerate(tables):
        blocks.append({"type": "table_caption", "text": f"表{i + 1} 大型统计表", "idx": 5 + i})
        blocks.append({"type": "table", "idx": 6 + i, "n_rows": 12, "n_cols": 6,
                       "header": ["变量"] * 6, "cells": cells})
    blocks += [
        {"type": "heading1", "text": "二、实证结果", "idx": 20},
        {"type": "body", "text": "回归结果稳健。" * 5, "idx": 21},
    ]
    return {"source": "x.docx", "blocks": blocks}


def test_sections_carry_share_and_three_metrics():
    res = report.check(_ir(), _rules())
    assert [s["title"] for s in res["sections"]] == ["引言", "实证结果"]
    total = sum(s["words"] for s in res["sections"])
    assert total > 0
    for s in res["sections"]:
        assert s["words"] > 0 and s["A"] > 0 and s["C"] >= s["A"]


def test_large_table_is_an_appendix_candidate():
    big = ["样本观测值统计" * 60 for _ in range(12)]
    res = report.check(_ir(tables=[big]), _rules())
    assert res["appendix_candidates"], "500 字以上的表没有被列为候选"
    c = res["appendix_candidates"][0]
    assert c["label"].startswith("表1"), c
    assert c["B"] >= 500 and c["rows"] == 12 and c["cols"] == 6
    html_out = report.render(_ir(tables=[big]), _rules(), res)
    assert "可附录化候选" in html_out
    assert "本工具不会搬动任何表" in html_out
    assert "表1 大型统计表" in html_out
    # 门槛说明必须在场，读者才知道候选是怎么筛出来的
    assert "500" in html_out


def test_small_table_is_not_a_candidate():
    small = ["小表"] * 6
    res = report.check(_ir(tables=[small]), _rules())
    assert res["appendix_candidates"] == []
    html_out = report.render(_ir(tables=[small]), _rules(), res)
    assert "没有达到" in html_out, "没有候选时也应该说明门槛，而不是静默消失"
