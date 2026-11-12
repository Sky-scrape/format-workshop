"""换刊时的人员信息承接提醒（DESIGN §13.3）。

作者页上的字段不会重复写进首页脚注：从要求作者页的刊换到不要求的刊，
已填内容不会自动跟着走。提示必须出现，但工具不替用户搬内容。
"""

import json
import sys
from pathlib import Path

from conftest import RULES, TOOLS

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(TOOLS))

import app  # noqa: E402


def _rules(jid):
    return json.loads((RULES / f"{jid}.json").read_text(encoding="utf-8"))


def test_author_page_to_plain_journal_warns():
    old, new = _rules("gjjryj"), _rules("tjyj")
    notes = app.author_info_handoff(old, new, {"author_name": "张三"})
    assert notes, "从要求作者页的刊换到不要求时没有提醒"
    joined = "\n".join(notes)
    assert "作者页" in joined and "首页脚注" in joined
    assert "国际金融研究" in joined and "统计研究" in joined


def test_plain_journal_to_author_page_warns():
    notes = app.author_info_handoff(_rules("tjyj"), _rules("gjjryj"))
    assert any("单独成页" in n for n in notes)


def test_lost_footnote_fields_are_named():
    old = {"journal": {"name_zh": "甲刊"}, "personal_fields": [
        {"id": "funding_info", "label_zh": "基金项目", "placement": "footnote"}]}
    new = {"journal": {"name_zh": "乙刊"}, "personal_fields": []}
    notes = app.author_info_handoff(old, new)
    assert any("基金项目" in n and "首页脚注" in n for n in notes)


def test_same_shape_journals_have_no_noise():
    a = {"journal": {"name_zh": "甲刊"}, "personal_fields": [
        {"id": "funding_info", "label_zh": "基金项目", "placement": "footnote"}]}
    b = {"journal": {"name_zh": "乙刊"}, "personal_fields": [
        {"id": "funding_info", "label_zh": "基金项目", "placement": "footnote"}]}
    assert app.author_info_handoff(a, b) == [], "规则集结构相同却弹提醒"
