"""参考文献条目：只改编号写法，不动一个字。

条目正文里改标点、改作者分隔符、补类型标识，都属于「改写内容」——
本产品不做。能做而且该做的是编号前缀的写法：各家期刊要求不同
（[1] / 〔1〕 / ［1］ / 1．），换刊重排时这一处必须跟着变。
"""

import json

import pytest
from docx import Document

from conftest import RULES, render_into

ITEMS = [
    "[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。",
    "[2] 李四、王五，2021：《融资约束的度量》，《金融研究》第7期。",
    "[3] Smith J., 2019: Digital Finance, Journal of Finance, 74(3): 1-30.",
]


def build_reference_doc(path, items=None):
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文段落，用于确认参考文献之前的内容不受影响。" * 3)
    doc.add_paragraph("参考文献")
    for item in (items or ITEMS):
        doc.add_paragraph(item)
    doc.save(str(path))
    return path


def ref_texts(pipeline):
    """取出成品里的参考文献条目文字。"""
    return [t for i, t in sorted(pipeline.output_texts.items())
            if pipeline.block_type.get(i) == "reference_item"]


@pytest.mark.parametrize("target,expected", [
    ("tjyj", "[1]"),      # 统计研究：[1]
    ("jmsc", "〔1〕"),    # 管理科学学报：〔1〕
])
def test_number_style_follows_target_journal(tmp_path, target, expected):
    """换一本刊，编号写法要跟着变。"""
    src = build_reference_doc(tmp_path / f"ref-{target}.docx")
    result = render_into(tmp_path / f"run-{target}", src, target=target)

    texts = ref_texts(result)
    assert texts, "成品里没有参考文献条目"
    assert texts[0].startswith(expected), f"编号没换成 {expected}：{texts[0][:12]}"


def test_number_only_prefix_changes(tmp_path):
    """除了编号前缀，条目正文必须一字不差。"""
    src = build_reference_doc(tmp_path / "ref.docx")
    result = render_into(tmp_path / "run", src, target="jmsc")

    texts = ref_texts(result)
    for original, after in zip(ITEMS, texts):
        assert after[after.index("〕") + 1:].lstrip() == original[original.index("]") + 1:].lstrip()


def test_numbering_is_audited(tmp_path):
    """编号改写要进报告，并带规则引用。"""
    src = build_reference_doc(tmp_path / "ref2.docx")
    result = render_into(tmp_path / "run2", src, target="jmsc")

    items = [it for it in result.items if it["action"] == "文献条目编号归一"]
    assert len(items) == len(ITEMS), f"应记 {len(ITEMS)} 条，实际 {len(items)}"
    for it in items:
        assert it["rule"] == "citation.reference_list.number_format"
        assert it["before"] and it["after"] and it["before"] != it["after"]


def test_unnumbered_entry_is_left_alone(tmp_path):
    """本来就没编号的条目，不替它编一个号——那是添加内容。"""
    items = ["张三，2020：《数字金融与企业融资》，《经济研究》第3期。"]
    src = build_reference_doc(tmp_path / "ref3.docx", items=items)
    result = render_into(tmp_path / "run3", src, target="tjyj")

    texts = ref_texts(result)
    assert texts[0] == items[0], f"无编号条目被改动了：{texts[0][:20]}"
    actions = {it["action"] for it in result.items}
    assert "文献条目编号归一" not in actions


def test_year_leading_entry_is_not_mistaken_for_number(tmp_path):
    """以年份开头的条目不能被当成「第 201 号文献」。"""
    items = ["2020：《数字金融与企业融资》，《经济研究》第3期。"]
    src = build_reference_doc(tmp_path / "ref4.docx", items=items)
    result = render_into(tmp_path / "run4", src, target="jmsc")

    texts = ref_texts(result)
    assert texts[0].startswith("2020"), f"年份被误当成编号改写：{texts[0][:16]}"


def test_numbering_can_be_rejected(tmp_path):
    """驳回后要保留原编号写法，并记录为已驳回。"""
    src = build_reference_doc(tmp_path / "ref5.docx")
    probe = render_into(tmp_path / "probe", src, target="jmsc")
    target_item = next(it for it in probe.items if it["action"] == "文献条目编号归一")

    rejected = render_into(tmp_path / "rejected", src, target="jmsc",
                           rejects=[target_item["id"]])
    texts = ref_texts(rejected)
    assert texts[0].startswith("[1]"), f"驳回后编号仍被改了：{texts[0][:12]}"

    marked = rejected.item_by_id(target_item["id"])
    assert marked["rejected"] is True
