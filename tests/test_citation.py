"""参考文献条目：只改编号写法，不动一个字。

条目正文里改标点、改作者分隔符、补类型标识，都属于「改写内容」——
本产品不做。能做而且该做的是编号前缀的写法：各家期刊要求不同
（[1] / 〔1〕 / ［1］ / 1．），换刊重排时这一处必须跟着变。

夹具只能用仓库自带的期刊；要测别的写法，就临时改一份规则集（rules_path）。
"""

import sys

import pytest

from conftest import ROOT, minimal_doc, mutated_rules, render_into

sys.path.insert(0, str(ROOT / "tools"))
from render_from_ir import format_ref_number, ref_number_prefix  # noqa: E402

ITEMS = [
    "1．张三，2020：《数字金融与企业融资》，《经济研究》第3期。",
    "2．李四、王五，2021：《融资约束的度量》，《金融研究》第7期。",
    "3．Smith J., 2019: Digital Finance, Journal of Finance, 74(3): 1-30.",
]


def build_reference_doc(path, items=None):
    minimal_doc(path, title="数字金融与企业融资约束",
                abstract="内容提要：本文考察数字金融对企业融资约束的影响。",
                extra_paragraphs=["正文段落，用于确认参考文献之前的内容不受影响。" * 3,
                                  "参考文献", *(items or ITEMS)])
    return path


def with_numbering(tmp_path, name, **overrides):
    """基于统计研究改一份规则集，用来测仓库里没有的编号写法。"""
    return mutated_rules(tmp_path, "tjyj", name,
                         **{f"citation.reference_list.{k}": v
                            for k, v in overrides.items()})


def ref_texts(pipeline):
    return [t for i, t in sorted(pipeline.output_texts.items())
            if pipeline.block_type.get(i) == "reference_item"]


# ---------- 纯逻辑：模板渲染 ----------

@pytest.mark.parametrize("fmt,space,expected", [
    ("[1]", None, "[7]"),
    ("[1]", 1, "[7] "),
    ("〔1〕", None, "〔7〕"),
    ("［1］", None, "［7］"),
    ("1．", None, "7．"),
    ("1.", None, "7."),
])
def test_format_ref_number(fmt, space, expected):
    assert format_ref_number(fmt, 7, space) == expected


def test_format_ref_number_returns_none_without_template():
    assert format_ref_number(None, 7) is None


@pytest.mark.parametrize("text,expected_num", [
    ("[1] 张三，2020", 1),
    ("〔12〕李四", 12),
    ("1．张三", 1),
    ("123 张三", 123),
    ("2020：《数字金融》", None),      # 年份开头，不是编号
    ("张三，2020", None),
    ("[1]无空格", 1),
])
def test_ref_number_prefix_recognition(text, expected_num):
    _, num = ref_number_prefix(text)
    assert num == expected_num


# ---------- 端到端：真的排一遍 ----------

def test_number_style_follows_target_journal(tmp_path):
    """换一本刊，编号写法要跟着变。"""
    src = build_reference_doc(tmp_path / "ref.docx")
    result = render_into(tmp_path / "run", src, target="tjyj")

    texts = ref_texts(result)
    assert texts, "成品里没有参考文献条目"
    assert texts[0].startswith("[1]"), f"编号没换成 [1]：{texts[0][:12]}"


def test_space_after_number_is_respected(tmp_path):
    """国际金融研究要求编号后留一个空格，统计研究不留。"""
    src = build_reference_doc(tmp_path / "ref2.docx")

    spaced = ref_texts(render_into(tmp_path / "gjjryj", src, target="gjjryj"))
    tight = ref_texts(render_into(tmp_path / "tjyj", src, target="tjyj"))

    assert spaced[0].startswith("[1] "), f"没有补空格：{spaced[0][:12]}"
    assert tight[0].startswith("[1]张"), f"不该有空格：{tight[0][:12]}"


def test_number_only_prefix_changes(tmp_path):
    """除了编号前缀，条目正文必须一字不差。"""
    src = build_reference_doc(tmp_path / "ref3.docx")
    texts = ref_texts(render_into(tmp_path / "run3", src, target="tjyj"))

    for original, after in zip(ITEMS, texts):
        assert after.split("]", 1)[1] == original.split("．", 1)[1]


def test_exotic_number_style_from_ruleset(tmp_path):
    """规则集写 〔1〕 时也要能排出来（仓库自带的刊里没有这种写法）。"""
    src = build_reference_doc(tmp_path / "ref4.docx")
    alt = with_numbering(tmp_path, "cjk.json", number_format="〔1〕")
    texts = ref_texts(render_into(tmp_path / "run4", src, target="tjyj", rules_path=alt))
    assert texts[0].startswith("〔1〕"), f"编号没换成 〔1〕：{texts[0][:12]}"


def test_numbering_is_audited(tmp_path):
    """编号改写要进报告，并带规则引用。"""
    src = build_reference_doc(tmp_path / "ref5.docx")
    result = render_into(tmp_path / "run5", src, target="tjyj")

    items = [it for it in result.items if it["action"] == "文献条目编号归一"]
    assert len(items) == len(ITEMS), f"应记 {len(ITEMS)} 条，实际 {len(items)}"
    for it in items:
        assert it["rule"] == "citation.reference_list.number_format"
        assert it["before"] and it["after"] and it["before"] != it["after"]


def test_unnumbered_entry_is_left_alone(tmp_path):
    """本来就没编号的条目，不替它编一个号——那是添加内容。"""
    items = ["张三，2020：《数字金融与企业融资》，《经济研究》第3期。"]
    src = build_reference_doc(tmp_path / "ref6.docx", items=items)
    result = render_into(tmp_path / "run6", src, target="tjyj")

    assert ref_texts(result)[0] == items[0], "无编号条目被改动了"
    assert "文献条目编号归一" not in {it["action"] for it in result.items}


def test_numbering_can_be_rejected(tmp_path):
    """驳回后要保留原编号写法，并记录为已驳回。"""
    src = build_reference_doc(tmp_path / "ref7.docx")
    probe = render_into(tmp_path / "probe", src, target="tjyj")
    target_item = next(it for it in probe.items if it["action"] == "文献条目编号归一")

    rejected = render_into(tmp_path / "rejected", src, target="tjyj",
                           rejects=[target_item["id"]])
    assert ref_texts(rejected)[0].startswith("1．"), "驳回后编号仍被改了"
    assert rejected.item_by_id(target_item["id"])["rejected"] is True
