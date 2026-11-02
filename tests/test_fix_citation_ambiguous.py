"""F04 回归钉：数量语境的编号样文本不得改写，真引用不得误杀。

上一轮独立验证留下的两条未闭症状：
- 残留误改：「调研对象为[90, 100]家企业。」「[90, 100]为样本量的上下界。」
  在成员恰被文献表收录时被静默改写为作者-年份——语义级错误；
- 假阴性：「另有[5, 6]支持这一区间估计。」里的「区间」属于前半句的
  「区间估计」，6 字窗口子串匹配把真引用 [5, 6] 整批误杀（双向体检随之
  判「文献未被引用」→ 转换功能整体关闭）。

修法口径（citation_ref.classify_spans）：区间词从窗口子串收紧为「紧邻」；
右邻紧贴量词/接续词的命中归入「可疑样文本」——不改写、不计入「已被引用」，
由 apply_citation 逐处写进审计（编号样文本未转换）。
"""

import sys
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import citation_ref as cr


def spans_text(ms):
    return [m.group(0) for m in ms]


def test_quantity_context_is_not_treated_as_citation():
    cites, ambig = cr.classify_spans("调研对象为[90, 100]家企业。")
    assert spans_text(cites) == []
    assert spans_text(ambig) == ["[90, 100]"]


def test_quantity_connector_word_is_not_treated_as_citation():
    cites, ambig = cr.classify_spans("[90, 100]为样本量的上下界。")
    assert spans_text(cites) == []
    assert spans_text(ambig) == ["[90, 100]"]


def test_adjacent_interval_word_is_not_treated_as_citation():
    for text in ("样本在[80, 95]之间。", "区间为[90, 100]。", "取值在[90, 100]之间。"):
        cites, ambig = cr.classify_spans(text)
        assert spans_text(cites) == [], text
        assert len(ambig) == 1, text


def test_distant_interval_word_does_not_kill_real_citations():
    # 「区间」属于前半句的「区间估计」，不归 [5, 6] 管
    cites, ambig = cr.classify_spans("另有[5, 6]支持这一区间估计。")
    assert spans_text(cites) == ["[5, 6]"]
    assert ambig == []


def test_mixed_sentence_keeps_citation_and_flags_range():
    cites, ambig = cr.classify_spans("文献[12, 15]表明，取值在[90, 100]之间。")
    assert spans_text(cites) == ["[12, 15]"]
    assert spans_text(ambig) == ["[90, 100]"]


def test_plain_citations_are_unaffected():
    for text in ("结果表明[1]。", "方法见[12, 15]。", "在[1, 2]中，作者提出",
                 "see [12] for details"):
        cites, ambig = cr.classify_spans(text)
        assert ambig == [], text
        assert len(cites) == 1, text


def _index():
    entries = [
        "[90] 张三．数字金融研究[J]．经济研究，2020，55(3)：1-15．",
        "[100] 李四．货币政策传导[J]．金融研究，2021，42(5)：20-34．",
    ]
    return cr.build_index(entries)


def test_ambiguous_span_is_not_rewritten_even_when_entries_exist():
    """成员恰好都被收录也不能改——这正是 F04 的语义事故场景。"""
    from render_from_ir import rewrite_citations

    index = _index()
    doc = Document()
    p = doc.add_paragraph("调研对象为[90, 100]家企业。")
    old = p.text
    assert rewrite_citations(p, index, {}) == 0
    assert p.text == old


def test_real_citation_still_converts():
    from render_from_ir import rewrite_citations

    index = cr.build_index([
        "[5] 王五．实证方法[J]．统计研究，2019，36(2)：8-20．",
        "[6] 赵六．数据科学[J]．统计研究，2020，37(1)：55-70．",
    ])
    doc = Document()
    p = doc.add_paragraph("另有[5, 6]支持这一结论。")
    assert rewrite_citations(p, index, {}) > 0
    assert "王五" in p.text and "[5, 6]" not in p.text


def test_ambiguous_span_does_not_count_as_cited():
    """可疑样文本不计入「已被引用」：双向体检据此整批拒绝，而非错改。"""
    index = _index()
    _cited, problems = cr.check(["调研对象为[90, 100]家企业。"], index,
                                {"citation": {"reference_list":
                                              {"must_require_bidirectional_match": True}}})
    assert any("没有被引用" in x for x in problems)
