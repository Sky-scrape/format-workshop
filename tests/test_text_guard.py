"""白名单守卫的回归：抓得住真改字，也不冤枉白名单内的归一。

用例全部用 python-docx 现造源稿/成品对，不依赖任何期刊规则集；最后一条
parity 用例跑真实流水线，防止搬家到 tools/text_guard.py 后与 conftest 的
判定漂移——同一不变式两份实现，迟早一边先松。
"""

import json

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from conftest import (body_texts, run_tool,
                      whitelist_violations as conftest_whitelist_violations)
from text_guard import document_texts, normalize, whitelist_violations


def write_docx(path, paragraphs):
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(str(path))
    return path


def write_sdt_docx(path, text):
    """body 里放一个内容控件，sdtContent 下挂一段正文。"""
    doc = Document()
    sdt = parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>")
    sdt.find(qn("w:sdtContent")).append(parse_xml(
        f'<w:p {nsdecls("w")}><w:r><w:t>{text}</w:t></w:r></w:p>'))
    doc.element.body.insert(0, sdt)
    doc.save(str(path))
    return path


# ---- 未审计的改字 ----

def test_unaudited_text_change_is_caught(tmp_path):
    """正文改了字、审计里没有任何条目——必须报出来。"""
    src = write_docx(tmp_path / "src.docx", ["原始句子。"])
    out = write_docx(tmp_path / "out.docx", ["改过的句子。"])

    got = whitelist_violations(document_texts(src), document_texts(out),
                               {0: "body"}, [])

    assert got == [(0, "body", "原始句子。", "改过的句子。")]


def test_layout_item_does_not_launder_text_change(tmp_path):
    """版式条目（before/after 是格式值）不能给改字背书。"""
    src = write_docx(tmp_path / "src.docx", ["原始句子。"])
    out = write_docx(tmp_path / "out.docx", ["改过的句子。"])
    items = [{"id": "L1", "idx": 0, "action": "正文套版", "rule": "typography.body"}]

    got = whitelist_violations(document_texts(src), document_texts(out),
                               {0: "body"}, items)

    assert [v[0] for v in got] == [0]


# ---- 白名单内的归一 ----

def test_caption_number_normalization_is_allowed(tmp_path):
    """题注只改编号与间隔（表1　→　表 2），文字本身逐字未动。"""
    src = write_docx(tmp_path / "src.docx", ["表1　数据来源"])
    out = write_docx(tmp_path / "out.docx", ["表 2 数据来源"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "table_caption"}, []) == []


def test_caption_word_change_is_caught(tmp_path):
    """题注编号之外的文字改了，仍算改字。"""
    src = write_docx(tmp_path / "src.docx", ["表1 数据来源"])
    out = write_docx(tmp_path / "out.docx", ["表1 数据说明"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "table_caption"}, [])


def test_reference_number_normalization_is_allowed(tmp_path):
    """文献条目只换编号写法（[1] → 〔1〕），条目正文逐字未动。"""
    src = write_docx(tmp_path / "src.docx", ["[1] 张三. 数字金融研究"])
    out = write_docx(tmp_path / "out.docx", ["〔1〕张三. 数字金融研究"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "reference_item"}, []) == []


def test_reference_entry_change_is_caught(tmp_path):
    """编号之后的条目文字改了，不在 N-07 白名单里。"""
    src = write_docx(tmp_path / "src.docx", ["[1] 张三. 数字金融研究"])
    out = write_docx(tmp_path / "out.docx", ["〔1〕李四. 数字金融研究"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "reference_item"}, [])


def test_label_normalization_is_allowed(tmp_path):
    """摘要标签整段换成目标刊标签，冒号随之变化，正文未动。"""
    src = write_docx(tmp_path / "src.docx", ["内容摘要：本文考察数字金融。"])
    out = write_docx(tmp_path / "out.docx", ["【摘要】本文考察数字金融。"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "abstract"}, [],
                                labels={"abstract": "【摘要】"}) == []


def test_keyword_separator_normalization_is_allowed(tmp_path):
    """关键词分隔符统一为「；」，词条逐个对得上。"""
    src = write_docx(tmp_path / "src.docx", ["关键词：数字金融，企业融资；创新"])
    out = write_docx(tmp_path / "out.docx", ["关键词：数字金融；企业融资；创新"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "keywords"}, [],
                                labels={"keywords": "关键词："},
                                kw_sep="；") == []


def test_heading_number_prefix_rewrite_is_allowed(tmp_path):
    """标题重新编号（一、→ 1.）只换前缀，标题文字未动。"""
    src = write_docx(tmp_path / "src.docx", ["一、研究背景"])
    out = write_docx(tmp_path / "out.docx", ["1. 研究背景"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "heading1"}, []) == []


def test_normalize_does_not_widen_to_unimplemented_rules():
    """N-02～N-06 与段尾空白仍是「暂不处理」：这些差异必须留在白名单外。"""
    assert normalize("研究  发现", "body") != normalize("研究 发现", "body")
    assert normalize("２０２０", "body") != normalize("2020", "body")
    assert normalize("数字金融 A 股", "body") != normalize("数字金融A股", "body")


# ---- 表注换位 ----

def test_table_note_reorder_is_allowed(tmp_path):
    """「注」与「资料来源」只换位置：整组文字守恒，不算改字。"""
    src = write_docx(tmp_path / "src.docx",
                     ["表1 回归结果", "注：数据来源于统计年鉴。", "资料来源：作者整理。"])
    out = write_docx(tmp_path / "out.docx",
                     ["表1 回归结果", "资料来源：作者整理。", "注：数据来源于统计年鉴。"])
    items = [{"id": "N1", "idx": 0, "action": "表注／图注层序调整",
              "rule": "tables.note_layers", "before": "source / general",
              "after": "general / source"}]

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "table_caption", 1: "table_note",
                                 2: "table_note"}, items) == []


def test_single_table_note_edit_is_caught(tmp_path):
    """只有一条表注、文字又变了，不是换位。"""
    src = write_docx(tmp_path / "src.docx", ["注：数据来源于甲。"])
    out = write_docx(tmp_path / "out.docx", ["注：数据来源于乙。"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "table_note"}, [])


# ---- 引用改写豁免 ----

def test_citation_rewrite_with_matching_audit_is_allowed(tmp_path):
    """--citation 的正文引用改写：动作与 rule 对得上才豁免。"""
    src = write_docx(tmp_path / "src.docx", ["张三（2020）发现数字金融降低融资约束。"])
    out = write_docx(tmp_path / "out.docx", ["Zhang（2020）发现数字金融降低融资约束。"])
    items = [{"id": "C1", "idx": 0, "action": "正文引用改作者-年份",
              "rule": "citation.in_text"}]

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "body"}, items) == []


def test_citation_rewrite_with_wrong_rule_is_caught(tmp_path):
    """动作名对、rule 指错位置——豁免不成立，照旧算改字。"""
    src = write_docx(tmp_path / "src.docx", ["张三（2020）发现数字金融降低融资约束。"])
    out = write_docx(tmp_path / "out.docx", ["Zhang（2020）发现数字金融降低融资约束。"])
    items = [{"id": "C1", "idx": 0, "action": "正文引用改作者-年份",
              "rule": "citation.in_text.typo"}]

    assert [v[0] for v in whitelist_violations(
        document_texts(src), document_texts(out), {0: "body"}, items)] == [0]


# ---- sdt 槽位 ----

def test_sdt_slot_change_without_audit_is_caught(tmp_path):
    """内容控件里的正文改了字、没有槽位审计——旧 body_texts 整槽跳过，现在必须抓到。"""
    src = write_sdt_docx(tmp_path / "src.docx", "内容控件里的原句。")
    out = write_sdt_docx(tmp_path / "out.docx", "内容控件里的改句。")
    source_texts, output_texts = document_texts(src), document_texts(out)

    assert source_texts[0] == "内容控件里的原句。"
    assert whitelist_violations(source_texts, output_texts, {0: "body"}, []) == [
        (0, "body", "内容控件里的原句。", "内容控件里的改句。")]


def test_sdt_slot_change_with_slot_audit_is_allowed(tmp_path):
    """槽位改动有 "slot." 文字动作条目覆盖，由条目里的快照供核对。"""
    src = write_sdt_docx(tmp_path / "src.docx", "内容控件里的原句。")
    out = write_sdt_docx(tmp_path / "out.docx", "内容控件里的改句。")
    items = [{"id": "S1", "idx": "slot.0", "action": "题注归一",
              "rule": "tables.caption_gap"}]

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "body"}, items) == []


def test_sdt_slot_layout_item_does_not_cover_text_change(tmp_path):
    """槽位审计也必须是文字动作；版式条目不能给槽内改字背书。"""
    src = write_sdt_docx(tmp_path / "src.docx", "内容控件里的原句。")
    out = write_sdt_docx(tmp_path / "out.docx", "内容控件里的改句。")
    items = [{"id": "S2", "idx": "slot.0", "action": "正文套版",
              "rule": "typography.body"}]

    assert [v[0] for v in whitelist_violations(
        document_texts(src), document_texts(out), {0: "body"}, items)] == [0]


def test_sdt_slot_indent_action_counts_as_text_audit(tmp_path):
    """N-01／N-08 的段首缩进删除也改写文字，槽位审计要认它。"""
    src = write_sdt_docx(tmp_path / "src.docx", "　　内容控件里的原句。")
    out = write_sdt_docx(tmp_path / "out.docx", "内容控件里的原句。")
    items = [{"id": "S3", "idx": "slot.0", "action": "段首缩进空格删除",
              "rule": "typography.body.first_line_indent_chars"}]

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "body"}, items) == []


def test_sdt_slot_audit_accepts_renderer_index_form(tmp_path):
    """渲染器把槽内条目写成 "0.0"（槽.序），不是 "slot.0"；守卫必须认。"""
    src = write_sdt_docx(tmp_path / "src.docx", "　　内容控件里的原句。")
    out = write_sdt_docx(tmp_path / "out.docx", "内容控件里的原句。")
    items = [{"id": "S4", "idx": "0.0", "action": "段首缩进空格删除",
              "rule": "typography.body.first_line_indent_chars"}]

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "body"}, items) == []


def _write_two_sdt_docx(path, texts):
    doc = Document()
    for pos, text in enumerate(texts):
        sdt = parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>")
        sdt.find(qn("w:sdtContent")).append(parse_xml(
            f'<w:p {nsdecls("w")}><w:r><w:t>{text}</w:t></w:r></w:p>'))
        doc.element.body.insert(pos, sdt)
    doc.save(str(path))
    return path


def test_slot_audit_does_not_cover_another_slot(tmp_path):
    """一条 "0.0" 条目只能覆盖 0 号槽，不能给 1 号槽的改字背书。"""
    src = _write_two_sdt_docx(tmp_path / "src.docx", ["原句甲。", "原句乙。"])
    out = _write_two_sdt_docx(tmp_path / "out.docx", ["改句甲。", "改句乙。"])
    items = [{"id": "S5", "idx": "0.0", "action": "题注归一",
              "rule": "tables.caption_gap"}]

    offenders = whitelist_violations(document_texts(src), document_texts(out),
                                     {0: "body", 1: "body"}, items)
    assert [v[0] for v in offenders] == [1]


def test_missing_block_is_caught_as_a_text_loss(tmp_path):
    """成品末尾少一段：只拿输出下标遍历会完全看不见，必须报成丢字。"""
    src = write_docx(tmp_path / "src.docx", ["第一段。", "第二段。"])
    out = write_docx(tmp_path / "out.docx", ["第一段。"])

    assert whitelist_violations(document_texts(src), document_texts(out),
                                {0: "body", 1: "body"}, []) == [
        (1, "body", "第二段。", "")]


def test_heading_digit_prefix_requires_a_separator():
    """「5G技术研究」不是编号标题；「1.2.3研究」「1 引言」仍是。"""
    assert normalize("5G技术研究", "heading1") == "5G技术研究"
    assert normalize("3D打印", "heading1") == "3D打印"
    assert normalize("1.2.3研究", "heading2") == "研究"
    assert normalize("1 引言", "heading1") == "引言"


def test_references_heading_change_needs_its_audit_item(tmp_path):
    """文献表标题改字必须有同 idx 的「文献表标题归一」条目，不能无条件放行。"""
    src = write_docx(tmp_path / "src.docx", ["参考文献"])
    out = write_docx(tmp_path / "out.docx", ["参考文献："])
    texts = (document_texts(src), document_texts(out))

    assert [v[0] for v in whitelist_violations(
        texts[0], texts[1], {0: "references_heading"}, [])] == [0]

    items = [{"id": "R1", "idx": 0, "action": "文献表标题归一",
              "rule": "blocks:references/label"}]
    assert whitelist_violations(texts[0], texts[1],
                                {0: "references_heading"}, items) == []


# ---- 取字口径 ----

def test_document_texts_separates_table_cells(tmp_path):
    """表格按单元格取字并用 \\x00 分隔，相邻单元格的文字不会粘成一段。"""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    for r, row in enumerate((("甲", "乙"), ("丙", "丁"))):
        for c, text in enumerate(row):
            table.cell(r, c).text = text
    path = tmp_path / "table.docx"
    doc.save(str(path))

    assert document_texts(path)[0] == "甲\x00乙\x00丙\x00丁"


def test_document_texts_matches_body_texts_on_plain_blocks(tmp_path):
    """普通段落/表格的取字与 conftest.body_texts 完全同源（除 sdt 新增那一支）。"""
    src = write_docx(tmp_path / "plain.docx", ["第一段", "第二段"])

    assert document_texts(src) == body_texts(src)


def test_sdt_slot_index_matches_ir(tmp_path):
    """槽位下标与 IR 的 idx 同源：解析器给的槽号能直接取到整槽文字。"""
    src = write_sdt_docx(tmp_path / "sdt.docx", "内容控件里的正文。")
    ir_path = tmp_path / "sdt.ir.json"
    run_tool("parse_docx.py", src, "-o", ir_path)
    blocks = json.loads(ir_path.read_text(encoding="utf-8"))["blocks"]
    block = next(b for b in blocks if "sdt_pos" in b)

    texts = document_texts(src)

    assert block["idx"] in texts
    assert texts[block["idx"]] == block["text"]


# ---- 搬家不漂移 ----

def test_whitelist_matches_conftest_on_real_pipeline(baseline):
    """真实跨刊渲染结果上，两边判定逐条相同（排序、截断、豁免都对齐）。"""
    got = whitelist_violations(baseline.source_texts, baseline.output_texts,
                               baseline.block_type, baseline.items,
                               labels=baseline.labels, kw_sep=baseline.keyword_sep)

    assert got == conftest_whitelist_violations(baseline)
