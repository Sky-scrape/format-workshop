"""F42 回归：apply_details.py 的纯函数单测 + 最小集成。

生产链第一棒（补录注入：插正文段落、写首页脚注、重排 IR idx）此前零
测试；DESIGN §10.5「审计比对基线 = 填充完成后的快照」从未被构造过。
纯函数离线单测：with_label / split_author_line / fill_placeholders；
集成面走 main()（子进程，参照 test_citation_convert.py）：作者信息页
带专用样式、首页脚注落位、IR idx 按 body 子元素 dense 整体重算并与
docx 实际段落逐块对齐。
"""

import json
import sys
from pathlib import Path

import pytest
from docx import Document

from conftest import TOOLS, body_texts, run_tool

sys.path.insert(0, str(TOOLS))
import apply_details as ad  # noqa: E402


# ---------------------------------------------------------------- with_label

def test_with_label_prepends_journal_label():
    assert ad.with_label("本文考察数字金融。", "内容摘要：") == "内容摘要：本文考察数字金融。"


def test_with_label_strips_user_typed_duplicate():
    """用户已手写同款标签时去重，不产出「内容摘要：内容摘要：…」。"""
    assert ad.with_label("内容摘要:本文考察数字金融。", "内容摘要：") == \
        "内容摘要：本文考察数字金融。"


def test_with_label_strips_stale_jel_prefix():
    """JEL 分类的旧前缀写法要先剥掉，标签由规则集统一给。"""
    assert ad.with_label("JEL: E62; O17", "JEL分类号：") == "JEL分类号：E62; O17"


def test_with_label_without_label_is_identity():
    assert ad.with_label("无标签文本", "") == "无标签文本"


# ---------------------------------------------------------------- 作者行拆分

def test_split_author_line_takes_both_fields():
    """「作者：张三　　单位：××大学」一行两字段必须拆开，否则成品出现
    「作者姓名：作者：张三……」双标签。"""
    assert ad.split_author_line("作者：张三\u3000\u3000单位：××大学经济学院") == \
        ("张三", "××大学经济学院")
    assert ad.split_author_line("作者:张三  单位:××大学") == ("张三", "××大学")


def test_split_author_line_bare_line_and_name_only():
    assert ad.split_author_line("张三，李四") == ("张三，李四", "")
    assert ad.split_author_line("作者：张三") == ("张三", "")


# ---------------------------------------------------------------- 占位符填充

def _template_doc(tmp_path):
    """前置区模板：占位符用全角括号包着，正是真实模板的形状。"""
    doc = Document()
    doc.add_paragraph("单位：[单位全称，城市，邮编]")
    doc.add_paragraph("作者：[作者姓名]")
    doc.add_paragraph("正文第一段。")
    doc.save(str(tmp_path / "tpl.docx"))
    doc2 = Document(str(tmp_path / "tpl.docx"))
    blocks = [{"idx": i, "type": "front_matter" if i < 2 else "body",
               "text": p.text} for i, p in enumerate(doc2.paragraphs)]
    block_el = {id(b): list(doc2.element.body)[b["idx"]] for b in blocks}
    return doc2, blocks, block_el


def test_fill_placeholders_fills_only_filled_fields(tmp_path):
    """复现 F42：填了单位 → 占位符就地替换；没填姓名 → 占位符原样保留，
    不擅自删字。IR 块的 text 同步更新。"""
    doc, blocks, block_el = _template_doc(tmp_path)

    done = ad.fill_placeholders(doc, blocks, block_el,
                                {"affiliation_zh": "××大学经济学院"})
    assert done == {"affiliation"}
    assert doc.paragraphs[0].text == "单位：××大学经济学院"
    assert doc.paragraphs[1].text == "作者：[作者姓名]", "未填字段不许动"
    assert blocks[0]["text"] == "单位：××大学经济学院", "IR 块文字没有同步"
    assert blocks[1]["text"] == "作者：[作者姓名]"

    done = ad.fill_placeholders(doc, blocks, block_el, {"author_name": "李四"})
    assert done == {"author_name"}
    assert doc.paragraphs[1].text == "作者：李四"


# ---------------------------------------------------------------- 最小集成

RULES = {
    "journal": {"id": "f42j", "name_zh": "集成测试刊"},
    "typography": {"body": {"size_pt": 10.5, "font_cn": "宋体"},
                   "footnote": {"size_pt": 9}},
    "submission": {"author_page": {
        "separate_page": True,
        "fields": ["paper_title", "author_name", "affiliation"]}},
    "personal_fields": [
        {"id": "paper_title", "label_zh": "文章标题"},
        {"id": "author_name", "label_zh": "作者姓名"},
        {"id": "affiliation", "label_zh": "作者单位"},
        {"id": "funding_info", "label_zh": "基金项目"},
    ],
    "blocks": [
        {"type": "title", "lang": "en", "label": "Title："},
        {"type": "abstract", "lang": "en", "label": "Abstract："},
    ],
}

DETAILS = {
    "affiliation_zh": "××大学经济学院",
    "english_title": "Digital Finance and the Real Economy",
    "english_abstract": "This paper examines digital finance.",
    "funding": "国家自然科学基金（72003012）",
}

BLOCKS = [
    {"idx": 0, "type": "front_matter", "text": "数字金融与实体经济"},
    {"idx": 1, "type": "front_matter", "text": "作者：张三\u3000\u3000单位：××大学"},
    {"idx": 2, "type": "abstract", "text": "内容摘要：本文考察数字金融。"},
    {"idx": 3, "type": "keywords", "text": "关键词：数字金融；实体经济"},
    {"idx": 4, "type": "body", "text": "正文第一段，用于核验注入后的块序对位。"},
    {"idx": 5, "type": "body", "text": "正文第二段，用于核验注入后的块序对位。"},
]


@pytest.fixture()
def injected(tmp_path):
    """最小稿件：手写 IR 与同序 docx，跑完 apply_details 的 main()。"""
    src = tmp_path / "manuscript.docx"
    doc = Document()
    for b in BLOCKS:
        doc.add_paragraph(b["text"])
    doc.save(str(src))

    ir = tmp_path / "manuscript.ir.json"
    ir.write_text(json.dumps({"source": src.name, "blocks": BLOCKS},
                             ensure_ascii=False), encoding="utf-8")
    details = tmp_path / "details.json"
    details.write_text(json.dumps(DETAILS, ensure_ascii=False), encoding="utf-8")
    rules = tmp_path / "rules.json"
    rules.write_text(json.dumps(RULES, ensure_ascii=False), encoding="utf-8")

    filled, filled_ir = tmp_path / "filled.docx", tmp_path / "filled.ir.json"
    out = run_tool("apply_details.py", src, ir, details, rules,
                   "--docx-out", filled, "--ir-out", filled_ir)
    return {
        "stdout": out,
        "docx": filled,
        "ir": json.loads(filled_ir.read_text(encoding="utf-8")),
        "source_texts": body_texts(filled),
    }


def test_author_page_paragraphs_carry_dedicated_style(injected):
    """作者信息页每段都带专用样式「作者信息页」：匿名稿据此整页删除，
    不必猜哪些段落属于这一页。"""
    doc = Document(str(injected["docx"]))
    page = [p for p in doc.paragraphs if p.style.name == ad.AUTHOR_PAGE_STYLE]
    assert len(page) == 3, f"作者信息页应有 3 行，实际 {len(page)}"
    assert page[0].text == "文章标题：数字金融与实体经济"
    assert page[1].text == "作者姓名：张三"
    assert page[2].text == "作者单位：××大学经济学院"
    # 分页符只放在本页最后一段：删页时整页一起消失
    assert page[-1]._p.xpath('.//w:br[@w:type="page"]'), "作者信息页末尾没有分页符"
    assert "作者信息页 3 行" in injected["stdout"]


def test_funding_lands_on_first_page_footer(injected):
    """基金项目写在首页脚注（作者页未列出的项才进脚注，不重复两遍）。"""
    doc = Document(str(injected["docx"]))
    sec = doc.sections[0]
    assert sec.different_first_page_header_footer is True
    footer = "\n".join(p.text for p in sec.first_page_footer.paragraphs)
    assert "基金项目：国家自然科学基金（72003012）" in footer
    assert "作者单位" not in footer, "作者页已有的项不该再进脚注"
    assert "首页脚注 1 行" in injected["stdout"]


def test_ir_idx_dense_realigned_with_docx(injected):
    """复现 F42 核心回归：注入后 IR 的 idx 必须「整体重算」，与 docx 实际
    body 子元素逐块 dense 对齐（DESIGN §10.5 审计比对基线的口径）。"""
    texts = injected["source_texts"]
    blocks = injected["ir"]["blocks"]
    assert all(texts[b["idx"]] == b["text"] for b in blocks), \
        "IR idx 与 docx 实际段落错位——重算没生效"

    # 旧映射是插前序号：作者页 3 段插在最前，原题名必须从 0 挪到 3
    title = next(b for b in blocks if b["text"] == "数字金融与实体经济")
    assert title["idx"] == 3, "原题名的 idx 还是插前旧值，重算没发生"

    # 注入的三个要件：中文单位 + 两个英文要件，均带 injected 标记与规则集给的标签
    en = [b for b in blocks if b.get("injected")]
    assert {b["type"] for b in en} == {"affiliations", "english_title", "english_abstract"}
    by_type = {b["type"]: b for b in en}
    assert by_type["affiliations"]["text"] == "××大学经济学院"
    assert by_type["english_title"]["text"] == \
        "Title：Digital Finance and the Real Economy"
    assert by_type["english_abstract"]["text"] == \
        "Abstract：This paper examines digital finance."
    assert by_type["english_abstract"]["idx"] == by_type["english_title"]["idx"] + 1
    assert "english_title" in injected["stdout"] and "english_abstract" in injected["stdout"]

    # 原正文的文字一字未动（本工具只插入、不改写）
    for b in BLOCKS:
        if b["type"] == "body":
            assert texts[next(x["idx"] for x in blocks if x["text"] == b["text"])] == \
                b["text"]
