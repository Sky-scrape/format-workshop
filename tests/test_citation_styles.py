"""英文刊引用插件与样式映射（DESIGN §9.2 / journal-schema §18.3，2026-11-15）。

三件事：
1. style_id 指向的 APA/Harvard/Chicago 插件真的把条目改写成目标样式，
   斜体区间（刊名/卷、书名）落在最终文本的正确位置；
2. 规则集写了 templates 时模板优先，插件不覆盖期刊自己的体例；
3. 样式映射表（style_map）按语义块类型查 Word 样式名，代码不含刊名特判。
"""

import json
import sys
from pathlib import Path

from docx import Document

from conftest import RULES, minimal_doc, render_into

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))

import citation_ref as cr  # noqa: E402
import cite_styles  # noqa: E402

ARTICLE = ("[1] Smith J, Jones A. Growth targets and local policy[J]. "
           "Journal of Comparative Economics, 2020, 48(3): 512-530.")
BOOK = "[2] Lee B. Economic growth[M]. Thousand Oaks: Sage, 2014."


def _entry(raw):
    return cr.parse_entry(raw)


# ---------------------------------------------------------------- 插件

def test_apa_formats_authors_year_journal_and_italics():
    text, spans = cite_styles.format_entry(_entry(ARTICLE), {"citation": {"style_id": "APA-7"}})
    assert text == ("Smith, J., & Jones, A. (2020). Growth targets and local policy. "
                    "Journal of Comparative Economics, 48(3), 512-530.")
    italic = [text[s:e] for s, e in spans]
    assert italic == ["Journal of Comparative Economics", "48"]


def test_harvard_variants_differ_in_year_position():
    rules = {"citation": {"style_id": "HARVARD-ELSEVIER"}}
    text, _ = cite_styles.format_entry(_entry(ARTICLE), rules)
    assert text.startswith("Smith, J., Jones, A., 2020. Growth targets")

    rules = {"citation": {"style_id": "HARVARD-WILEY"}}
    text, _ = cite_styles.format_entry(_entry(ARTICLE), rules)
    assert text.startswith("Smith, J. and Jones, A. 2020, 'Growth targets")


def test_chicago_quotes_title():
    text, spans = cite_styles.format_entry(_entry(ARTICLE), {"citation": {"style_id": "CHICAGO"}})
    assert '"Growth targets and local policy."' in text
    assert [text[s:e] for s, e in spans] == ["Journal of Comparative Economics", "48"]


def test_book_title_is_italic_and_no_double_period():
    for style, needle in (("APA-7", "Lee, B. (2014). "),
                          ("HARVARD-ELSEVIER", "Lee, B., 2014. "),
                          ("HARVARD-WILEY", "Lee, B., 2014. ")):
        text, spans = cite_styles.format_entry(_entry(BOOK), {"citation": {"style_id": style}})
        assert text.startswith(needle), (style, text)
        assert ".." not in text, (style, text)
        assert [text[s:e] for s, e in spans] == ["Economic growth"]


def test_unknown_style_returns_none():
    assert cite_styles.format_entry(_entry(ARTICLE), {"citation": {"style_id": "GB/T 7714"}}) == (None, [])
    assert cite_styles.resolve("") is None


def test_ruleset_templates_take_precedence_over_plugin():
    rules = {"citation": {"style_id": "APA-7",
                          "reference_list": {"templates": {
                              "en_journal": "{title}[J]. {journal}, {year}."}}}}
    text, _ = cr.format_entry_with_spans(_entry(ARTICLE), rules)
    assert text == "Growth targets and local policy[J]. Journal of Comparative Economics, 2020."
    assert cr.entry_from_plugin(_entry(ARTICLE), rules) is False
    assert cr.entry_from_plugin(_entry(ARTICLE), {"citation": {"style_id": "APA-7"}}) is True


def test_in_text_labels_follow_ruleset_examples():
    entry = _entry(ARTICLE)
    rules = {"citation": {
        "system": "author-year", "style_id": "APA-7",
        "in_text": {"bracket": "()", "one_author": "(Smith, 2020)",
                    "two_authors": "(Smith & Jones, 2020)",
                    "three_plus": "(Smith et al., 2020)",
                    "multiple_separator": "; "}}}
    assert cr.in_text_label(entry, rules) == "(Smith & Jones, 2020)"

    harvard = {"citation": {
        "system": "author-year", "style_id": "HARVARD-ELSEVIER",
        "in_text": {"bracket": "()", "one_author": "(Smith, 2020)",
                    "two_authors": "(Smith and Jones, 2020)",
                    "three_plus": "(Smith et al., 2020)"}}}
    assert cr.in_text_label(entry, harvard) == "(Smith and Jones, 2020)"


# ---------------------------------------------------------------- 端到端

def _english_doc(path):
    doc = Document()
    doc.add_paragraph("Growth Targets and Local Policy")
    doc.add_paragraph("Abstract: This paper studies growth targets.")
    doc.add_paragraph("Keywords: growth; policy")
    doc.add_paragraph("1. Introduction")
    doc.add_paragraph("Previous studies document [1]. Another study [2] finds more.")
    doc.add_paragraph("References")
    doc.add_paragraph(ARTICLE)
    doc.add_paragraph(BOOK)
    doc.save(str(path))
    return path


def test_end_to_end_apa_conversion_is_audited(tmp_path):
    src = _english_doc(tmp_path / "en.docx")
    rules = json.loads((RULES / "cer.json").read_text(encoding="utf-8"))
    rp = tmp_path / "cer.json"
    rp.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")

    p = render_into(tmp_path / "out", src, target="cer", rules_path=rp, citation=True)

    body = next(t for t in p.output_texts.values() if "Previous studies" in t)
    assert "(Smith & Jones, 2020)" in body and "(Lee, 2014)" in body
    entries = [t for i, t in p.output_texts.items()
               if p.block_type.get(i) == "reference_item"]
    assert any(t.startswith("Smith, J., & Jones, A. (2020).") for t in entries)
    assert any(t.startswith("Lee, B. (2014).") for t in entries)

    plugin_items = [it for it in p.items if it["action"] == "条目按样式插件重排"]
    assert len(plugin_items) == 2
    assert all(it["rule"] == "citation.style_id" for it in plugin_items)


# ---------------------------------------------------------------- style_map

def test_style_map_applies_declared_word_style(tmp_path):
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["style_map"] = {"body": "Body Text"}
    rp = tmp_path / "tjyj-style.json"
    rp.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    src = tmp_path / "sm.docx"
    minimal_doc(src, keywords="关键词：数字金融；融资约束",
                extra_paragraphs=["一、引言", "本文考察数字金融对企业融资约束的影响，样本为A股上市公司。"])

    p = render_into(tmp_path / "sm-out", src, target="tjyj", rules_path=rp)

    from docx import Document as _D
    doc = _D(str(p.docx))
    names = {par.text: par.style.name for par in doc.paragraphs}
    assert names["本文考察数字金融对企业融资约束的影响，样本为A股上市公司。"] == "Body Text"
    items = [it for it in p.items if it["action"] == "套用 Word 样式"]
    assert items and items[0]["rule"] == "style_map.body"
