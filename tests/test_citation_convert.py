"""引用体系重排：默认不碰，勾选后才改写。

这个能力会动正文引用和文献条目文字，是产品里唯一越线的开关，所以两条
底线必须钉住：
1. 不传 --citation 时，条目与正文一字不动（红线照旧）
2. 传了 --citation 时，正文每个编号都要能对上文献表，对不上就整批不转
"""

import json
import subprocess
import sys

import pytest
from docx import Document

from conftest import TOOLS, minimal_doc, mutated_rules

sys.path.insert(0, str(TOOLS))
import citation_ref as cr  # noqa: E402

TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
    "en_journal": "{authors}，{year}，“{title}”，{journal}{volume_part}{page_part}.",
    "en_book": "{authors}，{year}，“{title}”，{publisher}.",
    "en_other": "{authors}，{year}，“{title}”，{source}.",
}


def rules_with_templates(tmp_path, name="cite.json"):
    """基于统计研究改一份带条目模板的规则集（仓库自带的刊没这字段）。"""
    return mutated_rules(tmp_path, "tjyj", name,
                         **{"citation.system": "author-year",
                            "citation.in_text": {"en_three_plus": True},
                            "citation.reference_list.templates": TEMPLATES,
                            "citation.reference_list.must_require_bidirectional_match":
                                True})


def build_doc(path, items=None, body=None):
    default_body = [
        "正文第一段，引用了文献一[1]和文献二[2]。",
        "正文第二段，用区间写法引用了[1-2]。",
    ]
    default_items = [
        "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.",
    ]
    minimal_doc(path, title="数字金融与融资约束",
                abstract="内容提要：本文考察数字金融对企业融资约束的影响。",
                extra_paragraphs=[*(body or default_body), "参考文献",
                                  *(items or default_items)])
    return path


def run_render(tmp_path, docx, rules, citation):
    cmd = [sys.executable, str(TOOLS / "render_from_ir.py"),
           str(docx.with_suffix(".ir.json")), str(rules),
           "-o", str(docx.with_name("out.docx"))]
    if citation:
        cmd.append("--citation")
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(docx),
                    "-o", str(docx.with_suffix(".ir.json"))],
                   capture_output=True, text=True, check=True)
    r = subprocess.run(cmd, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr[-800:]
    return Document(str(docx.with_name("out.docx")))


# ---------- 纯逻辑 ----------

def test_parse_cn_journal():
    e = cr.parse_entry("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    assert e["lang"] == "cn" and e["slot"] == "journal"
    assert e["authors"] == ["张三"]
    assert e["year"] == "2020" and e["issue"] == "3" and e["pages"] == "1-10"


def test_parse_en_journal_keeps_comma_in_title():
    """刊名里带逗号时不能被拆断。"""
    e = cr.parse_entry(
        "[7] Parasuraman R, Sheridan T B, Wickens C D. A model[J]. "
        "IEEE Transactions on Systems, Man, and Cybernetics—Part A, 2000, 30(3): 286-297.")
    assert e["journal"] == "IEEE Transactions on Systems, Man, and Cybernetics—Part A"
    assert e["volume"] == "30" and e["issue"] == "3"


def test_url_is_not_taken_as_source():
    """「https:」的冒号不能被当成「出版地: 出版者」的分隔符。"""
    e = cr.parse_entry(
        "[15] 中国日报网. 某条新闻[EB/OL]. (2025-07-29)[2026-10-29]. https://cn.chinadaily.com.cn/a/x.html")
    assert e["source"] == ""
    assert e["year"] == "2025"


@pytest.mark.parametrize("raw,expected", [
    ("Kang J K", "Kang，J. K."),
    ("Maffettone P M", "Maffettone，P. M."),
    ("Weiss，D.", "Weiss，D."),
])
def test_en_author_initials(raw, expected):
    assert cr.en_author(raw) == expected


def test_in_text_three_authors_cn():
    e = {"authors": ["王化成", "曹丰", "叶康涛"], "year": "2015", "lang": "cn"}
    assert cr.in_text_label(e) == "（王化成等，2015）"


def test_in_text_two_authors_en():
    e = {"authors": ["Kang J K", "Kim J M"], "year": "2008", "lang": "en"}
    assert cr.in_text_label(e) == "（Kang and Kim，2008）"


@pytest.mark.parametrize("spec,expected", [
    ("1", [1]), ("1,3", [1, 3]), ("5-6", [5, 6]), ("2，4", [2, 4]),
])
def test_expand_numbers(spec, expected):
    assert cr.expand_numbers(spec) == expected


def test_check_reports_missing_entry():
    index = cr.build_index(["[1] 张三. 甲[J]. 某刊, 2020, 1(1): 1-2."])
    _, problems = cr.check(["正文引用了[2]。"], index, {"citation": {}})
    assert any("[2]" in p for p in problems)


def test_check_reports_uncited_entry():
    index = cr.build_index([
        "[1] 张三. 甲[J]. 某刊, 2020, 1(1): 1-2.",
        "[2] 李四. 乙[J]. 某刊, 2021, 1(1): 1-2.",
    ])
    rules = {"citation": {"reference_list": {"must_require_bidirectional_match": True}}}
    _, problems = cr.check(["正文只引用了[1]。"], index, rules)
    assert any("[2]" in p and "没有" in p for p in problems)


# ---------- 端到端 ----------

def test_entries_untouched_by_default(tmp_path):
    """不勾选时文献条目一字不动——这是产品的既有承诺。"""
    src = build_doc(tmp_path / "a.docx")
    out = run_render(tmp_path, src, rules_with_templates(tmp_path), citation=False)
    texts = [p.text for p in out.paragraphs]
    assert "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10." in texts
    assert any("正文第一段，引用了文献一[1]和文献二[2]。" == t for t in texts)


def test_citation_switch_rewrites_both_sides(tmp_path):
    src = build_doc(tmp_path / "b.docx")
    out = run_render(tmp_path, src, rules_with_templates(tmp_path), citation=True)
    texts = [p.text for p in out.paragraphs]
    joined = "\n".join(texts)

    assert "（张三，2020）" in joined, joined[:400]
    assert "（Smith，2021）" in joined
    assert "张三：《数字金融与企业融资》，《经济研究》，2020年第3期。" in joined
    assert "Smith，J.，2021，" in joined
    assert "[1]" not in joined and "[2]" not in joined


def test_citation_skipped_when_mapping_broken(tmp_path):
    """正文引用了文献表里没有的编号时，整批不转，并在审计里写明。"""
    src = build_doc(tmp_path / "c.docx", body=["正文引用了不存在的[9]。"])
    out = run_render(tmp_path, src, rules_with_templates(tmp_path), citation=True)
    joined = "\n".join(p.text for p in out.paragraphs)

    assert "[9]" in joined, "对不上映射时仍改了正文"
    assert "张三. 数字金融与企业融资[J]" in joined, "条目被改动了"

    audit = json.loads((tmp_path / "out.audit.json").read_text(encoding="utf-8"))
    skip = [i for i in audit["items"] if "引用体系转换（已跳过）" == i["action"]]
    assert skip and "[9]" in skip[0]["note"]


# ---------- 斜体保留（italic_fields，2026-11-13） ----------

def rules_with_italics(tmp_path, name="cite_italic.json"):
    """英文刊声明刊名斜体：模板占位符之外，斜体事实也必须来自规则集。"""
    return mutated_rules(
        tmp_path, "tjyj", name,
        **{"citation.system": "author-year",
           "citation.in_text": {"en_three_plus": True},
           "citation.reference_list.templates": TEMPLATES,
           "citation.reference_list.italic_fields": {"en_journal": ["journal"]},
           "citation.reference_list.must_require_bidirectional_match": True})


def test_format_entry_spans_declared_field_only():
    rules = {"citation": {"reference_list": {
        "templates": {"en_journal": "{authors}，{year}，“{title}”，{journal}"
                                    "{volume_part}{page_part}."},
        "italic_fields": {"en_journal": ["journal"]}}}}
    e = cr.parse_entry(
        "[2] Smith J. Digital finance and constraints[J]. "
        "Journal of Finance, 2021, 76(2): 20-40.")
    text, spans = cr.format_entry_with_spans(e, rules)
    assert spans, "声明了 italic_fields 却没有返回斜体区间"
    s, t = spans[0]
    assert text[s:t] == "Journal of Finance", (text, spans)
    assert cr.format_entry(e, rules) == text, "带区间渲染与纯文本渲染必须逐字一致"

    # 未声明时区间为空，行为与旧版一致
    plain = {"citation": {"reference_list": {
        "templates": {"en_journal": "{authors}，{year}，{journal}."}}}}
    text2, spans2 = cr.format_entry_with_spans(e, plain)
    assert spans2 == [] and "Journal of Finance" in text2


def test_italic_survives_end_to_end_rewrite(tmp_path):
    src = build_doc(tmp_path / "d.docx")
    out = run_render(tmp_path, src, rules_with_italics(tmp_path), citation=True)
    para = next(p for p in out.paragraphs if "Journal of Finance" in p.text)
    italic_runs = [r for r in para.runs if r.italic and "Journal of Finance" in r.text]
    assert italic_runs, [(r.text, r.italic) for r in para.runs]
    assert "Journal of Finance" in para.text
