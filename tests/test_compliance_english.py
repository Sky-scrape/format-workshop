"""英文稿的口径：另附要件、编辑部加注项、英文声明与英文单位。

三个报告口径只对中文稿成立，英文稿会集体误判：

1. Highlights 这类要件按刊要求是单独上传的文件，稿件正文里本来就没有。
   原先一律判「缺」，报告把作者推去正文里补一个不存在的段落。
2. 收稿日期由编辑部盖章，作者补不了。判成缺件后这本刊永远停在「还差 1 项」。
3. 声明与单位只按中文词匹配：英文稿里写好的 conflict of interest、英文单位行
   全被判「未检出」，声明被派成待办、单位判缺、作者行也连带判错。

用例全部从入库的 cjyj 合成规则集（mutated_rules），不依赖私有刊文件。
"""

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from conftest import mutated_rules  # noqa: E402

import compliance_report as crp  # noqa: E402


def _rules(tmp_path, name, **keys):
    """从 cjyj 复制一份规则集、按点路径改键（只改点名的键）。"""
    path = mutated_rules(tmp_path, "cjyj", name, **keys)
    return json.loads(path.read_text(encoding="utf-8"))


def _ir(*blocks):
    return {"source": "英文稿.docx", "blocks": list(blocks)}


# ---------------------------------------------------------------- 另附要件

def test_missing_highlights_is_external_not_missing(tmp_path):
    """Highlights 未检出：标「另附」，不进缺失、不派待办。"""
    rules = _rules(tmp_path, "hl.json",
                   blocks=[{"type": "highlights", "lang": None, "required": True}])
    ir = _ir({"type": "body", "text": "正文段落。"})
    res = crp.check(ir, rules)

    assert res["rows"] == [("Highlights", "另附", True)]
    assert ("highlights", None) in res["external"]

    page = crp.render(ir, rules, res)
    assert '<td class="muted">另附</td>' in page, "「另附」没用 muted 与真缺失区分"
    assert "Highlights · 缺" not in page, "另附要件被列进了缺失角标"
    assert "补齐Highlights" not in page, "另附要件被派成了正文待办"
    assert "必备要件齐全" in page, "另附要件不该计入「还差几项」"
    assert "另附文件提交" in page, "报告没说明这是另附文件"


def test_highlights_in_ir_counts_as_present(tmp_path):
    """稿件里真有 highlights 块时照常判「有」，不再提示另附。"""
    rules = _rules(tmp_path, "hl2.json",
                   blocks=[{"type": "highlights", "lang": None, "required": True}])
    ir = _ir({"type": "highlights", "text": "研究亮点。"})
    res = crp.check(ir, rules)

    assert res["rows"] == [("Highlights", "有", True)]
    assert res["external"] == set()
    assert "另附文件提交" not in crp.render(ir, rules, res)


# ---------------------------------------------------------------- 编辑部加注项

def test_received_date_is_not_delegated_to_author(tmp_path):
    """收稿日期由编辑部盖，不能算成作者的缺件。"""
    rules = _rules(tmp_path, "rd.json",
                   blocks=[{"type": "received_date", "required": True}])
    ir = _ir({"type": "body", "text": "正文段落。"})
    res = crp.check(ir, rules)

    assert res["rows"] == [("收稿日期", "缺", True)]
    page = crp.render(ir, rules, res)
    assert "收稿日期 由编辑部提供" in page, "编辑部给定项没说明由谁处理"
    assert "补齐收稿日期" not in page, "编辑部给定项被派给了作者"
    assert "缺失必备项：收稿日期" not in page
    assert "必备要件齐全" in page, "编辑部给定项不该计入「还差几项」"


def test_corresponding_author_has_a_chinese_label(tmp_path):
    """jae/jde 的 corresponding_author 不能把内部类型名甩给用户，也不能在报告里判缺。"""
    rules = _rules(tmp_path, "corr.json",
                   blocks=[{"type": "corresponding_author", "required": True}])
    ir = _ir({"type": "body", "text": "正文段落。"})
    res = crp.check(ir, rules)

    assert ("通讯作者", "补录", True) in res["rows"], res["rows"]
    page = crp.render(ir, rules, res)
    assert "corresponding_author" not in page.split("</head>", 1)[-1], \
        "内部类型名漏进了报告正文"
    assert "在网页补录页填写" in page, "补录项没说明由谁收集"
    assert "缺失必备项：通讯作者" not in page, "补录项被当成正文缺件了"


# ---------------------------------------------------------------- 英文声明

def test_english_statements_are_detected(tmp_path):
    """英文声明语句要能命中，否则每条都被派成「未检出」待办。"""
    rules = _rules(
        tmp_path, "en_statements.json",
        blocks=[{"type": "body", "required": True}],
        statements={
            "copyright_transfer": {"required": True},
            "conflict_of_interest": {"required": True},
            "data_availability": {"required": True},
            "ai_declaration": {"required": True},
            "originality": {"required": True},
            "confidentiality": {"required": True},
        })
    text = (
        "Conflict of interest: the authors declare no competing interests. "
        "Data availability: the data are available from the corresponding author. "
        "AI use: generative AI was not used in this study. "
        "Copyright transfer: the publisher holds the copyright. "
        "Originality: this work has not been published elsewhere. "
        "Confidentiality: the manuscript contains no confidential information."
    )
    res = crp.check(_ir({"type": "body", "text": text}), rules)

    expected = {"copyright_transfer", "conflict_of_interest", "data_availability",
                "ai_declaration", "originality", "confidentiality"}
    found = {s["key"] for s in res["statements"] if s["found"]}
    assert found == expected, f"英文声明没认全：缺 {expected - found}"

    page = crp.render(_ir({"type": "body", "text": text}), rules, res)
    assert page.count("已提到") == 6, "英文声明应在表里标「已提到」"
    for s in res["statements"]:
        assert f"该刊要求{s['label']}" not in page, f"{s['label']}已写明却被派成待办"


# ---------------------------------------------------------------- 英文单位

def test_english_affiliation_line_is_recognized(tmp_path):
    """英文单位行命中机构特征：作者行与单位行分别判对，不互相顶替。"""
    rules = _rules(tmp_path, "en_org.json",
                   blocks=[{"type": "authors", "required": True},
                           {"type": "affiliations", "required": True}])
    ir = _ir(
        {"type": "front_matter", "text": "Digital Finance and Firm Constraints"},
        {"type": "front_matter", "text": "John Smith"},
        {"type": "front_matter",
         "text": "Department of Finance, Renmin University of China"},
    )
    rows = crp.check(ir, rules)["rows"]
    assert ("作者", "有", True) in rows, "英文作者行没被判成作者"
    assert ("单位", "有", True) in rows, "英文单位行没命中机构特征"


@pytest.mark.parametrize("line", [
    "University of International Business and Economics",
    "Institute of Finance and Banking",
    "College of Economics",
    "School of Management",
    "Chinese Academy of Social Sciences",
    "People's Bank of China",
    "Research Center for Digital Finance",
    "Centre for Economic Research",
    "State Key Laboratory of Finance",
    "Department of Statistics",
    "Beijing Hospital",
    "Huawei Corporation",
    "Sample Company Limited",
    "Research Group on Finance",
])
def test_english_org_keywords_match(line):
    """规则集列的英文机构词逐个能命中（大小写不敏感）。"""
    assert crp.ORG.search(line), f"英文机构词没认出来：{line}"


def test_plain_english_author_line_is_not_an_organization():
    """机构特征不能宽到把普通英文作者行当成单位。"""
    assert not crp.ORG.search("John Smith")
