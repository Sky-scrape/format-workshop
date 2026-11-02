"""F04 返工回归钉：数字取值区间不再被当成引用（语境约束按形式分面）。

上一轮修复只拦数学邻接符号：CJK 前邻的逗号区间（括号内恰好都是已收录且
被引条目）仍被整段改写成作者-年份——语义级错改正文。且左邻排除对单数字
括号也生效，把「10%[12]」「图3[4]」「5‰[13]」这类正常引用也丢了（假阴性，
体检还把对应条目错记为「未被引用」）。

钉住的契约：
1. 含 0 的括号一律不是引用：文献表从 [1] 起编（GB/T 7714），
   [0]、[0, 1] 只会是取值区间或笔误，任何语境都拦下
2. 多数字括号在数学邻接符号、区间语境词下排除
3. ∈/≈/≡ 对所有形式生效；其余数学符号只拦多数字形式
4. check 与改写同一口：区间不再产生「正文引用了 [0]」的误导报错
5. 端到端：括号内恰为已收录条目号的区间不被改写，真引用照常转换
"""

import importlib.util
import json
import re
import shutil
import sys
from pathlib import Path

import pytest
from docx import Document

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"

from conftest import RULES, run_tool  # noqa: E402


def _load(name, filename):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


cr = _load("citation_ref_f04", "citation_ref.py")


# ---------- 含 0：任何语境都不是引用 ----------

@pytest.mark.parametrize("text", [
    "覆盖了[0, 1]区间的样本",
    "回归系数在 [0,1]内变动",
    "正文引用了[0]。",                     # 单独的 [0] 也不会再被体检误报
    "占比落在[0, 100]%之间。",
])
def test_zero_containing_bracket_is_not_citation(text):
    assert cr.citation_spans(text) == [], f"{text!r} 被当成了引用"


# ---------- 多数字括号：区间语境词 ----------

@pytest.mark.parametrize("text", [
    "样本量覆盖[90, 100]家企业的区间。",
    "参数取值在[44, 55]之间。",
    "行业集中度位于[90, 100]这一范围。",
])
def test_interval_words_block_multi_number(text):
    assert cr.citation_spans(text) == [], f"{text!r} 被当成了引用"


# ---------- 多数字括号：数学邻接符号（既有行为不退） ----------

@pytest.mark.parametrize("text", [
    "x ∈ [90, 100]",
    "占比在 [90, 100]%",
    "误差 ±[2, 3] 之间",
    "取值90[90, 100]",
])
def test_math_symbols_still_block_multi_number(text):
    assert cr.citation_spans(text) == [], f"{text!r} 被当成了引用"


# ---------- 单数字括号：数学邻接符号不再误伤 ----------

@pytest.mark.parametrize("text,expected", [
    ("增长10%[12]的结论。", ["[12]"]),
    ("误差为5‰[13]。", ["[13]"]),
    ("如图3[4]所示。", ["[4]"]),
    ("参见文献 [21]，其中[5, 6]是核心。", ["[21]", "[5, 6]"]),
])
def test_single_number_citation_after_math_char_restored(text, expected):
    got = [m.group(0) for m in cr.citation_spans(text)]
    assert got == expected, (text, got, expected)


# ---------- 不带区间语境的多数字括号仍是引用 ----------

@pytest.mark.parametrize("text,expected", [
    ("方法见[12, 15]所示。", ["[12, 15]"]),
    ("综述见[1-5]。", ["[1-5]"]),
    ("如文献[3]所示，另有[5, 6]。", ["[3]", "[5, 6]"]),
    ("as shown in [12]", ["[12]"]),          # 英文稿的正常引用不能堵死
])
def test_multi_number_citation_without_interval_context_kept(text, expected):
    got = [m.group(0) for m in cr.citation_spans(text)]
    assert got == expected, (text, got, expected)


# ---------- check 与 citation_spans 同口：误导性报错消失 ----------

def test_check_no_longer_misreports_zero_as_missing_entry():
    """「覆盖了[0, 1]区间」不再报「正文引用了 [0]，但文献表里没有这一条」。"""
    index = cr.build_index([
        "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.",
    ])
    cited, problems = cr.check(["覆盖了[0, 1]区间的样本"], index, {"citation": {}})
    assert not any("[0]" in p for p in problems), problems
    assert 0 not in cited


# ---------- 端到端：真实管线 + 带双向体检的规则集 ----------

TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
}


def rules_with_templates(tmp_path):
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["citation"]["system"] = "author-year"
    rl = rules["citation"].setdefault("reference_list", {})
    rl["templates"] = TEMPLATES
    rl["must_require_bidirectional_match"] = True
    path = tmp_path / "cite.json"
    path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return path


def test_interval_with_cited_entry_numbers_not_rewritten(tmp_path):
    """括号里恰好都是已收录且被引条目号的区间：不改写；真引用照常转换。

    修复前：「样本量覆盖[90, 100]家企业的区间」被改成「（作者90，1990；
    作者100，2000）家企业的区间」——语义级错改。
    """
    src = tmp_path / "interval.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文引用了文献[44]、[55]、[90]与[100]。")
    doc.add_paragraph("样本量覆盖[90, 100]家企业的区间。参数取值在[44, 55]之间。")
    doc.add_paragraph("参考文献")
    for it in [
        "[44] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.",
        "[55] 李四. 融资约束研究[J]. 金融研究, 2019, 40(5): 1-9.",
        "[90] 王五. 银行风险研究[J]. 财贸经济, 2018, 39(6): 1-8.",
        "[100] 赵六. 货币政策研究[J]. 世界经济, 2017, 40(9): 1-25.",
    ]:
        doc.add_paragraph(it)
    doc.save(str(src))

    d = tmp_path / "run"
    d.mkdir()
    shutil.copy(src, d / "manuscript.docx")   # ir 与 docx 同名同目录，渲染器按此找源稿
    run_tool("parse_docx.py", d / "manuscript.docx", "-o", d / "manuscript.ir.json")
    run_tool("render_from_ir.py", d / "manuscript.ir.json",
             rules_with_templates(tmp_path), "-o", d / "out.docx", "--citation")

    out = Document(str(d / "out.docx"))
    joined = "\n".join(p.text for p in out.paragraphs)
    assert "样本量覆盖[90, 100]家企业的区间" in joined, \
        f"区间被改写成了作者-年份：{joined[:300]}"
    assert "参数取值在[44, 55]之间" in joined, \
        f"区间被改写成了作者-年份：{joined[:300]}"
    assert "（张三，2020）" in joined and "（赵六，2017）" in joined, \
        f"真引用没有被转换：{joined[:300]}"
    assert not re.search(r"（作者\d+，\d{4}）", joined), "出现了解析失败的占位改写"

    audit = json.loads((d / "out.audit.json").read_text(encoding="utf-8"))
    assert not [i for i in audit["items"]
                if i["action"] == "引用体系转换（已跳过）"], \
        "区间被误当成引用后整批跳过"
