"""五个已确认问题的回归钉：文本改写口径（F01/F02/F03/F04/F09）。

每个用例都来自修复前的真实复现：

- F01  段落聚合文本（p.text）与物理 run/w:t 改写口径不一致：超链接、跨 run、
       w:ins 包裹的引用能过体检闸门却永远改不到；整段标题在链接内时
       replace_text 使文字翻倍；restyle 与 fmt_snapshot 看不见链接内 run。
- F02  标签/编号区含制表符或换行时，replace_prefix 按字符数扣 w:t 会静默
       吃掉正文一个字（「关键词\\t：数字金融」→「关键词：\\t字金融」）。
- F03  关键词分隔符归一把空格当分隔符，英文词组被拆散。
- F04  CITE 把「[0, 1]」「[90, 100]」这类数学区间当引用。
- F09  clear_para_formatting 无条件删 w:numPr，Word 自动编号静默消失。
"""

import importlib.util
import json
import re
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"

from conftest import RULES, body_texts, render_into  # noqa: E402


def _load(name, filename):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


render = _load("render_from_ir_fix_under_test", "render_from_ir.py")
sys.path.insert(0, str(TOOLS))
import citation_ref as cr  # noqa: E402


# ---------- 造段落的小工具 ----------

def new_para(*parts):
    """造段落；<TAB>/<BR> 占位符换成对应的物理节点。"""
    doc = Document()
    p = doc.add_paragraph()
    for part in parts:
        if part == "<TAB>":
            p.add_run()._r.append(OxmlElement("w:tab"))
        elif part == "<BR>":
            p.add_run().add_break()
        else:
            p.add_run(part)
    return p


def hyperlink_run(p, text, superscript=False):
    """把一个 run 装进 w:hyperlink（内部锚点链接，不需要关系件）。"""
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("w:anchor"), "sec1")
    r = OxmlElement("w:r")
    if superscript:
        rpr = OxmlElement("w:rPr")
        va = OxmlElement("w:vertAlign")
        va.set(qn("w:val"), "superscript")
        rpr.append(va)
        r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    hl.append(r)
    p._p.append(hl)
    return r


def ins_run(p, text):
    """把一个 run 装进 w:ins（修订插入）。p.text 与 p.runs 都看不见它。"""
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), "1")
    ins.set(qn("w:author"), "作者")
    ins.set(qn("w:date"), "2026-01-01")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    ins.append(r)
    p._p.append(ins)
    return r


def add_numpr(p, num_id=1, ilvl=0):
    """给段落挂直接自动编号（Word 列表）。"""
    ppr = p._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    lvl = OxmlElement("w:ilvl")
    lvl.set(qn("w:val"), str(ilvl))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(num_id))
    numpr.append(lvl)
    numpr.append(nid)
    ppr.append(numpr)


def agg_text(p):
    return render.para_full_text(p)


INDEX = cr.build_index([
    "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
    "[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.",
    "[3] 李四. 融资约束的度量[J]. 金融研究, 2019, 40(5): 1-9.",
])
RULES_BARE = {"citation": {}}


# ---------- F01：聚合口径统一 ----------

def test_citation_spanning_runs_is_rewritten():
    """跨 run 的「[1,2]」必须整段改到，正文其余文字原样。"""
    p = new_para("结果与[1,", "2]一致。")
    hits = render.rewrite_citations(p, INDEX, RULES_BARE)
    assert hits == 1
    assert agg_text(p) == "结果与（张三，2020；Smith，2021）一致。"


def test_citation_inside_hyperlink_is_rewritten_and_unsuperscripted():
    """超链接内的上标引用要改到，且落点 run 转回正常字号。"""
    p = new_para()
    hyperlink_run(p, "见文献", superscript=True)
    hyperlink_run(p, "[3]", superscript=True)
    tail = p.add_run("。")

    hits = render.rewrite_citations(p, INDEX, RULES_BARE)
    assert hits == 1
    assert agg_text(p) == "见文献（李四，2019）。"
    # 改写落点是「[3]」所在的 run：它的上标必须摘掉，「见文献」的不动
    for el in p._p.xpath(".//w:r"):
        txt = "".join(t.text or "" for t in el.findall(qn("w:t")))
        rpr = el.find(qn("w:rPr"))
        va = rpr is not None and len(rpr.findall(qn("w:vertAlign")))
        if txt == "见文献":
            assert va, "未改写的 run 不该被动格式"
        elif "李四" in txt:
            assert not va, "作者-年份写法必须转回正常字号"
    assert tail.text == "。"


def test_citation_inside_tracked_insert_is_seen_by_both_ends():
    """w:ins 包裹的引用 p.text 完全看不见；体检与改写都必须认到它。"""
    p = new_para()
    ins_run(p, "修订插入的[1]")
    assert p.text == ""                      # 旧口径下两端都瞎
    assert agg_text(p) == "修订插入的[1]"
    _, problems = cr.check([agg_text(p)], INDEX, RULES_BARE)
    assert not problems, "闸门必须看得见 w:ins 里的引用"
    hits = render.rewrite_citations(p, INDEX, RULES_BARE)
    assert hits == 1
    assert agg_text(p) == "修订插入的（张三，2020）"


def test_citation_inside_tab_spans_tab_is_dropped():
    """命中区里的制表符属于被替换的文字，随引用一起删掉。"""
    p = new_para("前文[1", "<TAB>", ",3]后文。")
    hits = render.rewrite_citations(p, INDEX, RULES_BARE)
    assert hits == 1
    assert agg_text(p) == "前文（张三，2020；李四，2019）后文。"
    assert not p._p.xpath(".//w:r/w:tab")


def test_replace_text_on_all_hyperlink_paragraph_does_not_double():
    """整段文字都在链接内时，replace_text 直接改链接里的节点，不再另起 run。"""
    p = new_para()
    hyperlink_run(p, "一、引言")
    render.replace_text(p, "一、引言")
    assert agg_text(p) == "一、引言"
    assert p.text == "一、引言" or len(p.runs) == 0     # 没有「引言一、引言」
    assert len(p._p.xpath(".//w:t")) == 1               # 文字节点没翻倍


def test_restyle_covers_runs_inside_hyperlink():
    """restyle 与 fmt_snapshot 必须看见链接内 run，否则版式漏套还查不到。"""
    p = new_para("普通")
    hyperlink_run(p, "链接")
    render.restyle(p, cn="宋体", latin="Times New Roman", size=10.5, align="both")
    sizes = set()
    for el in p._p.xpath(".//w:r"):
        rpr = el.find(qn("w:rPr"))
        sz = None if rpr is None else rpr.find(qn("w:sz"))
        assert sz is not None and sz.get(qn("w:val")) == "21", \
            "链接内 run 没有套上 10.5pt（半磅 21）"
        sizes.add(sz.get(qn("w:val")))
    snap = render.fmt_snapshot(p)
    assert snap["sizes"] == {10.5}, "fmt_snapshot 看不见链接内 run 的字号"


# ---------- F02：前缀/后缀替换的物理口径 ----------

def test_label_with_tab_keeps_body_chars():
    """证据复现：「关键词\\t：数字金融；实证」不得吃掉「数」。"""
    p = new_para("关键词", "<TAB>", "：数字金融；实证")
    old = agg_text(p)
    m = re.match(r"^(关\s*键\s*词)\s*[:：]", old)
    want = "关键词："
    new = want + old[m.end():]
    render.replace_prefix(p, m.end(), want)
    assert agg_text(p) == new == "关键词：数字金融；实证"


def test_label_with_line_break_keeps_body_chars():
    """证据复现：「摘\\n要：数字经济的测度」不得吃掉「数」。"""
    p = new_para("摘", "<BR>", "要：数字经济的测度")
    old = agg_text(p)
    m = re.match(r"^(摘\s*要)\s*[:：]", old)
    render.replace_prefix(p, m.end(), "摘要：")
    assert agg_text(p) == "摘要：数字经济的测度"


def test_ref_number_with_tab_keeps_authors():
    """证据复现：文献编号「1\\t何清…」改写法后「何」必须还在。"""
    p = new_para("1", "<TAB>", "何清，2020：《数字金融》。")
    prefix, num = render.ref_number_prefix(agg_text(p))
    assert (prefix, num) == ("1\t", 1)
    render.replace_prefix(p, len(prefix), render.format_ref_number("[1]", num))
    assert agg_text(p) == "[1]何清，2020：《数字金融》。"


def test_label_tab_case_end_to_end(tmp_path):
    """端到端：标签归一的审计 after 必须与成品逐字一致，正文一字不吞。"""
    src = tmp_path / "tab.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    p = doc.add_paragraph()
    p.add_run("关键词")
    p.runs[0]._r.append(OxmlElement("w:tab"))
    p.add_run("：数字金融；实证")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    out = {i: t for i, t in result.output_texts.items()
           if result.block_type.get(i) == "keywords"}
    assert list(out.values()) == ["关键词：数字金融；实证"], "标签区的制表符吞了正文"

    # 审计条目 id 现为「改前文本指纹|动作」（render_from_ir.Audit.key）；
    # 按「块序号＋动作名」定位条目，不耦合键的具体格式。
    item = next((i for i in result.items
                 if i["action"] == "标签归一" and i["idx"] == list(out)[0]), None)
    assert item and not item["rejected"]
    assert item["before"] == "关键词\t：数字金融；实证"
    assert item["after"] == out[list(out)[0]], "审计 after 与成品不符"


# ---------- F03：关键词分隔符 ----------

def test_split_keywords_keeps_english_phrases():
    """证据复现：单个空格是英文词间空格，不能当分隔符。"""
    assert render.split_keywords("数字金融 digital finance；货币政策", "；") == \
        ["数字金融 digital finance", "货币政策"]


def test_split_keywords_still_normalizes_double_space_separated_lists():
    """目标分隔符非空白时，连续两个以上的空白仍算分隔符——gjjryj 双空格
    关键词表换到 tjyj 的「；」这一既有归一不能退。"""
    assert render.split_keywords("宽松货币政策  美联储  金砖国家", "；") == \
        ["宽松货币政策", "美联储", "金砖国家"]
    assert render.split_keywords("数字金融　绿色金融", "；") == \
        ["数字金融", "绿色金融"]


def test_split_keywords_when_target_separator_is_whitespace():
    """目标刊的分隔符本身就是空白时，一切空白都参与切分。"""
    assert render.split_keywords("a b，c", "  ") == ["a", "b", "c"]


def test_english_phrase_keywords_untouched_end_to_end(tmp_path):
    """端到端：含英文词组的关键词行不被拆散，也不产生分隔符审计条目。"""
    src = tmp_path / "kw.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融 digital finance；货币政策")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    kws = [t for i, t in result.output_texts.items()
           if result.block_type.get(i) == "keywords"]
    assert kws == ["关键词：数字金融 digital finance；货币政策"]
    assert "关键词分隔符归一" not in {it["action"] for it in result.items}


# ---------- 端到端公共件（--citation 流程） ----------

def run_with_citation(tmp_path, src, rules_path):
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(src),
                    "-o", str(src.with_suffix(".ir.json"))],
                   capture_output=True, text=True, check=True)
    r = subprocess.run([sys.executable, str(TOOLS / "render_from_ir.py"),
                        str(src.with_suffix(".ir.json")), str(rules_path),
                        "-o", str(src.with_name("out.docx")), "--citation"],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr[-800:]
    return (json.loads((src.with_name("out.audit.json")).read_text(encoding="utf-8")),
            Document(str(src.with_name("out.docx"))))


TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
    "en_journal": "{authors}，{year}，“{title}”，{journal}{volume_part}{page_part}.",
    "en_book": "{authors}，{year}，“{title}”，{publisher}.",
    "en_other": "{authors}，{year}，“{title}”，{source}.",
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


# ---------- F04：数学区间不是引用 ----------

@pytest.mark.parametrize("text", [
    "x ∈ [0, 1] 且 y ∈ [90, 100]",
    "占比在 [90, 100]%",
    "误差 ±[2, 3] 之间",
])
def test_math_interval_is_not_a_citation(text):
    assert cr.citation_spans(text) == [], f"{text!r} 被当成了引用"


@pytest.mark.parametrize("text,expected", [
    ("如文献[3]所示，另有[5, 6]。", ["[3]", "[5, 6]"]),
    # 含 0 的括号一律不是引用：文献表从 [1] 起编，[0]、[0, 1] 只会是取值
    # 区间或笔误——不再靠「[0] 不在文献表」的体检误报兜底（理由有误导性）
    ("覆盖了[0, 1]区间的样本", []),
    ("回归系数在 [0,1]内变动", []),
    # 左邻是 %/数字/‰ 的单数字括号是真引用（「10%[12]」「图3[4]」），
    # 数学邻接符号只拦多数字区间形式
    ("增长10%[12]的结论。", ["[12]"]),
    ("如图3[4]所示。", ["[4]"]),
    ("as shown in [12]", ["[12]"]),           # 英文稿的正常引用不能堵死
])
def test_real_citations_still_found(text, expected):
    assert [m.group(0) for m in cr.citation_spans(text)] == expected


def test_interval_with_real_entry_numbers_not_rewritten(tmp_path):
    """端到端：括号里恰好都是已收录条目号时，∈ 区间也不得改成作者-年份。"""
    src = tmp_path / "interval.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("样本取值区间表明 x ∈ [90, 100]。")
    doc.add_paragraph("见文献[1]、[2]、[90]与[100]。")
    doc.add_paragraph("参考文献")
    for it in [
        "[1] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] 李四. 融资约束研究[J]. 金融研究, 2019, 40(5): 1-9.",
        "[90] 王五. 银行风险研究[J]. 财贸经济, 2018, 39(6): 1-8.",
        "[100] 赵六. 货币政策研究[J]. 世界经济, 2017, 40(9): 1-25.",
    ]:
        doc.add_paragraph(it)
    doc.save(str(src))

    audit, out = run_with_citation(tmp_path, src, rules_with_templates(tmp_path))
    joined = "\n".join(p.text for p in out.paragraphs)
    assert "∈ [90, 100]" in joined, "数学区间被改写成了作者-年份"
    assert "（张三，2020）" in joined and "（赵六，2017）" in joined, \
        "真引用没有被转换"
    assert not [i for i in audit["items"] if i["action"] == "引用体系转换（已跳过）"], \
        "区间被误当成引用后整批跳过"


# ---------- F01 端到端：超链接里的引用不再漏改 ----------

def test_hyperlink_citation_is_converted_end_to_end(tmp_path):
    """闸门放行的超链接引用必须真的改到；条目重排后不得残留错位编号。"""
    src = tmp_path / "hl.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    p = doc.add_paragraph()
    hyperlink_run(p, "见文献[3]和")
    p.add_run("[1]与[2]。")
    doc.add_paragraph("参考文献")
    for it in [
        "[1] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] 李四. 融资约束研究[J]. 金融研究, 2019, 40(5): 1-9.",
        "[3] 王五. 银行风险研究[J]. 财贸经济, 2018, 39(6): 1-8.",
    ]:
        doc.add_paragraph(it)
    doc.save(str(src))

    audit, out = run_with_citation(tmp_path, src, rules_with_templates(tmp_path))
    body = [t for i, t in body_texts(str(src.with_name("out.docx"))).items()]
    joined = "\n".join(body)

    assert "（王五，2018）" in joined and "（张三，2020）" in joined, \
        f"链接内/跨 run 的引用没改到：{joined[:200]}"
    assert not re.search(r"\[\d\]", joined), f"残留编号指向错位条目：{joined[:200]}"
    actions = {i["action"] for i in audit["items"]}
    assert "正文引用改作者-年份" in actions


# ---------- F09：自动编号不再无声消失 ----------

def test_body_paragraph_keeps_auto_numbering(tmp_path):
    """被判 body 的列表段：渲染后 numPr 必须还在，文字一字不动。"""
    src = tmp_path / "num.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；货币政策")
    p = doc.add_paragraph(
        "本文考察数字金融对企业融资约束的影响，这一关系在民营企业中尤为显著。")
    add_numpr(p)
    p2 = doc.add_paragraph("稳健性检验表明上述结论在更换度量方式后依然成立，"
                           "且不依赖于样本区间的选取。")
    add_numpr(p2)
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    out_doc = Document(str(result.docx))

    numbered = []
    for para in out_doc.paragraphs:
        ppr = para._p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            numbered.append(para.text)
    assert len(numbered) == 2, f"body 列表段的自动编号被删了：{len(numbered)} 段还在"
    assert numbered[0].startswith("本文考察数字金融")
    assert numbered[1].startswith("稳健性检验")


def test_heading_numbering_is_dropped_when_rewritten_as_literal(tmp_path):
    """标题分支的清除是必要的：编号已重写为字面文字，numPr 留着会双重编号。"""
    src = tmp_path / "head.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；货币政策")
    body = doc.add_paragraph("本章考察数字金融对企业融资约束的影响，样本为A股上市公司。")
    add_numpr(body)
    h = doc.add_paragraph("1 研究设计")
    add_numpr(h)
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    out_doc = Document(str(result.docx))

    texts = {t for t in body_texts(str(result.docx)).values()}
    assert any(t.startswith("一、研究设计") for t in texts), \
        f"标题没有重写为字面编号：{sorted(texts)[:6]}"
    for para in out_doc.paragraphs:
        if para.text.startswith("一、研究设计"):
            ppr = para._p.find(qn("w:pPr"))
            assert ppr is None or ppr.find(qn("w:numPr")) is None, \
                "标题重写为字面编号后自动编号还在，会双重编号"
