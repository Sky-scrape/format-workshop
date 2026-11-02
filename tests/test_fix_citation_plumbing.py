"""F20 / PB1 / PB2 回归钉：引用链路的管道接通。

三条线，各自钉一件事：

- **F20 模板键迁移**：DESIGN §1.4 与 citation_ref.format_entry 读
  `citation.reference_list.templates`，但入库三刊把条目模板放在别名键下
  （tjyj：cn_formats/en_formats；cjyj：formats），UI 的引用勾选闸
  （app.details_context 读 templates）永远开不出引用转换。修法取方向①
  数据迁移：键迁到 templates、格式串与出处逐字保留；format_entry 只认带
  {字段} 占位符的模板——GB/T 7714 的描述式著录写法是给人看的格式说明，
  拿去改写条目等于把真文献替换成一句说明书。
- **PB1 括号与分隔符**：合并引用的「（」「；」「）」原先硬编码，规则集
  citation.in_text.bracket / multiple_separator 无消费者，英文刊会产出
  全角括号。现从规则集读，缺省回退原字面量（中文刊逐字节不变）；声明
  只在 citation.system == "author-year" 时消费——编号制刊物的 bracket「[]」
  描述的是编号标注，不是作者-年份写法。
- **PB2 字间距进审计**：fmt_snapshot 原先不采 rPr 的 w:spacing，而
  set_run 会写它——声明 letter_spacing_pt 的刊（tjyj 0.6 / ej 1.3）标题
  字间距变化不进审计。快照补上这个字段后，restyle_audited 的前后对比
  自动覆盖；两遍渲染第二遍必须零新增。
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document

from conftest import RULES, TOOLS

sys.path.insert(0, str(TOOLS))

import citation_ref as cr  # noqa: E402
from render_from_ir import (  # noqa: E402
    Audit, fmt_delta, fmt_snapshot, rewrite_citations, restyle_audited, set_run,
)

INDEX = cr.build_index([
    "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
    "[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.",
])


def load_rules(jid):
    return json.loads((RULES / f"{jid}.json").read_text(encoding="utf-8"))


# ---------- F20：模板键迁移 ----------

def test_in_library_rulesets_expose_templates():
    """入库刊的条目模板必须在代码读的键下（app 勾选闸以此为闸）。"""
    tj = load_rules("tjyj")["citation"]["reference_list"]
    assert "cn_formats" not in tj and "en_formats" not in tj
    # 格式串逐字未动：键迁了，值一个字符都没改
    assert tj["templates"]["cn_journal"] == \
        "作者1, 作者2. 文献题名[J]. 期刊名, 年, 卷(期): 起止页码."
    assert tj["templates"]["en_journal"].startswith("Family name Initials")
    assert tj["templates"]["en_working_paper"].endswith("Working Paper Series, Year.")

    cj = load_rules("cjyj")["citation"]["reference_list"]
    assert "formats" not in cj
    assert cj["templates"]["journal"] == \
        "[序号]主要责任者.文献题名[J].刊名，年，（期）：起止页码。"

    # gjjryj 原本就没有模板键：无键可迁，templates 缺席是如实状态
    assert not (load_rules("gjjryj")["citation"]["reference_list"]).get("templates")


def test_format_entry_renders_placeholder_templates_only():
    """占位符模板照常渲染；描述式著录写法不可渲染（返回 None，条目不动）。"""
    entry = cr.parse_entry("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    # 入库刊迁过来的描述式写法：没有 {字段}，不能拿去改写条目
    assert cr.format_entry(entry, load_rules("tjyj")) is None
    assert cr.format_entry(
        cr.parse_entry("[1] 张军.制度基础[M].上海：上海人民出版社，2018：25-40。"),
        load_rules("cjyj")) is None
    # 真正的占位符模板（mwm 结构）照常工作
    rules = {"citation": {"reference_list": {"templates": {
        "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。"}}}}
    assert cr.format_entry(entry, rules) == \
        "张三：《数字金融与企业融资》，《经济研究》，2020年第3期。"


def _build_doc(path):
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("正文第二段，合并引用[1,2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.add_paragraph("[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.")
    doc.save(str(path))
    return path


def _render(docx, rules_path, stem, citation):
    """parse → render 子进程流水（参照 test_citation_convert.py）。"""
    ir = docx.with_suffix(".ir.json")
    r = subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(docx),
                        "-o", str(ir)], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=str(TOOLS.parent))
    assert r.returncode == 0, r.stderr[-800:]
    out = docx.with_name(f"{stem}.docx")
    cmd = [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir),
           str(rules_path), "-o", str(out)]
    if citation:
        cmd.append("--citation")
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=str(TOOLS.parent))
    assert r.returncode == 0, r.stderr[-800:]
    return out, json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))


def test_in_library_ruleset_runs_conversion_end_to_end(tmp_path):
    """入库规则集（tjyj.json，非自建）勾选后端到端跑通引用转换。

    正文编号引用换成作者-年份（分隔符用该刊声明的「；」）；条目模板是
    描述式著录写法，条目文字必须原样保留——尤其不能被格式说明替换。
    """
    src = _build_doc(tmp_path / "src.docx")
    out, audit = _render(src, RULES / "tjyj.json", "out.docx", citation=True)
    texts = [p.text for p in Document(str(out)).paragraphs]
    joined = "\n".join(texts)

    assert "正文第一段，引用了文献一（张三，2020）和文献二（Smith，2021）。" in texts
    assert "正文第二段，合并引用（张三，2020；Smith，2021）。" in texts
    assert "[1]" not in "".join(t for t in texts if t.startswith("正文"))
    # 条目原样：描述式模板没被拿去改写（修复前这里会原样返回格式说明）
    assert "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10." in texts
    assert "作者1, 作者2. 文献题名" not in joined, "格式说明泄漏进了成品条目"
    # 正文改写有审计记录
    assert any(i["action"] == "正文引用改作者-年份" for i in audit["items"])


def test_workspace_gate_opens_for_in_library_ruleset(tmp_path):
    """app 的引用勾选闸读 templates：迁移后入库刊开闸，无模板刊仍关闭。"""
    import app

    d = tmp_path
    ir = {"blocks": [{"idx": 0, "type": "body", "text": "引用了文献[1]和[2]。"}]}
    (d / "manuscript.ir.json").write_text(json.dumps(ir, ensure_ascii=False),
                                          encoding="utf-8")
    offered = app.details_context(d, {"journal": "tjyj"})["citation_offered"]
    assert offered is True, "入库刊迁移后勾选闸仍打不开"

    denied = app.details_context(d, {"journal": "gjjryj"})["citation_offered"]
    assert denied is False, "没有条目模板的刊不该提供引用转换"


# ---------- PB1：括号与分隔符来自规则集 ----------

HALFWIDTH = {"citation": {"system": "author-year",
                          "in_text": {"bracket": "()", "multiple_separator": "; "}}}


def test_declared_bracket_and_separator_are_consumed():
    doc = Document()
    p = doc.add_paragraph("合并引用[1,2]与单引用[2]。")
    assert rewrite_citations(p, INDEX, HALFWIDTH) == 2
    assert p.text == "合并引用(张三，2020; Smith，2021)与单引用(Smith，2021)。"


def test_default_behaviour_is_byte_identical():
    """缺省（规则集未声明）与硬编码时期逐字节一致。"""
    for rules in ({}, load_rules("gjjryj")):   # gjjryj 的 citation.in_text 未声明括号与分隔符
        doc = Document()
        p = doc.add_paragraph("合并引用[1,2]。")
        rewrite_citations(p, INDEX, rules)
        assert p.text == "合并引用（张三，2020；Smith，2021）。"


def test_numbered_bracket_declaration_is_not_misused_for_author_year():
    """编号制刊物的 bracket「[]」描述编号标注，作者-年份转换不得错拿。

    cjyj（system=numbered）声明 bracket「[]」：那说的是它正文里 [1] 式标注的
    括号；转换产出的是作者-年份写法，用「[]」等于产出刊物从未声明的形态，
    也违反「中文刊行为不变」。
    """
    rules = load_rules("cjyj")
    assert rules["citation"]["system"] == "numbered"
    doc = Document()
    p = doc.add_paragraph("合并引用[1,2]。")
    rewrite_citations(p, INDEX, rules)
    assert p.text == "合并引用（张三，2020；Smith，2021）。"
    assert "[" not in p.text and "]" not in p.text


def test_declared_fullwidth_separator_is_consumed():
    """tjyj 声明 multiple_separator「；」：值来自规则集而非代码。"""
    doc = Document()
    p = doc.add_paragraph("合并引用[1,2]。")
    rewrite_citations(p, INDEX, load_rules("tjyj"))
    assert p.text == "合并引用（张三，2020；Smith，2021）。"


def test_in_text_label_uses_declared_bracket():
    entry = {"authors": ["Smith J"], "year": "2021", "lang": "en"}
    assert cr.in_text_label(entry, HALFWIDTH) == "(Smith，2021)"
    assert cr.label_body(entry, HALFWIDTH) == "Smith，2021"
    assert cr.in_text_label(entry) == "（Smith，2021）"
    assert cr.label_body(entry) == "Smith，2021"


def test_halfwidth_bracket_end_to_end(tmp_path):
    """自建规则集声明半角括号：整条流水线产出 (A，2020; B，2021) 形态。"""
    rules = load_rules("tjyj")
    rules["citation"]["in_text"] = dict(HALFWIDTH["citation"]["in_text"])
    rp = tmp_path / "halfwidth.json"
    rp.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")

    src = _build_doc(tmp_path / "src.docx")
    out, _audit = _render(src, rp, "out.docx", citation=True)
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "正文第二段，合并引用(张三，2020; Smith，2021)。" in texts
    assert not any("（张三" in t for t in texts), "声明了半角就不能再出全角括号"


# ---------- PB2：字间距进快照与审计 ----------

def test_fmt_snapshot_captures_letter_spacing():
    from docx.oxml.ns import qn

    doc = Document()
    p = doc.add_paragraph("标题")
    set_run(p.runs[0], cn="黑体", latin="Times New Roman", size=16, bold=False)
    before = fmt_snapshot(p)
    assert before["spacing"] == set()
    set_run(p.runs[0], letter_spacing=1.3)      # ej 的标题字间距：1.3pt = 26 twips
    after = fmt_snapshot(p)
    assert after["spacing"] == {"26"}
    assert before != after, "仅字间距变了仍被判成没变化，审计漏记（PB2）"
    assert fmt_delta(before, after, "16pt") == "改字间距 → 16pt"
    # 没有字间距的 run 不虚报
    plain = fmt_snapshot(doc.add_paragraph("普通"))
    assert plain["spacing"] == set()


def test_restyle_audited_reports_spacing_only_change():
    """段落其余全部合规、只差字间距时，也要出审计条目。"""
    kw = dict(cn="黑体", latin="Times New Roman", size=16, align="center",
              line_pt=24, bold=False)
    doc = Document()
    p = doc.add_paragraph("标题")
    # 首遍建立目标版式（原始段落处处不合规，出条目是应当的）
    first = Audit()
    restyle_audited(p, first, 1, "front_matter", "文章标题套版", "16pt",
                    "typography.title", **kw)
    assert len(first.items) == 1
    # 同参数重跑必须零新增——快照与写入往返幂等，也是「两遍渲染零新增」的单元面
    rerun = Audit()
    restyle_audited(p, rerun, 1, "front_matter", "文章标题套版", "16pt",
                    "typography.title", **kw)
    assert rerun.items == [], "同参数重跑仍虚报改动（快照不幂等）"
    # 只差字间距：恰好一条，且改动名里写「字间距」
    spaced = Audit()
    restyle_audited(p, spaced, 1, "front_matter", "文章标题套版", "16pt",
                    "typography.title", letter_spacing=0.6, **kw)
    assert len(spaced.items) == 1 and "字间距" in spaced.items[0]["note"], spaced.items
    # 字间距已写到位：再跑不得再出条目
    again = Audit()
    restyle_audited(p, again, 1, "front_matter", "文章标题套版", "16pt",
                    "typography.title", letter_spacing=0.6, **kw)
    assert again.items == []


def test_letter_spacing_audited_once_and_second_pass_clean(tmp_path):
    """声明 letter_spacing_pt 的刊（tjyj 0.6）：首遍进审计，第二遍零新增。"""
    src = _build_doc(tmp_path / "src.docx")
    out1, audit1 = _render(src, RULES / "tjyj.json", "pass1.docx", citation=False)
    title_notes = [i["note"] for i in audit1["items"] if i["action"] == "文章标题套版"]
    assert title_notes and any("字间距" in n for n in title_notes), \
        "letter_spacing_pt 的变化没进审计"

    out2, audit2 = _render(out1, RULES / "tjyj.json", "pass2.docx", citation=False)
    assert audit2["items"] == [], f"第二遍渲染新增了 {len(audit2['items'])} 条改动"
