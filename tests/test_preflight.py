"""输入稿版式体检（tools/preflight.py）的验收。

两条最关键的约束在这里：
1. 幂等闸：先用 render_from_ir 按 tjyj 排一遍，再拿成品过体检，fix=="auto"
   必须为 0。它证明「体检报 auto 的项，在渲染器成品上已经不再出现」；它
   不覆盖体检根本没建检查项的属性（如标题字间距），所以另有字间距专项和
   「渲染器二次套版不改字」两条补位。
2. 只读闸：体检前后源 docx 的 sha256 必须相同。
"""

import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

import preflight
from rules_ref import resolvable

ROOT = Path(__file__).resolve().parent.parent
RULES = ROOT / "rules" / "journal"


def load_rules(jid="tjyj"):
    return json.loads((RULES / f"{jid}.json").read_text(encoding="utf-8"))


def _run(p, text, cn="宋体", latin="Times New Roman", size=10.5):
    r = p.add_run(text)
    r.font.name = latin
    r.font.size = Pt(size)
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), cn)
    return r


def _field_paragraph(doc, instr, text):
    """插一个 w:fldSimple 复杂域（EndNote 引文就是这种形态）。"""
    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    fld.append(r)
    p._p.append(fld)
    return p


@pytest.fixture()
def wrong_docx(tmp_path):
    """一份「处处写错」的稿子：正文字体/字号/行距/对齐、页边距、表线都对不上。"""
    path = tmp_path / "wrong.docx"
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Mm(10)
    sec.bottom_margin = Mm(10)
    sec.left_margin = Mm(10)
    sec.right_margin = Mm(10)

    _run(doc.add_paragraph(), "美国宽松货币政策对金砖国家的溢出效应研究",
         cn="宋体", latin="宋体", size=12)

    # 正文必须够长：parse_docx 会把前三段里不足 80 字的短段当前置区
    p = doc.add_paragraph()
    _run(p, "本文以美联储及金砖五国宏观月度数据为样本，考察美国宽松货币政策"
            "对金砖国家宏观经济与金融市场的溢出效应，并检验其传导渠道与异质性。"
            "这一段故意写成楷体十五磅、左对齐、二十八磅行距，用来验证体检能报偏差。",
         cn="楷体", latin="Arial", size=15)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = Pt(28)

    p = doc.add_paragraph()
    _run(p, "一、引言", cn="宋体", latin="宋体", size=9)

    # 再垫一段，让「表1」不落在前三段的位置推断里；相邻表格由 bind_captions 认题注
    p = doc.add_paragraph()
    _run(p, "第二段正文也故意用黑体十二磅，行距二十八磅，用于核验正文套版的判定。",
         cn="黑体", latin="Times New Roman", size=12)
    p.paragraph_format.line_spacing = Pt(28)

    _run(doc.add_paragraph(), "表1 主要变量描述性统计", cn="宋体", latin="宋体", size=10.5)

    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    for row in table.rows:
        for cell in row.cells:
            cell.text = "单元格"
            for r in cell.paragraphs[0].runs:
                r.font.name = "宋体"
                r.font.size = Pt(12)
                r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")

    _run(doc.add_paragraph(), "注：括号内为标准差。", cn="宋体", latin="宋体", size=10.5)
    doc.save(str(path))
    return path


def _auto_items(res):
    return {(it["block"], it["property"]): it
            for it in res["items"] if it["fix"] == "auto"}


def test_wrong_docx_reports_items_and_rules(wrong_docx):
    rules = load_rules()
    res = preflight.analyze(str(wrong_docx), rules)

    autos = _auto_items(res)
    # 正文：显式写了楷体/15pt/左对齐，都应判 auto
    assert ("body", "中文字体") in autos
    item = autos[("body", "中文字体")]
    assert item["expected"] == "宋体"
    assert "楷体" in item["actual"]
    assert item["rule"] and resolvable(rules, item["rule"])
    # 条目字段齐全
    for key in ("idx", "section", "block", "property", "rule",
                "expected", "actual", "fix", "note"):
        assert key in item
    assert ("body", "字号") in autos
    assert autos[("body", "字号")]["expected"] == "10.5pt"
    assert ("body", "对齐") in autos
    # 页面：10mm 的页边距与 tjyj 的上 22 / 下 20 / 左 25 / 右 24 不符
    assert ("page", "上边距") in autos
    assert ("page", "纸张") in autos
    # 表格：Table Grid 的表级边框与三线表不符；表内 12pt 也不是 7.5pt
    assert ("table", "三线表边框") in autos
    assert ("table", "表内字号") in autos
    assert autos[("table", "表内字号")]["expected"] == "7.5pt"


def test_rendered_output_has_no_auto_fix(baseline):
    """幂等闸：按 tjyj 排过的成品，体检不应再报任何「自动修正」。

    体检的期望值全部取自 render_from_ir 的分支；表格边框更是直接在 XML 深拷贝
    上跑 apply_table 取目标快照。两边一旦漂移，这里就会出现 auto 条目。
    """
    rules = load_rules("tjyj")
    res = preflight.analyze(str(baseline.docx), rules)
    autos = [(it["idx"], it["block"], it["property"], it["expected"], it["actual"])
             for it in res["items"] if it["fix"] == "auto"]
    assert autos == []
    # 成品是渲染器自己写的显式格式，应当没有「随样式」可判
    assert res["stats"]["inherited"] == 0
    assert res["coverage"]["blocks_checked"] == res["coverage"]["blocks_total"]


def test_title_letter_spacing_is_checked(baseline, tmp_path):
    """标题字间距：渲染器会写 w:spacing，体检必须也有这一项。

    tjyj 的 typography.title.letter_spacing_pt = 0.6；成品标题应被体检为
    「字间距 0.6pt / ok」。把成品标题的字间距清掉后，体检必须报 auto——
    否则「体检说合规、排版却改字」的口子就会从这里漏过去。
    """
    rules = load_rules()
    res = preflight.analyze(str(baseline.docx), rules)
    ok = [it for it in res["items"]
          if it["block"] == "front_matter" and it["property"] == "字间距"
          and it["fix"] == "ok"]
    assert ok, "成品标题的 0.6pt 字间距没有被体检识别"

    # 源稿没有显式字间距 → inherited（可能随样式），不能再当成「没有这一项」
    broken = tmp_path / "no-spacing.docx"
    shutil.copy2(baseline.docx, broken)
    doc = Document(str(broken))
    for r in doc.paragraphs[0].runs:
        rpr = r._element.find(qn("w:rPr"))
        sp = rpr.find(qn("w:spacing")) if rpr is not None else None
        if sp is not None:
            rpr.remove(sp)
    doc.save(str(broken))

    res2 = preflight.analyze(str(broken), rules)
    inherited = [it for it in res2["items"]
                 if it["block"] == "front_matter" and it["property"] == "字间距"
                 and it["fix"] == "inherited"]
    assert inherited, "标题没有显式字间距时体检没有建「随样式」项"
    assert inherited[0]["rule"] == "typography.title.letter_spacing_pt"

    # 显式写成 0（模板切到 0pt 的刊）：渲染器会清除，体检必须报 auto
    wrong = tmp_path / "zero-spacing.docx"
    shutil.copy2(baseline.docx, wrong)
    doc = Document(str(wrong))
    for r in doc.paragraphs[0].runs:
        rpr = r._element.get_or_add_rPr()
        old = rpr.find(qn("w:spacing"))
        if old is not None:
            rpr.remove(old)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), "0")
        rpr.append(sp)
    doc.save(str(wrong))

    res3 = preflight.analyze(str(wrong), rules)
    autos = [it for it in res3["items"]
             if it["block"] == "front_matter" and it["property"] == "字间距"
             and it["fix"] == "auto"]
    assert autos, "显式 0 字间距没有被体检报 auto"
    assert autos[0]["rule"] == "typography.title.letter_spacing_pt"


def test_figures_border_is_declared_not_handled(baseline):
    """图边框：渲染器不处理，体检必须在 not_handled 里明说，不能假装检查过。"""
    res = preflight.analyze(str(baseline.docx), load_rules())
    rows = [r for r in res["not_handled"] if r["rule"] == "figures.border"]
    assert rows, "figures.border 没有进 not_handled"
    assert rows[0]["fix"] == "manual"


def test_renderer_does_not_change_its_own_output(baseline, tmp_path):
    """渲染器二次套版不改字：把成品再排一遍，changed=True 的审计条目必须为 0。

    体检的 auto==[] 只说明「体检报的项都已落地」；本用例从渲染器一侧补上
    「已经合规的成品不会再被改」，两条一起才盖住「体检漏建检查项」的缺口。
    """
    from conftest import render_into

    again = render_into(tmp_path / "second-pass", baseline.docx)
    changed = [it for it in again.items if it.get("changed", True)]
    assert changed == [], f"二次套版仍改动了 {len(changed)} 处：{changed[:3]}"


def test_preflight_does_not_touch_source(wrong_docx):
    before = hashlib.sha256(wrong_docx.read_bytes()).hexdigest()
    preflight.analyze(str(wrong_docx), load_rules())
    after = hashlib.sha256(wrong_docx.read_bytes()).hexdigest()
    assert before == after


def test_field_and_pseudo_table_are_unchecked(tmp_path):
    path = tmp_path / "fields.docx"
    doc = Document()
    doc.add_paragraph("测试标题")
    p = doc.add_paragraph()
    _run(p, "这是正文段落，刻意写得足够长，避开解析器把前三段短句当前置区的位置推断，"
            "后面紧跟一个引文域和一个伪表格公式，用来核验体检的未检查清单。")
    _field_paragraph(doc, "ADDIN EN.CITE", "（张三，2020）")
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "y = α + βx"
    table.rows[0].cells[1].text = "（1）"
    doc.save(str(path))

    res = preflight.analyze(str(path), load_rules())
    unchecked = res["unchecked"]
    assert any("citation" in u["reason"] for u in unchecked)
    assert any(u["block"] == "equation_pseudo_table" for u in unchecked)
    # 域锁定段与伪表格都不该被当成版式偏差
    locked_idx = {u["idx"] for u in unchecked}
    assert not [it for it in res["items"]
                if it["fix"] == "auto" and it["idx"] in locked_idx]
    assert res["stats"]["unchecked"] == len(unchecked)


def test_every_rule_reference_resolves(wrong_docx, baseline):
    """报告里的每个 rule 都要能落到规则集的一个值上（DESIGN §6/§8.5）。"""
    rules = load_rules()
    for docx in (wrong_docx, baseline.docx):
        res = preflight.analyze(str(docx), rules)
        for it in res["items"]:
            assert it["fix"] in (preflight.FIX_AUTO, preflight.FIX_MANUAL,
                                 preflight.FIX_OK, preflight.FIX_INHERITED)
            assert it["rule"], it
            assert resolvable(rules, it["rule"]), (it["property"], it["rule"])
        for row in res["not_handled"]:
            assert row["rule"] and resolvable(rules, row["rule"]), row
        for key in ("auto", "manual", "ok", "unchecked", "inherited"):
            assert key in res["stats"]


def test_html_summary_cards(wrong_docx):
    res = preflight.analyze(str(wrong_docx), load_rules())
    page = preflight.render_html(res)
    for label in ("自动修正", "需你处理", "已合规", "未检查"):
        assert label in page
    assert page.startswith("<!DOCTYPE html>")
    # HTML 里出现了规则引用与偏差属性，说明明细真的渲染进去了
    assert "三线表边框" in page
    assert "tables.border_pt" in page


def test_cli_writes_json_and_html(wrong_docx, tmp_path):
    out = tmp_path / "preflight.json"
    page = tmp_path / "preflight.html"
    r = subprocess.run(
        [sys.executable, str(ROOT / "tools" / "preflight.py"), str(wrong_docx),
         str(RULES / "tjyj.json"), "-o", str(out), "--html", str(page)],
        capture_output=True, text=True, cwd=str(ROOT))
    assert r.returncode == 0, r.stderr
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["journal"] == "统计研究"
    assert data["stats"]["auto"] > 0
    assert page.is_file() and "自动修正" in page.read_text(encoding="utf-8")


def test_heading4_within_max_level_is_checked(tmp_path, monkeypatch):
    """max_level 4 的刊：四级标题按 level4 检查，不再一律跳过。"""
    path = tmp_path / "h4.docx"
    doc = Document()
    doc.add_paragraph("(1)四级标题")
    doc.save(str(path))

    # 直接给出 IR，避免依赖解析器对四级编号的推断：这里要钉的是体检分支。
    monkeypatch.setattr(preflight.parse_docx, "parse", lambda *a, **k: {
        "blocks": [{"idx": 0, "type": "heading4", "level": 4,
                    "text": "(1)四级标题"}]})
    res = preflight.analyze(str(path), load_rules("cjyj"))

    assert not [u for u in res["unchecked"] if u["block"] == "heading4"]
    assert not [it for it in res["items"]
                if it["block"] == "heading4" and it["property"] == "标题层级"]
    # 编号前缀按 level4 的模板 (1) 判
    prefix = [it for it in res["items"]
              if it["block"] == "heading4" and it["property"] == "编号前缀"]
    assert prefix and prefix[0]["fix"] == "ok"


def test_heading4_beyond_max_level_is_manual_not_silent(tmp_path, monkeypatch):
    """max_level 3 的刊：四级标题按正文套版，体检出一条 manual 说明。"""
    path = tmp_path / "h4.docx"
    doc = Document()
    doc.add_paragraph("（1）四级标题")
    doc.save(str(path))

    monkeypatch.setattr(preflight.parse_docx, "parse", lambda *a, **k: {
        "blocks": [{"idx": 0, "type": "heading4", "level": 4,
                    "text": "（1）四级标题"}]})
    res = preflight.analyze(str(path), load_rules("tjyj"))

    manual = [it for it in res["items"]
              if it["block"] == "heading4" and it["property"] == "标题层级"]
    assert manual and manual[0]["fix"] == "manual"
    assert "只排到 3 级" in manual[0]["expected"]


def test_left_indent_without_first_line_is_not_reported_ok(tmp_path, monkeypatch):
    """只有左缩进的段落：渲染器会清掉整个 w:ind，体检必须报 auto，不能 ok。"""
    path = tmp_path / "left-indent.docx"
    doc = Document()
    p = doc.add_paragraph("一、引言")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "420")
    ind.set(qn("w:leftChars"), "200")
    p._p.get_or_add_pPr().append(ind)
    doc.save(str(path))

    monkeypatch.setattr(preflight.parse_docx, "parse", lambda *a, **k: {
        "blocks": [{"idx": 0, "type": "heading1", "level": 1,
                    "text": "一、引言"}]})
    res = preflight.analyze(str(path), load_rules())
    item = next(it for it in res["items"]
                if it["block"] == "heading1" and it["property"] == "缩进")
    assert item["fix"] == "auto", item
    assert item["expected"] == "无缩进"
    assert "左缩进" in item["actual"], item


def test_first_line_plus_left_indent_is_not_ok(tmp_path, monkeypatch):
    """首行缩进对上了、但还带左缩进：渲染器会清掉左缩进，体检不能报 ok。"""
    path = tmp_path / "first-left.docx"
    doc = Document()
    p = doc.add_paragraph("本文使用上市公司数据，考察数字金融对企业融资约束的影响。")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLineChars"), "200")
    ind.set(qn("w:leftChars"), "200")
    ind.set(qn("w:left"), "420")
    p._p.get_or_add_pPr().append(ind)
    doc.save(str(path))

    monkeypatch.setattr(preflight.parse_docx, "parse", lambda *a, **k: {
        "blocks": [{"idx": 0, "type": "body", "text": p.text}]})
    res = preflight.analyze(str(path), load_rules())
    item = next(it for it in res["items"]
                if it["block"] == "body" and it["property"] == "缩进")
    assert item["fix"] == "auto", item
    assert item["expected"] == "首行缩进 2 字符"
    assert "左缩进" in item["actual"], item


def test_zero_first_line_with_left_indent_is_not_ok(tmp_path, monkeypatch):
    """firstLineChars=0 但还有左缩进：目标「无缩进」不能因为 0 就判 ok。"""
    path = tmp_path / "zero-first-left.docx"
    doc = Document()
    p = doc.add_paragraph("一、引言")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:leftChars"), "200")
    ind.set(qn("w:left"), "420")
    p._p.get_or_add_pPr().append(ind)
    doc.save(str(path))

    monkeypatch.setattr(preflight.parse_docx, "parse", lambda *a, **k: {
        "blocks": [{"idx": 0, "type": "heading1", "level": 1,
                    "text": "一、引言"}]})
    res = preflight.analyze(str(path), load_rules())
    item = next(it for it in res["items"]
                if it["block"] == "heading1" and it["property"] == "缩进")
    assert item["fix"] == "auto", item
    assert "左缩进" in item["actual"], item


def test_title_line_spacing_is_reported_as_not_handled():
    """规则集声明了 title.line_spacing，渲染器按代码约定折算——必须在「不处理」里。"""
    rules = load_rules()
    rules["typography"]["title"]["line_spacing"] = {"mode": "multiple", "value": 1.2}
    rows = preflight._not_handled(rules)
    hit = [r for r in rows if r["property"] == "主标题行距"]
    assert hit and hit[0]["rule"] == "typography.title.line_spacing", rows
