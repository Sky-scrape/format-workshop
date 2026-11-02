"""审读第二轮五个缺口的回归钉。

1. 「可驳回」声明与实现矛盾：N-07（文献条目编号归一）与关键词分隔符归一
   同为 B 档、渲染器都有 audit.blocked 分支，却不在 render_audit 的
   REJECTABLE 集合——审计页不渲染勾选框，DESIGN §4.1「全部逐条进审计、
   可驳回（B 档）」与 README 的明文声明落空。
2. table_snapshot 只看边框 val/sz 与 run 的 eastAsia/ascii，而 apply_table
   无条件重写六边 w:color→000000、w:space→0 并重写 run 的 w:hAnsi——
   这些真变更不产生审计条目，违反模块红线「全部进审计报告」。
3. 「统计范围」格对没有声明 limits.total.value 的刊照抄默认文案
   「全部正文段落（不含表格内文字）」，而 A/B/C 数字实为全文口径——
   标注与计算自相矛盾（F18 要消除的正是这个）。
4. 「全文（含图表）」口径实际数不出表格内文字：count_words 只取
   b.get("text")，而 parse_docx 产出的 table 块不带文字。
5. 「页面设置」条目（idx=-1）无条件输出：第二遍渲染版式全合规时，
   审计仍有一条无差异记录。
"""

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from conftest import RULES, ROOT, TOOLS, render_into  # noqa: E402

import compliance_report as crp  # noqa: E402
from lxml import etree  # noqa: E402
from render_audit import render as render_audit_html  # noqa: E402


def _rules(name):
    """直接加载规则库里的刊（本机 23 本全有；缺失时跳过而不是误报）。"""
    p = RULES / f"{name}.json"
    if not p.exists():
        pytest.skip(f"{name}.json 不在本机规则库")
    return json.loads(p.read_text(encoding="utf-8"))


def _audit_item(audit, action, idx=None):
    return next((it for it in audit["items"]
                 if it["action"] == action and (idx is None or it["idx"] == idx)), None)


def _scope_row(html_out):
    m = re.search(r"统计范围</td><td>(.*?)</td>", html_out, re.S)
    assert m, "报告里没有「统计范围」这一行"
    return m.group(1)


# ---------- 1. B 档动作的驳回勾选框 ----------

def _one_item_html(action):
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        {"id": f"6|{action}", "idx": 6, "sec": "", "block": "body", "action": action,
         "before": "前", "after": "后", "note": "", "rule": "r", "rejected": False}]}
    return render_audit_html(audit, job="t1")


@pytest.mark.parametrize("action", ["文献条目编号归一", "关键词分隔符归一"])
def test_b_grade_actions_render_reject_checkbox(action):
    """B 档动作全部渲染勾选框——声明「逐条可驳回」的动作不能在 UI 上失踪。"""
    html_out = _one_item_html(action)
    assert f'value="6|{action}"' in html_out, f"{action} 没有驳回勾选框"
    assert "保留原文" in html_out


def test_a_grade_actions_stay_unrejectable():
    """A 档是纯版式，驳回后成品就不合规，仍不给勾选框。"""
    assert 'value="1|正文套版"' not in _one_item_html("正文套版")


def test_keyword_separator_reject_round_trips(baseline, sample_docx, tmp_path):
    """勾选框的 id 提交回渲染器要真的生效：保留原文、条目记「已驳回」。"""
    target = next(it for it in baseline.items if it["action"] == "关键词分隔符归一")
    after = render_into(tmp_path / "rej", sample_docx, rejects=[target["id"]])
    item = after.item_by_id(target["id"])
    assert item is not None and item["rejected"] is True, "驳回没被记录进审计"
    assert after.output_texts[target["idx"]] == baseline.source_texts[target["idx"]], \
        "驳回后关键词分隔符仍被改了"


# ---------- 2. apply_table 的静默改动必须进审计 ----------

def _style_table_runs(doc, hansi="Times New Roman"):
    """表内 run 按 tjyj 目标预置（宋体/Times New Roman/7.5pt），只留指定缺口。"""
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(7.5)
                        rpr = run._element.get_or_add_rPr()
                        rf = rpr.find(qn("w:rFonts"))
                        if rf is None:
                            rf = OxmlElement("w:rFonts")
                            rpr.append(rf)
                        rf.set(qn("w:eastAsia"), "宋体")
                        rf.set(qn("w:ascii"), "Times New Roman")
                        rf.set(qn("w:hAnsi"), hansi)


def _build_table_docx(path, border_color, border_space, hansi="Times New Roman"):
    """三行两列的表：边框 val/sz 与 tjyj 目标一致（顶底 1.5pt、栏目线 0.5pt），
    颜色与间距由参数控制——构造「形状合规、上色不同」的审读复现输入。"""
    doc = Document()
    t = doc.add_table(rows=3, cols=2)
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            cell.text = f"数据{ri}{ci}abcdefghijkl"
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")          # 1.5pt × 8
        else:
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), str(border_space))
        el.set(qn("w:color"), border_color)
        borders.append(el)
    tblPr.append(borders)
    for cell in t.rows[0].cells:              # 首行栏目线
        tcPr = cell._tc.get_or_add_tcPr()
        tcb = OxmlElement("w:tcBorders")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")                # 0.5pt × 8
        b.set(qn("w:space"), str(border_space))
        b.set(qn("w:color"), border_color)
        tcb.append(b)
        tcPr.append(tcb)
    _style_table_runs(doc, hansi=hansi)
    doc.add_paragraph("此段用于核验正文版式。")
    doc.save(str(path))


def _parse_and_render(docx, rules, tag):
    """与 test_citation_convert.run_render 同款：子进程跑解析＋渲染。"""
    ir = docx.with_suffix(".ir.json")
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(docx), "-o", str(ir)],
                   check=True, capture_output=True, text=True, cwd=str(ROOT))
    out = docx.with_name(f"{tag}.docx")
    r = subprocess.run([sys.executable, str(TOOLS / "render_from_ir.py"),
                        str(ir), str(rules), "-o", str(out)],
                       capture_output=True, text=True, cwd=str(ROOT))
    assert r.returncode == 0, r.stderr[-800:]
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    return out, audit


def _tbl_edge_paint(docx_path):
    tbl = Document(str(docx_path)).tables[0]._tbl
    el = tbl.tblPr.find(qn("w:tblBorders"))
    return {etree.QName(c).localname: (c.get(qn("w:color")), c.get(qn("w:space")))
            for c in el}


def test_color_only_table_change_is_audited(tmp_path):
    """边框形状合规、仅颜色不同：修前 0 条审计（静默改动），修后必须有一条。"""
    src = tmp_path / "red.docx"
    _build_table_docx(src, border_color="FF0000", border_space="4")
    out, audit = _parse_and_render(src, RULES / "tjyj.json", "red_out")

    item = _audit_item(audit, "表格重设为三线表")
    assert item is not None, "颜色重写是真实改动，不能静默"
    assert "颜色" in item["note"] or "重写" in item["note"], item["note"]
    # 成品确实归一了：六边黑色、间距 0
    paint = _tbl_edge_paint(out)
    assert set(paint.values()) == {("000000", "0")}, paint

    # 第二遍：表已完全合规，快照相同，不再出条目（F28 在更全的快照口径下仍成立）
    out2, audit2 = _parse_and_render(out, RULES / "tjyj.json", "red_out2")
    assert _audit_item(audit2, "表格重设为三线表") is None, "完全合规的表不该再出条目"


def test_hansi_only_change_is_audited(tmp_path):
    """run 的 ascii 已合规、仅 hAnsi 不同：set_run 会重写它，快照必须看得见。"""
    src = tmp_path / "hansi.docx"
    _build_table_docx(src, border_color="000000", border_space="0", hansi="Arial")
    out, audit = _parse_and_render(src, RULES / "tjyj.json", "hansi_out")

    item = _audit_item(audit, "表格重设为三线表")
    assert item is not None, "hAnsi 重写是真实改动，不能静默"
    assert "hAnsi" in item["note"], item["note"]
    rf = Document(str(out)).tables[0].rows[0].cells[0].paragraphs[0].runs[0] \
        ._element.rPr.find(qn("w:rFonts"))
    assert rf.get(qn("w:hAnsi")) == "Times New Roman"


# ---------- 3/4. 统计范围标注与计算一致 ----------

def _scope_ir():
    """500 摘要 + 5000 正文 + 500 文献（审读用的复现 IR，B 口径全文 6000）。"""
    return {"source": "repro.docx", "blocks": [
        {"type": "abstract", "text": "摘" * 500},
        {"type": "body", "text": "正" * 5000},
        {"type": "reference_item", "text": "文" * 500},
    ]}


@pytest.mark.parametrize("name", ["cdxb", "cer"])
def test_scope_cell_without_total_value_says_full_scope(name):
    """没声明 total.value 的刊：数字是全文口径，标注不得再写「全部正文段落」。"""
    rules = _rules(name)
    res = crp.check(_scope_ir(), rules)
    assert res["total_limit"] is None, "用例前提：该刊没有声明的总字数上限"
    cell = _scope_row(crp.render(_scope_ir(), rules, res))
    assert "全部正文段落" not in cell, f"{name} 的标注仍与全文口径的计算相反"
    assert "按全文口径统计" in cell or cell == "全文", cell


def test_scope_cell_without_total_value_but_declared_scope():
    """声明了 scope 但没有上限值：照抄声明，且标注与实际统计的块一致。"""
    rules = _rules("tjyj")
    rules["limits"]["total"]["value"] = None
    res = crp.check(_scope_ir(), rules)
    assert res["words_scope"] == "body" and res["words_scope_declared"] == "正文"
    assert _scope_row(crp.render(_scope_ir(), rules, res)) == "正文"


def test_count_words_all_scope_includes_table_cells():
    """「全文（含图表）」口径必须数表格内文字；figure 无文本层文字，天然不计。"""
    blocks = [
        {"type": "body", "text": "正" * 100},                      # B: 100
        {"type": "table", "idx": 3, "n_rows": 2, "n_cols": 2,
         "cells": ["甲" * 3, "乙" * 3, "丙" * 3, "丁" * 3]},        # B: +12
        {"type": "figure", "text": "", "idx": 4},                  # 不计
    ]
    assert crp.count_words(blocks, scope="all")["B"] == 112
    assert crp.count_words(blocks, scope="body")["B"] == 100, "正文口径不受影响"


def test_parse_produces_table_cells(tmp_path):
    """parse_docx 的 table 块要带 cells，全文口径的统计才有表内文字可数。"""
    doc = Document()
    doc.add_paragraph("此段用于核验正文版式。")
    t = doc.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text = "变量"
    t.rows[0].cells[1].text = "系数"
    t.rows[1].cells[0].text = "规模"
    t.rows[1].cells[1].text = "0.37"
    doc.save(str(tmp_path / "t.docx"))

    ir_path = tmp_path / "t.ir.json"
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"),
                    str(tmp_path / "t.docx"), "-o", str(ir_path)],
                   check=True, capture_output=True, text=True, cwd=str(ROOT))
    ir = json.loads(ir_path.read_text(encoding="utf-8"))
    tb = next(b for b in ir["blocks"] if b["type"] == "table")
    assert tb["cells"] == ["变量", "系数", "规模", "0.37"], tb.get("cells")

    # 全文口径把表内文字计入（jmsc 是唯一声明「全文（含图表）」的入库刊）
    rules = _rules("jmsc")
    res = crp.check(ir, rules)
    assert res["words"][1][1].startswith(f"{crp.count_words(ir['blocks'])['B']:,}")
    assert _scope_row(crp.render(ir, rules, res)) == "全文（含图表）"


# ---------- 5. 页面设置条目只在真有差异时输出 ----------

def test_page_setup_entry_only_when_changed(sample_docx, tmp_path):
    """gjjryj→tjyj 页边距真有差异，第一遍记一条；第二遍版式已合规，不再有。"""
    first = render_into(tmp_path / "p1", sample_docx)
    assert _audit_item({"items": first.items}, "页面设置", idx=-1) is not None, \
        "第一遍真有差异，「页面设置」必须进审计"

    second = render_into(tmp_path / "p2", first.dir / "manuscript.tjyj.docx")
    assert [it for it in second.items if it["idx"] == -1] == [], \
        "第二遍版式全合规，不该再有无差异条目"
    assert second.output_texts == first.output_texts, "去条目不能改变成品"
