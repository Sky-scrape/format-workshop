"""合并单元格的三线表：栏目线必须落在「表头区最后一行」下面。

经管论文的表头经常跨两行——「变量」跨行、分组标题跨列。旧实现按 `ri == 0`
给每个单元格设下边框，结果是一张表里三种情况同时出现：

- 跨行单元格（vMerge restart）的底边落在合并区底部 → 线在第 1 行下（碰巧对）
- 跨列分组标题（gridSpan）在第 0 行 → 线画到了表头中间（错）
- 真正的列名在第 1 行，没人给它画线（错）

看起来就是一条断成几截、位置还不对的线。这不是排版瑕疵，是三线表画错了。
"""

import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import render_from_ir as R  # noqa: E402

BP = {"top": 1.5, "bottom": 1.5, "header_bottom": 0.75}


def build_merged_table(doc):
    """两行表头 + 两行表体：

        | 变量   | 主效应        | 控制变量      |   ← 第 0 行：变量跨 2 行，
        |        | 系数  | t 值   | 系数  | t 值  |     两个分组各跨 2 列
        | 规模   | 0.32  | 3.10   | 0.11  | 1.02  |
        | 杠杆   | -0.21 | -2.33  | 0.07  | 0.88  |
    """
    t = doc.add_table(rows=4, cols=5)
    t.cell(0, 0).merge(t.cell(1, 0)).text = "变量"
    t.cell(0, 1).merge(t.cell(0, 2)).text = "主效应"
    t.cell(0, 3).merge(t.cell(0, 4)).text = "控制变量"
    for col, text in ((1, "系数"), (2, "t 值"), (3, "系数"), (4, "t 值")):
        t.cell(1, col).text = text
    for row, values in ((2, ("规模", "0.32", "3.10", "0.11", "1.02")),
                        (3, ("杠杆", "-0.21", "-2.33", "0.07", "0.88"))):
        for col, text in enumerate(values):
            t.cell(row, col).text = text
    return t


def _tc_text(tc):
    return "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()


def _has_bottom(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return False
    tcb = tcPr.find(qn("w:tcBorders"))
    if tcb is None:
        return False
    b = tcb.find(qn("w:bottom"))
    return b is not None and b.get(qn("w:val")) == "single"


def _rows(table):
    return [tr.findall(qn("w:tc")) for tr in table._tbl.findall(qn("w:tr"))]


# ---------------------------------------------------------------- 辅助函数

def test_grid_span_and_vmerge_are_read_back():
    """铺格矩阵要能正确认出合并：跨列占多格、跨行续行回填成起始 tc。"""
    doc = Document()
    table = build_merged_table(doc)

    grid = R._table_grid(table._tbl)
    assert len(grid) == 4 and all(len(row) == 5 for row in grid)
    # 第 0 列两行是同一个逻辑单元格
    assert grid[0][0] is grid[1][0]
    # 跨列分组：第 1、2 列同一个 tc；第 3、4 列同一个 tc
    assert grid[0][1] is grid[0][2]
    assert grid[0][3] is grid[0][4]
    # 表头两行（第 0 行有跨到第 1 行的合并）
    assert R._header_rows(grid) == 2


# ---------------------------------------------------------------- 栏目线

def test_header_line_lands_below_the_last_header_row():
    """栏目线画在表头区最后一行下面，不是无脑画在第 0 行下面。"""
    doc = Document()
    table = build_merged_table(doc)
    R.apply_table(table, BP, 9, "宋体", "Times New Roman")

    rows = _rows(table)
    # 第 1 行是列名行。前 4 个逻辑单元格该有线；第 0 列是「变量」的续行空壳，
    # 边框必须挂在起始 tc 上——续行空壳的边框 Word 会直接忽略，挂了也没用
    for tc in rows[1]:
        if R._vmerge_state(tc) == "continue":
            assert not _has_bottom(tc), "续行空壳不该挂边框（Word 会忽略它）"
        else:
            assert _has_bottom(tc), f"列名单元格缺栏目线：{_tc_text(tc)}"
    # 跨列分组标题在第 0 行，底边也在第 0 行 —— 不该有线（旧实现正是在这里画错）
    assert not any(_has_bottom(tc) for tc in rows[0][1:]), \
        [(_tc_text(tc), _has_bottom(tc)) for tc in rows[0][1:]]
    # 跨行单元格「变量」底边在合并区底部（第 1 行），该有线
    assert _has_bottom(rows[0][0]), "跨行表头单元格的底边在合并区底部，必须有栏目线"


def test_body_cells_get_no_cell_borders():
    """表体单元格不设 cell 级边框：三线表只有表头下那一条，其余靠表级边框。"""
    doc = Document()
    table = build_merged_table(doc)
    R.apply_table(table, BP, 9, "宋体", "Times New Roman")

    for r in (2, 3):
        for tc in _rows(table)[r]:
            assert not _has_bottom(tc), f"第 {r} 行不该有 cell 级下边框：{_tc_text(tc)}"


def test_plain_table_keeps_old_behaviour():
    """没有合并的普通表格：第 0 行有栏目线，行为与改动前一致。"""
    doc = Document()
    t = doc.add_table(rows=3, cols=3)
    for r in range(3):
        for c in range(3):
            t.cell(r, c).text = f"r{r}c{c}"
    R.apply_table(t, BP, 9, "宋体", "Times New Roman")

    rows = _rows(t)
    assert all(_has_bottom(tc) for tc in rows[0])
    assert not any(_has_bottom(tc) for tc in rows[1] + rows[2])


def test_single_row_table_still_gets_a_line():
    """单行表格：表头即表体，仍然要有栏目线（不能因为「留一行表体」的保护丢掉）。"""
    doc = Document()
    t = doc.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "a"
    t.cell(0, 1).text = "b"
    R.apply_table(t, BP, 9, "宋体", "Times New Roman")

    assert all(_has_bottom(tc) for tc in _rows(t)[0])


# ---------------------------------------------------------------- 幂等

def test_snapshot_ignores_untouched_continuation_cells():
    """快照与 apply_table 同口径：续行空壳不参与比对。

    两边口径不一致时，apply_table 没碰过的续行空壳会进快照，
    「完全合规的表」每跑一次都被判成改动。
    """
    doc = Document()
    table = build_merged_table(doc)
    R.apply_table(table, BP, 9, "宋体", "Times New Roman")

    first = R.table_snapshot(table)
    R.apply_table(table, BP, 9, "宋体", "Times New Roman")
    second = R.table_snapshot(table)
    assert first == second, "同一张表连续套两次版，快照应当完全相同"


# ---------------------------------------------------------------- 完整流程

def test_merged_table_runs_through_pipeline(tmp_path):
    """真走一遍流水线：带合并表头的表格能被识别、套成三线表，重跑零改动。

    单测只证明了 apply_table 的手艺，这条证明它在真实流水线里也活着——
    识别、审计、幂等三关都过。
    """
    from conftest import minimal_doc, render_into

    src = tmp_path / "with-table.docx"
    doc = minimal_doc(src, title="企业创新的驱动因素研究",
                      abstract="内容提要：本文考察企业创新的驱动因素。",
                      keywords="关键词：企业创新；研发投入")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("表 1 报告了主要变量的描述性统计。")
    doc.add_paragraph("表1  主要变量描述性统计")
    build_merged_table(doc)
    doc.save(str(src))

    first = render_into(tmp_path / "run1", src, target="tjyj")
    tables = [b for b in first.ir["blocks"] if b["type"] == "table"]
    assert tables, "流水线没把这份表识别成 table"

    acts = [it["action"] for it in first.audit["items"]]
    assert "表格重设为三线表" in acts, f"表格没进审计：{acts}"

    second = render_into(tmp_path / "run2", first.docx, target="tjyj")
    leftover = [it for it in second.audit["items"] if it.get("changed", True)]
    names = "、".join(f"{it['action']}@{it['idx']}" for it in leftover[:5])
    assert leftover == [], f"带合并表的成品重跑仍报 {len(leftover)} 条改动：{names}"
