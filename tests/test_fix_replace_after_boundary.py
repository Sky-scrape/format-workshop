"""F02 返工回归钉：replace_after 在「保留区恰好落在物理单元边界」时不得丢字。

上一轮修复 F02 引入的 replace_after 只在「跨界节点」里写新文字：当 keep_len
恰好等于保留区若干完整单元的长度（标签独占一个加粗 run、或标签归一刚把
「关键词\\t：」合并成一个节点）时，没有任何节点进入写分支，整段关键词正文
被静默清空，审计 after 还记着替换前的期望值——比原 F02 更重（整段 vs 一字）。

钉住的契约：
1. 边界对齐时新文字写进替换区第一个文字节点（继承正文格式，不吃标签加粗）
2. 保留区里的 tab/br 原样保留；替换区里的 tab/br 随正文一起删除
3. 替换区没有文字节点时兜底 add_run
4. 端到端：成品文字正确，审计 after 与成品逐字一致
"""

import importlib.util
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"

from conftest import render_into  # noqa: E402


def _load(name, filename):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


render = _load("render_from_ir_f02", "render_from_ir.py")
qn = render.qn


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


def agg_text(p):
    return render.para_full_text(p)


def runs_with_text(p):
    return ["".join(t.text or "" for t in r.findall(qn("w:t")))
            for r in p._p.xpath(".//w:r")]


# ---------- 单元级：边界布局不丢字 ----------

def test_bold_label_own_run_body_kept():
    """标签独占加粗 run（最常见边界布局）：正文 9 字完好，加粗不扩散。"""
    doc = Document()
    p = doc.add_paragraph()
    label = p.add_run("关键词：")
    label.bold = True
    body = p.add_run("数字金融，实体经济")
    render.replace_after(p, 4, "数字金融；实体经济")
    assert agg_text(p) == "关键词：数字金融；实体经济"
    assert body.text == "数字金融；实体经济"          # 新文字落在正文自己的 run
    assert label.text == "关键词：" and label.bold is True
    plain = [r for r in p._p.xpath(".//w:r")
             if r.find(qn("w:rPr")) is None
             or r.find(qn("w:rPr")).find(qn("w:b")) is None]
    assert all("".join(t.text or "" for t in r.findall(qn("w:t"))) == "数字金融；实体经济"
               for r in plain), runs_with_text(p)


def test_chain_prefix_then_after_keeps_everything():
    """标签归一→分隔符归一两连跳：首节点恰被前缀填满时，第二步不得吞正文。"""
    p = new_para("关键词", "<TAB>", "：数字金融，实体经济")
    old = agg_text(p)
    m = re.match(r"^(关\s*键\s*词)\s*[:：]", old)
    render.replace_prefix(p, m.end(), "关键词：")
    assert agg_text(p) == "关键词：数字金融，实体经济"
    before = agg_text(p)
    sep = "；"
    items = render.split_keywords(before[len("关键词："):], sep)
    render.replace_after(p, len("关键词："), sep.join(items))
    assert agg_text(p) == "关键词：数字金融；实体经济"


def test_boundary_at_tab_end_keeps_tab():
    """keep_len 落在 tab 末尾（保留区含 tab）：tab 留着，正文整体替换。"""
    p = new_para("关键词", "<TAB>", "数字金融，实体经济")
    render.replace_after(p, 4, "数字金融；实体经济")
    assert agg_text(p) == "关键词\t数字金融；实体经济"
    assert p._p.xpath(".//w:r/w:tab"), "保留区里的 tab 被误删"


def test_replacement_zone_starting_with_tab_drops_tab():
    """替换区以 tab 开头：tab 属于被替换内容，删除后新文字落在文字节点。"""
    p = new_para("关键词：", "<TAB>", "数字金融，实体经济")
    render.replace_after(p, 4, "数字金融；实体经济")
    assert agg_text(p) == "关键词：数字金融；实体经济"
    assert not p._p.xpath(".//w:r/w:tab")


def test_replacement_zone_starting_with_break_drops_break():
    p = new_para("关键词：", "<BR>", "数字金融，实体经济")
    render.replace_after(p, 4, "数字金融；实体经济")
    assert agg_text(p) == "关键词：数字金融；实体经济"


def test_cross_boundary_run_regression():
    """跨界节点（keep_len 落在 w:t 内部）：保留半截标签、接新文字——既有行为。"""
    p = new_para("关键词：数字金融，实体经济")
    render.replace_after(p, 4, "数字金融；实体经济")
    assert agg_text(p) == "关键词：数字金融；实体经济"


def test_full_keep_zone_without_text_node_appends():
    """keep_len 吃满全段（替换区无文字节点）：兜底 add_run，文字不得凭空消失。"""
    p = new_para("关键词：")
    render.replace_after(p, 4, "X")
    assert agg_text(p) == "关键词：X"


# ---------- 端到端：真实管线 + 真实规则集（tjyj） ----------

def _kw_audit_item(result, idx):
    return result.item_by_id(f"{idx}|关键词分隔符归一")


def test_bold_label_run_end_to_end(tmp_path):
    """加粗标签独占 run 的关键词段：成品正文完好，审计 after 与成品逐字一致。

    修复前：成品只剩「关键词：」，9 字正文被清空，审计 after 仍记期望值。
    """
    src = tmp_path / "kwbold.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    p = doc.add_paragraph()
    label = p.add_run("关键词：")
    label.bold = True
    p.add_run("数字金融，实体经济")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    out = {i: t for i, t in result.output_texts.items()
           if result.block_type.get(i) == "keywords"}
    assert list(out.values()) == ["关键词：数字金融；实体经济"], out

    item = _kw_audit_item(result, list(out)[0])
    assert item and not item["rejected"]
    assert item["after"] == out[list(out)[0]], "审计 after 与成品不符"
    assert "数字金融" in item["before"] and "，" in item["before"]


def test_label_tab_then_separator_end_to_end(tmp_path):
    """标签+tab 先归一标签、再归一分隔符：两连跳后正文一字不吞。"""
    src = tmp_path / "kwtab.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    p = doc.add_paragraph()
    p.add_run("关键词")
    p.runs[0]._r.append(OxmlElement("w:tab"))
    p.add_run("：数字金融，实体经济")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    out = {i: t for i, t in result.output_texts.items()
           if result.block_type.get(i) == "keywords"}
    idx = list(out)[0]
    assert out[idx] == "关键词：数字金融；实体经济", out

    item = _kw_audit_item(result, idx)
    assert item and not item["rejected"]
    assert item["after"] == out[idx], "审计 after 与成品不符"

    lab = result.item_by_id(f"{idx}|标签归一")
    assert lab and lab["after"] == "关键词：数字金融，实体经济", \
        "标签归一这一步的成品快照不得预先包含第二步的改动"
