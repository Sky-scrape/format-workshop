"""审读修复回归钉：判级与空白分隔一级标题的两处误判（b7 审读 ①②）。

- 判级①  detect_arabic_level 的预扫用 PAT["h4"] 直接匹配原文、不带标题护栏，
  正文枚举句「（1）样本选择说明。」就足以置 has_arabic3，让全文没有「一、」
  的阿拉伯层级稿（1／1.1）把「1.」判成二级；渲染到 cwe 再解析得
  「0.1 1. Introduction」——正是 IDMPOT 自述修掉的「编号越跑越多」。
  合成稿双渲染矩阵测不出它：首次渲染已把「1. 」改写成「1.1 」，二次 parse
  走 h2_num 稳定复现（幂等但首遍已错）。→ 「1.」判二级只认「一、」在场，
  「（1）」单独在场不再推翻判一级（预扫分不出「（1）标题」与「（1）正文句」）。
- 判级②  h1_ws 只有「短＋无句读」护栏，挡不住它自己注释点名的「一 些研究
  表明」——数量词被空格拆开的正文短句既短又无句读，且命中即 high 置信，
  AI 不得纠正（DESIGN §14.2）。→ 补与 h1_num 同口径的分类词负向断言（+些）。

夹具姿势：单元断言直接 import parse_docx（同 test_fix_b7_render.py）；
端到端走子进程 parse / render_from_ir（同文件 parse_ir 姿势）。
"""

import json
import re
import subprocess
import sys

import pytest
from docx import Document

from conftest import ROOT, RULES, TOOLS

sys.path.insert(0, str(TOOLS))
import parse_docx as pd  # noqa: E402


def _para(text):
    p = pd.ET.Element(pd.W + "p")
    r = pd.ET.SubElement(p, pd.W + "r")
    t = pd.ET.SubElement(r, pd.W + "t")
    t.text = text
    return (0, p, None)


def _parse(docx_path, out_path):
    r = subprocess.run(
        [sys.executable, str(TOOLS / "parse_docx.py"), str(docx_path),
         "-o", str(out_path), "--no-ai"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(ROOT))
    assert r.returncode == 0, r.stderr[-400:]
    return json.loads(out_path.read_text(encoding="utf-8"))


def _enum_docx(tmp_path, name="manuscript.docx"):
    """阿拉伯层级体例合成稿 + 一句以「（1）」开头的正文枚举。"""
    doc = Document()
    doc.add_paragraph("1. Introduction")
    doc.add_paragraph("1.1 Data")
    doc.add_paragraph("（1）样本选择说明。")
    doc.add_paragraph("正文段落说明数据处理方式。")
    src = tmp_path / name
    doc.save(str(src))
    return src


# ---------------------------------------------------------------- 判级①

def test_body_paren_enum_does_not_make_1dot_level2():
    """正文枚举句「（1）…」在场但无「一、」：「1.」仍判一级（审读根因单元钉）。"""
    doc = [
        _para("1. Introduction"),
        _para("1.1 Data"),
        _para("（1）样本选择说明。"),        # 正文枚举句，与「（1）标题」同形
        _para("（2）稳健性检验说明。"),
    ]
    assert pd.detect_arabic_level(doc) == (1, True, 4)


def test_paren_still_follows_1dot_when_cn1_present():
    """对照：一、／1.／（1）体系铁证不变——1. 二级、（1）顺延三级。"""
    assert pd.detect_arabic_level(
        [_para("一、引言"), _para("1. 理论基础"), _para("（1）货币政策传导机制")]) \
        == (2, True, 3)


def test_cn2_and_pure_arabic_cases_unchanged():
    """（一）在场判三级、纯阿拉伯层级判一级：既有口径不动（防修复回退）。"""
    assert pd.detect_arabic_level([_para("（一）理论基础"), _para("1. 小节")]) \
        == (3, True, 4)
    assert pd.detect_arabic_level([_para("1. 引言"), _para("2.1 文献综述")]) \
        == (1, True, 4)


def test_first_parse_keeps_1dot_level1_despite_body_enum(tmp_path):
    """端到端：正文（1）枚举句在场，首次 parse「1. Introduction」就是一级。

    审读实测的病灶在**首遍**：首次渲染已把「1. 」改写成「1.1 」，二次 parse
    走 h2_num 稳定复现——所以只钉首遍的判级结果。
    """
    ir = _parse(_enum_docx(tmp_path), tmp_path / "m.ir.json")
    by_text = {b["text"]: b["type"] for b in ir["blocks"]}
    assert by_text["1. Introduction"] == "heading1", by_text
    assert by_text["1.1 Data"] == "heading2", by_text
    # 枚举句在文首前三段里会按位置推断成 front_matter（既有行为）；要钉的
    # 是它不得成为标题、更不得把「1.」顶成二级
    assert not by_text["（1）样本选择说明。"].startswith("heading"), by_text


def test_cwe_roundtrip_no_runaway_numbering(tmp_path):
    """审读复现链全程：parse → 渲染到 cwe → 再 parse，不得出现「0.1 …」。"""
    rules = RULES / "cwe.json"
    if not rules.exists():
        pytest.skip("cwe 不在本机规则库")
    src = _enum_docx(tmp_path)
    # IR 与源稿同名配套：render_from_ir 按 IR 路径找同名 .docx
    ir1_path = tmp_path / "manuscript.ir.json"
    _parse(src, ir1_path)
    out = tmp_path / "m.cwe.docx"
    r = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir1_path),
         str(rules), "-o", str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(ROOT))
    assert r.returncode == 0, r.stderr[-500:]
    ir2 = _parse(out, tmp_path / "m2.ir.json")
    texts = [b["text"] for b in ir2["blocks"] if b["type"].startswith("heading")]
    assert "1. Introduction" in texts, f"一级标题读不回：{texts}"
    assert not [t for t in texts if re.match(r"0\.\d", t)], \
        f"编号越跑越多（「0.1 …」）：{texts}"


# ---------------------------------------------------------------- 判级②

def test_h1_ws_rejects_split_measure_word_sentences():
    """数量词被空格拆开的正文短句不得冒充一级标题（审读探针原样）。"""
    for t in ("一 些研究表明", "一 种新的研究方法被提出", "一 个例子",
              "三 种类型的检验", "一 条经验证据"):
        assert pd.classify_heading(t) is None, t


def test_h1_ws_still_accepts_real_ws_headings():
    """真「一　标题」不受分类词排除误伤；顿号体例与长句护栏维持原口径。"""
    assert pd.classify_heading("一　引  言") == (1, "high")
    assert pd.classify_heading("十二　文献综述与理论框架") == (1, "high")
    assert pd.classify_heading("一、引言") == (1, "high")
    assert pd.classify_heading(
        "一 些研究表明上述结论在更换度量方式后依然成立，样本为A股。") is None
