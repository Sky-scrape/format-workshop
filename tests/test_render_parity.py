"""样板与交付稿必须用同一套版式写入。

render_sample.py 与 render_from_ir.py 原先各有一份 set_run／缩进写入。复制品
的典型漂移是：样板按 11.8pt 截成 11.5pt，交付稿却四舍五入到 12pt；或者规则集
把字间距归零时，一边删旧 w:spacing、另一边留着。用户拿样板核对版式，交付稿
却对不上，样板就不成其为基准。

公共实现已收进 tools/docx_fmt.py（见那里的模块说明）。本文件钉两层：
1. 同一组字号/字间距，两个模块写出的 w:sz、w:rFonts、w:spacing 完全一致，
   且半磅取整、字间距清零（先删旧元素）的行为固定下来；
2. 同一规则集下，「render_sample 生成样板」与「样板→render_from_ir 转换」
   在正文/标题/题注的字体、字号、字间距、对齐、行距、缩进上一致。

第 2 层只比 run 与段落的关键版式属性，不比段前/段后距：样板为了在 Word 里
可读，标题上下留白由 add_para 的默认值决定，而交付稿按规则集另算，两者本来
就不是同一个概念。
"""

import re

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import docx_fmt
import render_from_ir
import render_sample
from conftest import RULES, run_tool, render_into

TITLE_TEXT = "美国宽松货币政策对金砖国家的溢出效应研究"
BODY_MARK = "此段用于核验正文版式"
CAPTION_RE = re.compile(r"^表\s*1\s*主要变量描述性统计$")


def _is_title(text):
    return text.strip() == TITLE_TEXT


def _is_body(text):
    return BODY_MARK in text


def _is_caption(text):
    return bool(CAPTION_RE.match(text.strip()))


# ---------------------------------------------------------------- 一、写入原语

def _seed_spacing(run, val="99"):
    """在 run 上先放一个旧字间距，模拟「从带字间距的模板切刊」。"""
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), val)
    run._element.get_or_add_rPr().append(sp)


def _written_run_state(run):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    sz = rpr.find(qn("w:sz"))
    spaces = rpr.findall(qn("w:spacing"))
    return {
        "eastAsia": rf.get(qn("w:eastAsia")) if rf is not None else None,
        "ascii": rf.get(qn("w:ascii")) if rf is not None else None,
        "hAnsi": rf.get(qn("w:hAnsi")) if rf is not None else None,
        "sz": sz.get(qn("w:val")) if sz is not None else None,
        "spacing": [s.get(qn("w:val")) for s in spaces],
    }


@pytest.mark.parametrize("size", [7.5, 9, 10.5, 11.8, 22])
@pytest.mark.parametrize("letter_spacing", [None, 0, 0.6, 1.3])
def test_set_run_parity_between_renderers(size, letter_spacing):
    """同一输入下两个渲染器写出的 run 属性逐项相同。

    旧实现的差异点：样板不取半磅（11.8 被 python-docx 截成 11.5），且字间距
    只在真值时才追加、从不删旧元素（声明 0 时残留旧值）。共用 docx_fmt.set_run
    后，两条路径必须给出同一份 XML。
    """
    states = []
    for mod in (render_from_ir, render_sample):
        doc = Document()
        run = doc.add_paragraph().add_run("核验")
        _seed_spacing(run)                      # 旧值 99，检验是否被正确覆盖/清除
        mod.set_run(run, cn="宋体", latin="Times New Roman", size=size,
                    letter_spacing=letter_spacing)
        states.append(_written_run_state(run))

    sample_state, output_state = states
    assert sample_state == output_state, (
        f"两个渲染器对 size={size} letter_spacing={letter_spacing} 写出的 run 属性不同："
        f"\nrender_sample {sample_state}\nrender_from_ir {output_state}")

    # 半磅刻度：w:sz 是整数 half-point，11.8pt 必须落在 12pt（24）而不是 11.5（23）
    assert sample_state["sz"] == str(int(round(size * 2)))
    assert sample_state["eastAsia"] == "宋体"
    assert sample_state["ascii"] == "Times New Roman"
    assert sample_state["hAnsi"] == "Times New Roman"

    if letter_spacing is None:
        # None 表示规则集没声明：不动 run 上已有的字间距
        assert sample_state["spacing"] == ["99"]
    elif letter_spacing == 0:
        # 声明 0 必须把旧 w:spacing 删掉，不能留下两个或残留旧值
        assert sample_state["spacing"] == []
    else:
        assert sample_state["spacing"] == [str(int(round(letter_spacing * 20)))]


def test_renderers_share_one_implementation():
    """两个模块必须 import 同一份实现，而不是各自留一份同名函数。"""
    assert render_from_ir.DEFAULTS is docx_fmt.DEFAULTS
    assert render_sample.DEFAULTS is docx_fmt.DEFAULTS
    assert render_from_ir.set_run is docx_fmt.set_run
    assert render_sample.set_run is docx_fmt.set_run
    assert render_from_ir.set_chars_indent is docx_fmt.set_chars_indent
    # render_from_ir 的历史公开名 set_hanging 是 docx_fmt.set_hanging_indent 的别名
    assert render_from_ir.set_hanging is docx_fmt.set_hanging_indent
    assert render_sample.set_hanging_indent is docx_fmt.set_hanging_indent
    assert render_sample.set_line_unit_space is docx_fmt.set_line_unit_space


# ---------------------------------------------------------------- 二、样板 ↔ 成品

def _paragraphs(path):
    doc = Document(str(path))
    return [Paragraph(el, doc) for el in doc.element.body if el.tag == qn("w:p")]


def _find_one(paras, pred, label, path):
    hits = [p for p in paras if pred(p.text)]
    assert len(hits) == 1, f"{path.name}: {label} 段落应恰有一个，实际 {len(hits)} 个"
    return hits[0]


def _style_east_asia(style):
    rpr = style.element.find(qn("w:rPr"))
    rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
    return rf.get(qn("w:eastAsia")) if rf is not None else None


def _inherited(p, getter):
    """沿段落样式的继承链取第一个非空值（样板没有显式写 ascii 时靠它继承）。"""
    style = p.style
    seen = set()
    while style is not None and id(style.element) not in seen:
        seen.add(id(style.element))
        value = getter(style)
        if value is not None:
            return value
        style = style.base_style
    return None


def _effective_run_state(p, run):
    rpr = run._element.rPr
    rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
    sp = rpr.find(qn("w:spacing")) if rpr is not None else None

    size = run.font.size or _inherited(p, lambda s: s.font.size)
    bold = run.font.bold
    if bold is None:
        bold = _inherited(p, lambda s: s.font.bold)

    ascii_font = (rf.get(qn("w:ascii")) if rf is not None else None) \
        or _inherited(p, lambda s: s.font.name)
    return {
        # 不比 run 文字：这条用例比的是版式。跨刊夹具的样板正文里写着「行距
        # 20 磅」这类自述，源稿不同时文字本来就该不同（文字零改动由其他用例守）。
        "eastAsia": (rf.get(qn("w:eastAsia")) if rf is not None else None)
                    or _inherited(p, _style_east_asia),
        # hAnsi 缺省时 Word 按 ascii 的西文字体排，等价比较要跟着这个口径
        "hAnsi": (rf.get(qn("w:hAnsi")) if rf is not None else None)
                 or (rf.get(qn("w:ascii")) if rf is not None else None)
                 or _inherited(p, lambda s: s.font.name),
        "ascii": ascii_font,
        "sz": round(size.pt, 2) if size is not None else None,
        "spacing": sp.get(qn("w:val")) if sp is not None else None,
        "bold": bool(bold) if bold is not None else False,
    }


def _line_sig(p):
    ls = p.paragraph_format.line_spacing
    if hasattr(ls, "pt"):
        return ("pt", round(ls.pt, 2))
    if isinstance(ls, (int, float)) and not isinstance(ls, bool):
        return ("mult", float(ls))
    return None


def _paragraph_state(p):
    """一段的关键版式属性：run 字体/字号/字间距 + 对齐/行距/缩进。"""
    ppr = p._p.pPr
    ind = ppr.find(qn("w:ind")) if ppr is not None else None
    return {
        "align": p.alignment,
        "line": _line_sig(p),
        "indent": tuple(sorted(ind.attrib.items())) if ind is not None else None,
        "runs": [_effective_run_state(p, r) for r in p.runs],
    }


def _assert_key_attributes_match(sample_path, output_path, kinds):
    sparas = _paragraphs(sample_path)
    oparas = _paragraphs(output_path)
    for label, pred in kinds:
        sp = _find_one(sparas, pred, label, sample_path)
        op = _find_one(oparas, pred, label, output_path)
        sample_state = _paragraph_state(sp)
        output_state = _paragraph_state(op)
        assert sample_state == output_state, (
            f"{label}的关键版式属性在样板与转换成品之间不一致："
            f"\n样板 {sample_state}\n成品 {output_state}")


@pytest.fixture(scope="module")
def tjyj_sample(tmp_path_factory):
    """tjyj 样板：该刊没有独立作者信息页，正文标题会被解析成 front_matter。

    gjjryj/cjyj 的样板前面有整页作者信息，解析器把真正的正文标题判成 body
    （不是本测试要管的渲染问题），转换会按正文套版——那两本只比正文/题注。
    """
    out = tmp_path_factory.mktemp("parity-sample") / "tjyj.sample.docx"
    run_tool("render_sample.py", RULES / "tjyj.json", "-o", out)
    return out


def test_same_ruleset_sample_and_conversion_agree(tjyj_sample, tmp_path):
    """tjyj：样板 → 同刊转换，标题/正文/题注的关键属性逐项一致。"""
    pipeline = render_into(tmp_path / "tjyj-self", tjyj_sample, target="tjyj")
    _assert_key_attributes_match(
        tjyj_sample, pipeline.docx,
        [("标题", _is_title), ("正文", _is_body), ("题注", _is_caption)])


def test_shipped_sample_body_and_caption_agree(sample_docx, tmp_path):
    """gjjryj 样板 → 同刊转换：正文与题注的关键属性一致。

    标题不在此列：该刊样板带独立作者信息页，解析结果把正文标题标成 body，
    转换按正文套版。这是解析分类的口径问题（另案），不是两个渲染器写值不同。
    """
    pipeline = render_into(tmp_path / "gjjryj-self", sample_docx, target="gjjryj")
    _assert_key_attributes_match(
        sample_docx, pipeline.docx,
        [("正文", _is_body), ("题注", _is_caption)])


def test_baseline_matches_target_sample(baseline, tjyj_sample):
    """conftest 的 baseline（gjjryj 源稿→tjyj）与 tjyj 样板同版式。

    两份成品的源稿结构不同，但目标刊是同一本；正文/题注的字体字号由目标规则集
    决定，必须与 tjyj 样板一致。这条同时守住「样板夹具」与「跨刊流水线夹具」
    不会各自漂移。
    """
    _assert_key_attributes_match(
        tjyj_sample, baseline.docx,
        [("正文", _is_body), ("题注", _is_caption)])
