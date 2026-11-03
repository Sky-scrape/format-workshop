"""三个真实事故的回归钉：标题漏判、标签归一污染整段格式、编号双重缀。

三条都出自同一份真实投稿（《管理世界》目标刊）：
- 24 个标题原稿都套了 Heading2/3，解析器却只看文字，「1 引言」不带点 → 全判成正文
- 「摘要：」改成「内容摘要：」时整段文字塞进第一个 run，标签的加粗把整段摘要染粗
- 标题从「1 引言」归一为「一、」时，旧正则只认「1.」，「一、1 引言」双重编号
"""

import importlib.util
import sys
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"


def _load(name, filename):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


parse = _load("parse_docx_under_test", "parse_docx.py")
render = _load("render_from_ir_under_test", "render_from_ir.py")


def _para(text_runs):
    """造一个只含文字与加粗的段落：(文本, 是否加粗) 列表。"""
    doc = Document()
    p = doc.add_paragraph()
    for txt, bold in text_runs:
        r = p.add_run(txt)
        r.font.bold = bold
    return p


def test_numbered_heading_without_dot_is_level1():
    """「1 引言」是 Word 多级列表的常见写法，必须认成一级标题。"""
    assert parse.classify_heading("1 引言") == (1, "high")
    assert parse.classify_heading("9 结论与局限") == (1, "high")


def test_numbered_subheading_without_dot_is_level2():
    """「2.1 研究方法」不能因为缺了点号就掉回正文。"""
    assert parse.classify_heading("2.1 AI驱动科学研究的范式与进展") == (2, "high")
    assert parse.classify_heading("5.4 局限与应对") == (2, "high")


def test_quantity_phrase_is_not_heading():
    """「3 个案例」这类正文短句不能被当标题，否则整篇编号被带偏。"""
    assert parse.classify_heading("3 个案例") is None
    assert parse.classify_heading("1 项结论") is None


def test_short_heading_survives_digit_ratio():
    """数字编号标题要绕过正文的数字占比检查。

    「1 引言」只有 4 个字符，数字占比 0.25 已逼近 0.3 的正文阈值；再短一点的
    「1 节」就会超过。所以数字编号规则必须排在 digit_ratio 检查之前。
    """
    assert parse.digit_ratio("1 节") > 0.3      # 正文阈值会挡掉它
    assert parse.classify_heading("1 节") is not None
    assert parse.classify_heading("1 引言") is not None


def test_single_level_arabic_heading_is_ambiguous():
    """「1.标题」是二级还是三级，文本本身说了不算。

    《经济管理》体例：一、／1.／（1）—— 「1.」是二级。
    「一、／（一）／1.」体系里它是三级。
    硬编码成三级会让前一类刊的合规稿整篇二级标题被降级保护打成正文，
    所以这里只给候选标记，由 parse() 按文档已出现的层级定级。
    """
    assert parse.classify_heading("1.目标责任制下的增长承诺") == (parse.AMBIGUOUS_H3, "medium")
    assert parse.classify_heading("2.两种对立的理论预期") == (parse.AMBIGUOUS_H3, "medium")


def test_bracketed_number_is_unambiguously_level4():
    """有中文括号包着的「（1）」不受歧义影响，仍是四级。"""
    assert parse.classify_heading("（1）模型设定与变量说明") == (4, "high")


def test_ambiguous_level_is_decided_once_per_document(tmp_path):
    """全文只定一次级：同一份稿子里「1.」不会一会儿二级一会儿三级。"""
    doc = Document()
    doc.add_paragraph("一、研究背景")          # 一级：中文顿号
    doc.add_paragraph("1.第一小节")            # 「1.」首次出现，此前无「（一）」
    doc.add_paragraph("正文内容一段。")
    doc.add_paragraph("1.第二小节")            # 同一章下重新从 1 起
    doc.add_paragraph("2.第三小节")
    path = tmp_path / "amb.docx"
    doc.save(str(path))

    ir = parse.parse(str(path), ai_cfg={"enabled": False})
    h2 = [b["text"] for b in ir["blocks"] if b["type"] == "heading2"]
    assert h2 == ["1.第一小节", "1.第二小节", "2.第三小节"], h2


def test_bracketed_second_level_wins_over_arabic(tmp_path):
    """稿子里先出现过「（一）」，说明二级另有写法，「1.」就该是三级。"""
    doc = Document()
    doc.add_paragraph("一、研究背景")
    doc.add_paragraph("（一）问题提出")        # 二级：中文括号
    doc.add_paragraph("1.第一点说明")          # 此时「1.」只能是三级
    path = tmp_path / "amb2.docx"
    doc.save(str(path))

    ir = parse.parse(str(path), ai_cfg={"enabled": False})
    assert [b["text"] for b in ir["blocks"] if b["type"] == "heading3"] == ["1.第一点说明"]
    assert [b["text"] for b in ir["blocks"] if b["type"] == "heading2"] == ["（一）问题提出"]


def test_bracket_second_level_appearing_later_still_wins(tmp_path):
    """「（一）」出现在第一个「1.」之后，全文仍按三级体系判。

    这条守的是「全文预扫 vs 边扫边判」的差别：边扫会把第一个「1.」定成
    二级，从此整篇级别错位。引言章直接进正文章节、二级标题晚于三级出现，
    在真实稿件里并不罕见。

    结构本身异常（三级早于二级）时宁可把那个降成正文——不套错样式，
    原文一个字不动。
    """
    doc = Document()
    doc.add_paragraph("一、研究背景")
    doc.add_paragraph("1.早出现的节")          # 位置异常：此时还没有任何二级
    doc.add_paragraph("（一）问题提出")
    doc.add_paragraph("1.具体做法")            # 二级之后的三级
    path = tmp_path / "amb3.docx"
    doc.save(str(path))

    ir = parse.parse(str(path), ai_cfg={"enabled": False})
    assert [b["text"] for b in ir["blocks"] if b["type"] == "heading2"] == ["（一）问题提出"]
    assert [b["text"] for b in ir["blocks"] if b["type"] == "heading3"] == ["1.具体做法"]
    early = next(b for b in ir["blocks"] if b["text"] == "1.早出现的节")
    assert early["type"] == "body", "三级早于二级时不套标题样式，降成正文更安全"


def test_no_second_level_trace_marks_inference(tmp_path):
    """全文只有「一、」和「1.」时，判二级是推断，置信度必须标出来。

    报告里要能看出「这一级是猜的」，不能把推断混进「期刊要求」。
    """
    doc = Document()
    doc.add_paragraph("一、研究背景")
    doc.add_paragraph("1.第一小节")
    doc.add_paragraph("正文内容一段。")
    doc.add_paragraph("1.第二小节")
    path = tmp_path / "amb4.docx"
    doc.save(str(path))

    ir = parse.parse(str(path), ai_cfg={"enabled": False})
    h2 = [b for b in ir["blocks"] if b["type"] == "heading2"]
    assert [b["text"] for b in h2] == ["1.第一小节", "1.第二小节"]
    assert all(b.get("confidence") == "ambiguous" for b in h2), \
        [b.get("confidence") for b in h2]


def test_arabic_third_marker_makes_level_certain(tmp_path):
    """出现「（1）」说明是《经济管理》体，判二级有依据，不再标推断。"""
    doc = Document()
    doc.add_paragraph("一、研究背景")
    doc.add_paragraph("1.第一小节")
    doc.add_paragraph("（1）具体做法")          # 阿拉伯三级编号，反证「1.」是二级
    path = tmp_path / "amb5.docx"
    doc.save(str(path))

    ir = parse.parse(str(path), ai_cfg={"enabled": False})
    h2 = [b for b in ir["blocks"] if b["type"] == "heading2"]
    assert [b["text"] for b in h2] == ["1.第一小节"]
    assert h2[0].get("confidence") == "medium", h2[0].get("confidence")


def test_label_rewrite_keeps_body_unbold():
    """标签归一只能改前缀：标签仍是粗体，正文必须保持原字重。

    旧实现把整段文字塞进第一个 run（带 <w:b/>），摘要整段跟着变粗——
    这是用户一眼就看出来的「奇怪」。
    """
    p = _para([("摘要", True), ("：人工智能正由辅助工具转变。", False)])
    render.replace_prefix(p, 3, "内容摘要：")

    assert p.text == "内容摘要：人工智能正由辅助工具转变。"
    assert p.runs[0].font.bold is True          # 标签继承原标签的加粗
    assert p.runs[0].text == "内容摘要："
    assert p.runs[1].font.bold is False         # 正文不被带跑
    assert p.runs[1].text.startswith("人工智能")


def test_label_rewrite_when_prefix_spans_runs():
    """标签横跨两个 run（「关键词」+「：」）时也要只改前缀。"""
    p = _para([("关键词", True), ("：科研智能体；人机协同", False)])
    render.replace_prefix(p, 4, "关键词：")
    assert p.text == "关键词：科研智能体；人机协同"
    assert p.runs[1].font.bold is False


def test_num_prefix_strips_both_forms():
    """一级编号剥离要同时认「1. 引言」和「1 引言」，且不吃掉「2.1」的前半截。"""
    rx = render.NUM_PREFIX_NUM[1]
    assert rx.sub("", "1 引言") == "引言"
    assert rx.sub("", "1. 引言") == "引言"
    assert rx.sub("", "2.1 研究设计") == "2.1 研究设计"   # 二级标题不能被一级规则削


def test_level1_heading_not_double_numbered():
    """「1 引言」归一成「一、引言」，不是「一、1 引言」。"""
    rx = render.NUM_PREFIX_NUM[1]
    core = rx.sub("", "1 引言").strip()
    assert f"一、{core}" == "一、引言"
