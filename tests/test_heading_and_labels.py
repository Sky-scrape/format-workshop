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
