"""标题编号：模板写哪种数字，编号就必须跟着哪种数字递增。

两个真实事故：

1. 《经济管理》《南开管理评论》的二级标题是「1.」。渲染器只拿中文数字正则去
   替换模板，模板里没有中文数字，替换空转，整篇二级标题都印成「1.」。
2. 一级标题到第 11 个就 IndexError（`CN_SEQ` 只有十位），整篇转换直接失败。

两处都只在**换到特定刊物**时才暴露，原有矩阵测试覆盖的三本刊都走中文序号，
所以一直没被踩到。
"""

import json
import re

import pytest

from conftest import RULES, TARGET_JOURNAL, body_texts, mutated_rules, render_into
from docx import Document

import importlib.util
import sys
from pathlib import Path

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


render = _load("render_from_ir_numbering_under_test", "render_from_ir.py")


def build_headings_docx(path, sections=3, subs_per_section=2, style="cn"):
    """造一份只有各级标题的稿子，用来单看编号。

    style="cn" 用「一、／（一）」，style="num" 用「1 ／（1）」——后者用来
    制造源稿与目标刊编号体系不一致的场面，编号重排才会真的发生。

    中文数字用字面量清单（F59）：夹具数据不能经过被测函数 cn_numeral，
    否则它错了夹具跟着错，「照着实现抄数据」的用例永远绿。
    """
    cn = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"]
    assert max(sections, subs_per_section) <= len(cn), "字面量清单只备到十二"
    doc = Document()
    doc.add_paragraph("标题编号测试稿")
    for s in range(1, sections + 1):
        head = f"{cn[s - 1]}、第{s}章" if style == "cn" else f"{s} 第{s}章"
        doc.add_paragraph(head)
        doc.add_paragraph(f"第{s}章的正文。")
        for t in range(1, subs_per_section + 1):
            sub = (f"（{cn[t - 1]}）第{s}章第{t}节" if style == "cn"
                   else f"（{t}）第{s}章第{t}节")
            doc.add_paragraph(sub)
            doc.add_paragraph(f"第{s}章第{t}节的正文。")
    doc.save(str(path))
    return path


HEADING_LIKE = re.compile(
    r"^(\d{1,2}\s|[一二三四五六七八九十]{1,3}、"
    r"|[（(][一二三四五六七八九十]{1,3}[）)]|[（(]\d{1,2}[）)]|\d{1,2}[.．]\S)")


def heading_texts(docx_path):
    """按出现顺序取标题段文字（只认带编号前缀的短行）。"""
    out = []
    for p in Document(str(docx_path)).paragraphs:
        t = p.text.strip()
        if t and len(t) <= 24 and HEADING_LIKE.match(t):
            out.append(t)
    return out


# ---------- 纯函数 ----------

def test_cn_numeral_past_ten():
    """超过十不能崩，也不能退化成阿拉伯数字。"""
    assert render.cn_numeral(1) == "一"
    assert render.cn_numeral(10) == "十"
    assert render.cn_numeral(11) == "十一"
    assert render.cn_numeral(12) == "十二"
    assert render.cn_numeral(20) == "二十"
    assert render.cn_numeral(21) == "二十一"


def test_fill_template_handles_both_numeral_styles():
    """模板里有中文数字改中文数字，有阿拉伯数字改阿拉伯数字。"""
    assert render.fill_level_template("一、", 3) == "三、"
    assert render.fill_level_template("（一）", 3) == "（三）"
    assert render.fill_level_template("1.", 3) == "3."
    assert render.fill_level_template("（1）", 3) == "（3）"
    assert render.fill_level_template("一、", 11) == "十一、"


def test_fill_template_only_touches_the_sequence_placeholder():
    """「（一）」里的括号不能被吃掉。"""
    assert render.fill_level_template("（一） ", 2) == "（二） "


# ---------- 端到端 ----------

def test_arabic_level2_template_numbers_increment(tmp_path):
    """二级标题模板是「1.」时，编号必须 1、2、3 递增，而不是全印成 1.。"""
    ruleset = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    assert ruleset["headings"]["numbering"] == "cn-hierarchical"
    # 二级标题改阿拉伯数字（经济管理 / 南开管理评论 的写法）；复制规则集改一键
    # 走公共入口（F58），列表下标用数值段表示
    rules_path = mutated_rules(tmp_path, TARGET_JOURNAL, "arabic-level2.json",
                               **{"headings.levels.1": "1."})

    src = build_headings_docx(tmp_path / "src.docx", sections=2, subs_per_section=3)
    p = render_into(tmp_path / "run", src, target=TARGET_JOURNAL, rules_path=rules_path)

    got = [t for t in heading_texts(p.docx) if t[0].isdigit()]
    assert got == ["1.第1章第1节", "2.第1章第2节", "3.第1章第3节",
                   "1.第2章第1节", "2.第2章第2节", "3.第2章第3节"], (
        f"二级编号没有按章递增：{got}")


def test_eleven_sections_do_not_crash(tmp_path):
    """第 11 个一级标题以前会 IndexError，把整篇转换崩成「转换失败」。"""
    src = build_headings_docx(tmp_path / "src12.docx", sections=12, subs_per_section=0)
    p = render_into(tmp_path / "run12", src, target=TARGET_JOURNAL)

    texts = body_texts(p.docx).values()
    assert any(t.startswith("十一、") for t in texts), (
        f"第十一个一级标题没编出来，实际标题：{[t for t in texts if '章' in t][:14]}")
    assert any(t.startswith("十二、") for t in texts)


def test_numbering_changes_are_audited(tmp_path):
    """编号重排要逐条进报告，用户才能驳回。"""
    # 源稿用「1 第1章」这种数字编号，目标刊要「一、」，重排一定发生
    src = build_headings_docx(tmp_path / "src_audit.docx", sections=2,
                              subs_per_section=2, style="num")
    p = render_into(tmp_path / "run_audit", src, target=TARGET_JOURNAL)

    items = [it for it in p.items if it["action"] == "标题重新编号"]
    assert items, "编号改过却没有审计条目"
    assert all(it["rule"] == "headings.numbering" for it in items)
    # 二级标题在统计研究里是「（一）」，与源稿写法一致时不产生条目；
    # 只要有一条记录，它的 after 就该带着目标刊的编号样式
    assert all(it["after"].strip() for it in items)
