"""覆盖广度：换刊矩阵 + 一份「乱排」稿件。

单个组合通过没有说服力——不同期刊走的分支不一样，缺哪个分支只有跑到了才知道。
这份稿件故意做得很脏：半角标点、多余空格、制表符撑开的表题、错序的表注、
带固定行距的图片、手写编号的各级标题。真实投稿稿大致就长这样。
"""

import base64
from io import BytesIO

import pytest
from docx import Document
from docx.shared import Pt

from conftest import invariant_failures, render_into

SHIPPED_JOURNALS = ["cjyj", "tjyj", "gjjryj"]

PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def build_messy_manuscript(path):
    """造一份格式很乱的成稿，尽量把渲染器的分支都踩一遍。"""
    doc = Document()

    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("——基于上市公司的经验证据")
    doc.add_paragraph("张三　李四")

    # 半角标点 + 多余空格 + 手写缩进空格
    doc.add_paragraph("内容摘要:本文利用2011-2020年A股上市公司数据 , 考察数字金融"
                      "对企业融资约束的影响 。研究发现,数字金融显著缓解了融资约束 。")
    doc.add_paragraph("关键词：数字金融　融资约束　企业创新")   # 全角空格分隔

    # 一级标题用全角空格撑开，二级用括号中文，三级用阿拉伯数字
    doc.add_paragraph("一、引　　言")
    doc.add_paragraph("近年来,数字金融快速发展 , 对微观主体产生了深远影响 。")
    doc.add_paragraph("（二）研究假设")
    doc.add_paragraph("假设1：数字金融缓解融资约束。")
    doc.add_paragraph("1. 样本选择")
    doc.add_paragraph("本文选取沪深A股上市公司为样本 , 剔除金融类企业 。")

    # 表题用制表符撑开；表注顺序故意写反（注在前、资料来源在后）
    doc.add_paragraph("表1\t变量定义与度量")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "变量"
    table.cell(0, 1).text = "定义"
    table.cell(1, 0).text = "Size"
    table.cell(1, 1).text = "企业规模"
    doc.add_paragraph("注：括号内为标准差；*** 表示 1% 水平显著。")
    doc.add_paragraph("资料来源：Wind数据库，作者整理。")

    # 图片段落带固定行距——这是最容易被裁的地方
    fig = doc.add_paragraph()
    fig.add_run().add_picture(BytesIO(PNG_1PX))
    fig.paragraph_format.line_spacing = Pt(20)
    doc.add_paragraph("图1\t融资约束变化趋势")
    doc.add_paragraph("资料来源：作者绘制。")

    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.add_paragraph("[2] 李四，2021：《融资约束的度量》，《金融研究》第7期。")

    doc.save(str(path))
    return path


@pytest.fixture(scope="module")
def messy_docx(tmp_path_factory):
    d = tmp_path_factory.mktemp("messy")
    return build_messy_manuscript(d / "messy.docx")


def test_messy_manuscript_covers_many_block_types(messy_docx, tmp_path):
    """先确认这份稿子确实踩到了各类块，否则下面的用例测了个寂寞。"""
    from conftest import body_texts, run_tool
    import json

    d = tmp_path / "probe"
    d.mkdir()
    run_tool("parse_docx.py", messy_docx, "-o", d / "m.ir.json")
    ir = json.loads((d / "m.ir.json").read_text(encoding="utf-8"))
    kinds = {b["type"] for b in ir["blocks"]}

    for expected in ("front_matter", "abstract", "keywords", "heading1", "heading2",
                     "heading3", "body", "table", "table_note", "figure", "references_heading"):
        assert expected in kinds, f"乱排稿件没有覆盖到 {expected}，实际：{sorted(kinds)}"


@pytest.mark.parametrize("target", SHIPPED_JOURNALS)
def test_sample_manuscript_to_each_journal(sample_docx, tmp_path, target):
    """同一份源稿排到每本示例期刊，四条不变式都要成立。"""
    p = render_into(tmp_path / f"sample-{target}", sample_docx, target=target)
    assert not invariant_failures(p), f"{target}: " + "；".join(invariant_failures(p))


@pytest.mark.parametrize("target", SHIPPED_JOURNALS)
def test_messy_manuscript_to_each_journal(messy_docx, tmp_path, target):
    """乱排稿件排到每本示例期刊，四条不变式同样要成立。"""
    p = render_into(tmp_path / f"messy-{target}", messy_docx, target=target)
    assert not invariant_failures(p), f"{target}: " + "；".join(invariant_failures(p))
    assert p.items, "乱排稿件没有产生任何改动，夹具可能没用上"
