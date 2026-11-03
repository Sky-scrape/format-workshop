"""带图稿件：图片段落不能被设成固定行距。

Word 里「固定值」行距会把内嵌图裁成一条——12cm 高的图只剩 18 磅高。
这是排版工具最容易犯、也最难在页面上发现的错：图不会报错，只会消失一半。
"""

import base64
from io import BytesIO

from docx import Document
from docx.shared import Pt

from conftest import NS_W, minimal_doc, render_into

# 1x1 透明 PNG，用来构造真实的图片段落
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def build_manuscript_with_figure(path):
    """造一份最小稿件：标题 + 一张图（原稿带固定行距）+ 图注 + 正文。"""
    doc = minimal_doc(path, title="经济增长与企业创新研究",
                      abstract="内容提要：本文考察企业创新产出的影响因素。")

    fig = doc.add_paragraph()
    fig.add_run().add_picture(BytesIO(PNG_1PX))
    fig.paragraph_format.line_spacing = Pt(18)   # 原稿常见的固定行距

    doc.add_paragraph("图1 企业创新产出趋势")
    doc.add_paragraph("本文发现创新产出与研发投入显著正相关。")
    doc.save(str(path))
    return path


def find_drawing_paragraph(docx_path):
    """返回正文里第一个含图片的段落（及其 body 下标）。"""
    doc = Document(str(docx_path))
    for i, child in enumerate(doc.element.body):
        if child.tag != NS_W + "p":
            continue
        if any(n.tag == NS_W + "drawing" for n in child.iter()):
            return i, child
    return None, None


def test_figure_paragraph_does_not_get_fixed_line_spacing(tmp_path):
    """图片段落必须是单倍/自动行距，不能是固定值。"""
    src = build_manuscript_with_figure(tmp_path / "with_figure.docx")
    idx, _ = find_drawing_paragraph(src)
    assert idx is not None, "构造的稿件里没有图片段落，夹具本身有问题"

    result = render_into(tmp_path / "run", src, target="tjyj")
    out_idx, node = find_drawing_paragraph(result.docx)
    assert out_idx is not None, "成品里图片丢了——比裁图更严重"

    spacing = node.find(f".//{NS_W}pPr/{NS_W}spacing")
    if spacing is not None:
        assert spacing.get(NS_W + "lineRule") != "exact", (
            f"图片段的行距被设成了固定值 {spacing.get(NS_W + 'line')}，会把图裁掉")


def test_figure_restyling_is_audited(tmp_path):
    """解除固定行距这件事也要进报告，否则用户不知道这里动过。"""
    src = build_manuscript_with_figure(tmp_path / "with_figure2.docx")
    result = render_into(tmp_path / "run2", src, target="tjyj")
    actions = {it["action"] for it in result.items}
    assert "图片段解除固定行距" in actions, f"没有记录图片段处理，实际只有：{sorted(actions)}"
