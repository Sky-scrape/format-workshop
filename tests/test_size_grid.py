"""字号必须落在 Word 的半磅刻度上。

Word 的字号存成 half-point 整数（w:sz=24 即 12pt）。规则集里若写 11.8，
python-docx 会截成 11.5——规范页说 11.8、交付稿却是 11.5，作者按规范核对时会以为
自己排错了。实测值带小数是正常的（PDF 里的字号读数），但进规则集前必须归整到刻度上，
原始读数保留在 measured 里。
"""

import json
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
RULES = ROOT / "rules" / "journal"
sys.path.insert(0, str(ROOT / "tools"))

SIZE_KEYS = {"size_pt", "font_size_pt", "caption_size_pt", "value_pt"}
SKIP = {"measured", "measured_raw", "size_pt_raw"}


def iter_sizes(node, path=()):
    if isinstance(node, dict):
        for k, v in node.items():
            if k in SKIP:
                continue
            if k in SIZE_KEYS and isinstance(v, (int, float)) and not isinstance(v, bool):
                yield path + (k,), float(v)
            else:
                yield from iter_sizes(v, path + (k,))
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from iter_sizes(v, path + (i,))


@pytest.mark.parametrize("rules_path", sorted(RULES.glob("*.json")), ids=lambda p: p.stem)
def test_sizes_on_half_point_grid(rules_path):
    rules = json.loads(rules_path.read_text(encoding="utf-8"))
    bad = [
        (".".join(str(p) for p in path), val)
        for section in ("typography", "headings", "tables", "figures")
        for path, val in iter_sizes(rules.get(section, {}), (section,))
        if abs(val * 2 - round(val * 2)) > 1e-9
    ]
    assert not bad, f"{rules_path.stem} 有非半磅字号：{bad}"


def test_renderer_rounds_instead_of_truncating():
    """渲染器自己也要兜住：11.8 进来，写进 docx 的是 12pt，不是 11.5pt。"""
    import render_from_ir

    doc = Document()
    run = doc.add_paragraph().add_run("核验")
    render_from_ir.set_run(run, cn="宋体", size=11.8)
    sz = run._element.rPr.find(qn("w:sz")).get(qn("w:val"))
    assert sz == "24", f"11.8pt 应写成 24 半磅（12pt），实际 {sz}"
