"""脚注编号格式与重编号范围的自动对齐（DESIGN §12.1，2026-11-15）。

渲染器只写文档级设置（settings.xml 的 w:footnotePr），不碰正文与脚注
文字。真脚注对象才处理；尾注仍归「需你处理」。体检侧同一偏差改标
「自动修正」，不再让用户去 Word 里手工改。
"""

import importlib.util
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))


def _render():
    name = "render_from_ir_footnote_under_test"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / "render_from_ir.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _doc_with_footnote():
    doc = Document()
    p = doc.add_paragraph("正文提到一个注释")
    ref = OxmlElement("w:footnoteReference")
    ref.set(qn("w:id"), "1")
    run = p.add_run()
    run._r.append(ref)
    return doc


def _settings(doc):
    return doc.settings.element


def _set_pr(doc, fmt=None, restart=None):
    settings = _settings(doc)
    pr = settings.find(qn("w:footnotePr"))
    if pr is None:
        pr = OxmlElement("w:footnotePr")
        settings.append(pr)
    if fmt:
        el = OxmlElement("w:numFmt")
        el.set(qn("w:val"), fmt)
        pr.append(el)
    if restart:
        el = OxmlElement("w:numRestart")
        el.set(qn("w:val"), restart)
        pr.append(el)


def _read_pr(doc):
    pr = _settings(doc).find(qn("w:footnotePr"))
    if pr is None:
        return None, None
    fmt = pr.find(qn("w:numFmt"))
    restart = pr.find(qn("w:numRestart"))
    return (fmt.get(qn("w:val")) if fmt is not None else None,
            restart.get(qn("w:val")) if restart is not None else None)


def test_aligns_format_and_scope_and_audits():
    mod = _render()
    doc = _doc_with_footnote()
    _set_pr(doc, fmt="decimal", restart="continuous")
    rules = {"numbering": {"footnotes": "①②③", "footnotes_scope": "per_page"}}
    audit = mod.Audit()

    mod.align_footnote_numbering(doc, rules, audit)

    assert _read_pr(doc) == ("decimalEnclosedCircle", "eachPage")
    actions = {it["action"]: it for it in audit.items}
    assert set(actions) == {"脚注编号样式对齐", "脚注重编号范围对齐"}
    assert actions["脚注编号样式对齐"]["rule"] == "numbering.footnotes"
    assert actions["脚注重编号范围对齐"]["rule"] == "numbering.footnotes_scope"
    assert all("Word 脚注" in it["note"] for it in audit.items)


def test_no_change_no_audit_and_idempotent():
    mod = _render()
    doc = _doc_with_footnote()
    _set_pr(doc, fmt="decimalEnclosedCircle", restart="eachPage")
    rules = {"numbering": {"footnotes": "①②③", "footnotes_scope": "per_page"}}
    audit = mod.Audit()

    mod.align_footnote_numbering(doc, rules, audit)
    assert audit.items == []

    doc2 = _doc_with_footnote()
    _set_pr(doc2, fmt="decimal", restart="continuous")
    audit2 = mod.Audit()
    mod.align_footnote_numbering(doc2, rules, audit2)
    first = len(audit2.items)
    mod.align_footnote_numbering(doc2, rules, audit2)
    assert len(audit2.items) == first, "第二次对齐不应再出条目"


def test_no_footnote_reference_no_settings_touched():
    mod = _render()
    doc = Document()
    doc.add_paragraph("没有脚注的正文")
    rules = {"numbering": {"footnotes": "①②③", "footnotes_scope": "per_page"}}
    audit = mod.Audit()

    mod.align_footnote_numbering(doc, rules, audit)

    assert audit.items == []
    assert _read_pr(doc) == (None, None), "无脚注稿件不应被写入 footnotePr"
