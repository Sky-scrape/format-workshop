"""四级标题的排版支持。

两条边界：
- 目标刊 `max_level: 4` 时，Heading 4 样式段按 `headings.level4` 套版并重写编号；
  `inline_with_body: true` 的刊按正文套版，只换编号前缀。
- 目标刊只排到三级时，不硬套标题模板：按正文套版，并在审计里出一条
  「标题层级超出本刊」的说明（changed=False），编号原样保留。
"""

import json
import subprocess
import sys
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
RULES = ROOT / "rules" / "journal"


def _manuscript(path):
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融。")
    doc.add_paragraph("关键词：数字金融；融资约束")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("这是正文段落，用来测试四级标题的识别。")
    p = doc.add_paragraph("（1）机制检验")
    p.style = doc.styles["Heading 4"]
    doc.add_paragraph("这是四级标题下面的正文段落。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.save(str(path))
    return path


def _render(tmp_path, jid):
    src = _manuscript(tmp_path / "src.docx")
    ir = tmp_path / "src.ir.json"
    subprocess.run([sys.executable, str(ROOT / "tools" / "parse_docx.py"),
                    str(src), "-o", str(ir)], check=True, cwd=str(ROOT))
    ir_data = json.loads(ir.read_text(encoding="utf-8"))
    idx = next(b["idx"] for b in ir_data["blocks"] if b.get("type") == "heading4")
    out = tmp_path / f"out-{jid}.docx"
    subprocess.run([sys.executable, str(ROOT / "tools" / "render_from_ir.py"),
                    str(ir), str(RULES / f"{jid}.json"), "-o", str(out)],
                   check=True, cwd=str(ROOT))
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    return out, audit, idx


def _heading4_para(docx):
    doc = Document(str(docx))
    return next(p for p in doc.paragraphs if "机制检验" in p.text)


def test_heading4_within_max_level_is_styled_and_renumbered(tmp_path):
    out, audit, idx = _render(tmp_path, "cjyj")
    p = _heading4_para(out)
    assert p.text.startswith("(1)"), p.text
    # cjyj level4 是 inherit: body + inline_with_body：字号跟正文 10pt
    assert {r.font.size.pt for r in p.runs} == {10.0}
    items = [it for it in audit["items"] if it.get("idx") == idx]
    assert any(it["action"] == "标题套版" and it["rule"].startswith("headings.level4")
               for it in items), items
    assert any(it["action"] == "标题重新编号" for it in items), items
    assert not [it for it in items if it["action"] == "标题层级超出本刊"]


def test_heading4_beyond_max_level_falls_back_to_body_with_note(tmp_path):
    out, audit, idx = _render(tmp_path, "tjyj")
    p = _heading4_para(out)
    assert p.text.startswith("（1）"), "超出层级时不应按目标模板重编号"
    assert {r.font.size.pt for r in p.runs} == {10.5}
    items = [it for it in audit["items"] if it.get("idx") == idx]
    note = next(it for it in items if it["action"] == "标题层级超出本刊")
    assert note["changed"] is False and note["rule"] == "headings.max_level"
    assert any(it["action"] == "正文套版" for it in items), items


def test_heading4_renumber_passes_text_guard(tmp_path):
    """四级标题的编号重写必须被白名单认，不能被运行期文字闸误杀。"""
    out, audit, _idx = _render(tmp_path, "cjyj")
    src = tmp_path / "src.docx"
    ir = json.loads((tmp_path / "src.ir.json").read_text(encoding="utf-8"))
    block_types = {b["idx"]: b.get("type", "") for b in ir["blocks"]}

    from text_guard import document_texts, whitelist_violations

    violations = whitelist_violations(
        document_texts(src), document_texts(out), block_types, audit["items"])
    assert violations == [], violations


def test_heading4_second_render_is_idempotent(tmp_path):
    out, _audit, _idx = _render(tmp_path, "cjyj")
    # 以成品为源再排一遍：Heading 4 样式还在，四级标题不应再被改。
    # render_from_ir 按 IR 文件名找同 stem 的 docx，所以 IR 必须与源同名。
    ir = out.with_suffix(".ir.json")
    subprocess.run([sys.executable, str(ROOT / "tools" / "parse_docx.py"),
                    str(out), "-o", str(ir)], check=True, cwd=str(ROOT))
    again = tmp_path / "again.docx"
    subprocess.run([sys.executable, str(ROOT / "tools" / "render_from_ir.py"),
                    str(ir), str(RULES / "cjyj.json"), "-o", str(again)],
                   check=True, cwd=str(ROOT))
    audit = json.loads(again.with_suffix(".audit.json").read_text(encoding="utf-8"))
    h4 = [it for it in audit["items"]
          if it.get("block") == "heading4" and it.get("changed", True)]
    assert h4 == [], h4
