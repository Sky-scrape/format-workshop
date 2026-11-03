"""匿名稿：删多了能补回，删少了是身份泄露。

匿名稿是从署名稿派生的一份副本，署名稿本身完整保留。所以删多了的代价
（基金信息缺失）可以从署名稿补回来；删少了的代价是审稿人看到你的名字，
不可挽回。因此策略是「宁可多删」，但多删必须在报告里说清楚。

规则集的 submission.anonymize_blocks 声明了各刊要删哪些块，写法差异很大：
有的含「通讯作者」，有的只要「致谢 + 基金」，甚至有一本要求匿名但清单为空。
"""

import json

import pytest
from docx import Document

from conftest import RULES, mutated_rules, run_tool

BIO = "作者简介：张三，某大学经济学院教授，电子信箱 zs@example.edu.cn"
ACK_HEAD = "致谢"
ACK_BODY = "感谢匿名审稿人的宝贵意见。"
AUTHOR_LINE = "张三　李四"


def build_docx(path, with_bio=True, with_ack=True):
    """作者行必须排第二段：匿名工具只保留第一个 front_matter 块（标题），
    其余前置块按位置删除（tools/anonymize.py 的 fm_seen 逻辑）——这段骨架
    不能换成 minimal_doc 的固定顺序，保持手拼。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph(AUTHOR_LINE)
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("正文段落，用于确认正文不会被删。")
    if with_bio:
        doc.add_paragraph(BIO)
    if with_ack:
        doc.add_paragraph(ACK_HEAD)
        doc.add_paragraph(ACK_BODY)
    doc.save(str(path))
    return path


def parse_and_patch_ir(tmp_path, docx_path):
    """解析后按文字内容改写块类型，构造出作者简介、致谢等块。

    真实流程里这些块由 apply_details 注入并带类型；这里手工构造，
    目的是只测 anonymize 是否按规则集执行。
    """
    ir_path = tmp_path / "m.ir.json"
    run_tool("parse_docx.py", docx_path, "-o", ir_path)
    ir = json.loads(ir_path.read_text(encoding="utf-8"))
    for b in ir["blocks"]:
        text = b.get("text", "")
        if text.startswith("作者简介"):
            b["type"] = "author_note"
        elif text.strip() == ACK_HEAD:
            b["type"] = "acknowledgement"
    ir_path.write_text(json.dumps(ir, ensure_ascii=False), encoding="utf-8")
    return ir_path


def alt_rules(tmp_path, name, blocks):
    """复制 tjyj 改匿名相关键（F58 公共入口；键已存在，点路径赋值即可）。"""
    return mutated_rules(tmp_path, "tjyj", name,
                         **{"submission.anonymous": True,
                            "submission.anonymize_blocks": blocks})


def run_anonymize(tmp_path, docx_path, ir_path, rules_path, name="anon.docx"):
    out = tmp_path / name
    stdout = run_tool("anonymize.py", docx_path, ir_path, rules_path, "-o", out)
    removed = json.loads(stdout.strip().splitlines()[-1])["removed"]
    return out, removed


def body_text(docx_path):
    return "\n".join(p.text for p in Document(str(docx_path)).paragraphs)


def test_author_line_is_removed_title_kept(tmp_path):
    doc = build_docx(tmp_path / "a.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    out, _ = run_anonymize(tmp_path, doc, ir, RULES / "tjyj.json")

    text = body_text(out)
    assert AUTHOR_LINE not in text, "作者署名行没删"
    assert "数字金融与企业融资约束" in text, "标题被误删"
    assert "正文段落" in text, "正文被误删"


def test_acknowledgement_removed_when_declared(tmp_path):
    doc = build_docx(tmp_path / "b.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    out, removed = run_anonymize(tmp_path, doc, ir, RULES / "tjyj.json")

    assert ACK_BODY not in body_text(out), "规则集声明要删致谢，却没删"
    assert any("致谢" in r for r in removed)


def test_acknowledgement_kept_when_not_declared(tmp_path):
    """规则集没声明删致谢时，致谢要留着。"""
    doc = build_docx(tmp_path / "c.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    rules = alt_rules(tmp_path, "no_ack.json", ["authors", "affiliations"])
    out, _ = run_anonymize(tmp_path, doc, ir, rules)

    assert ACK_BODY in body_text(out), "规则集没要求删致谢，却被删了"


def test_undeclared_block_is_removed_but_flagged(tmp_path):
    """规则集没声明作者简介，仍按安全优先删掉，但要在报告里标出来。"""
    doc = build_docx(tmp_path / "d.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    rules = alt_rules(tmp_path, "no_bio.json", ["acknowledgement"])
    out, removed = run_anonymize(tmp_path, doc, ir, rules)

    assert BIO not in body_text(out), "作者简介含邮箱，必须删"
    flagged = [r for r in removed if "author_note" in r]
    assert flagged and "规则集未声明" in flagged[0], f"多删了却没标注：{removed}"


def test_declared_block_is_not_flagged(tmp_path):
    """规则集声明过的，不该打「未声明」的标记。"""
    doc = build_docx(tmp_path / "e.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    rules = alt_rules(tmp_path, "declared.json",
                      ["acknowledgement", "author_note", "authors", "affiliations"])
    _, removed = run_anonymize(tmp_path, doc, ir, rules)

    bio = next(r for r in removed if "author_note" in r)
    assert "规则集未声明" not in bio, f"声明过的块被标成未声明：{bio}"


def test_unknown_declared_type_is_reported(tmp_path):
    """规则集声明了工具不认识的类型，必须报出来，不能静默跳过。"""
    doc = build_docx(tmp_path / "f.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    rules = alt_rules(tmp_path, "unknown.json",
                      ["acknowledgement", "biography_page"])
    _, removed = run_anonymize(tmp_path, doc, ir, rules)

    assert any("未识别" in r and "biography_page" in r for r in removed), (
        f"未识别类型没有报出来：{removed}")


def test_empty_declaration_still_anonymizes_identity(tmp_path):
    """要求匿名但清单为空时（数量经济技术经济研究就是这样），身份信息仍要删。"""
    doc = build_docx(tmp_path / "g.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    rules = alt_rules(tmp_path, "empty.json", [])
    out, removed = run_anonymize(tmp_path, doc, ir, rules)

    text = body_text(out)
    assert AUTHOR_LINE not in text, "清单为空时作者行没删"
    assert BIO not in text, "清单为空时作者简介没删"
    assert any("规则集未声明" in r for r in removed), "多删没有标注"


def test_signature_copy_is_untouched(tmp_path):
    """匿名稿是副本，署名稿必须原封不动。"""
    doc = build_docx(tmp_path / "h.docx")
    before = doc.read_bytes()
    ir = parse_and_patch_ir(tmp_path, doc)
    run_anonymize(tmp_path, doc, ir, RULES / "tjyj.json")
    assert doc.read_bytes() == before, "生成匿名稿时改动了署名稿"


def test_journal_without_anonymous_refuses(tmp_path):
    """不要求匿名的刊，不该产出匿名稿。"""
    doc = build_docx(tmp_path / "i.docx")
    ir = parse_and_patch_ir(tmp_path, doc)
    # 不走 alt_rules（它固定 anonymous=True），一次到位改两个键（F58）
    rules = mutated_rules(tmp_path, "tjyj", "no_anon.json",
                          **{"submission.anonymous": False,
                             "submission.anonymize_blocks": []})

    import subprocess
    import sys
    from conftest import TOOLS
    r = subprocess.run([sys.executable, str(TOOLS / "anonymize.py"), str(doc), str(ir),
                        str(rules), "-o", str(tmp_path / "x.docx")],
                       capture_output=True, text=True)
    assert r.returncode != 0, "不要求匿名却产出了匿名稿"
