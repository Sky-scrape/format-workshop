"""匿名覆盖缺口的回归测试：F05 / F06 / F10，用例直接来自审计证据的复现。

F06:通讯作者/作者简介类脚注没有单位词时，「单位+身份」双条件漏删。
F10:「-」打头的前置区单位/地址行被当副标题静默保留，且不进 removed。
F05:页眉、默认页脚、批注正文、尾注既不在匿名清理范围，闸门也查不到。

批注/脚注/尾注部件不在 python-docx 对象模型里，夹具按包结构手工注入，
再确认匿名稿产出后部件内容与正文引用标记都被清掉。
"""

import json
import re
import sys
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import docx_meta  # noqa: E402

from conftest import RULES, run_tool  # noqa: E402

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = {
    "footnotes": ("application/vnd.openxmlformats-officedocument."
                  "wordprocessingml.footnotes+xml"),
    "endnotes": ("application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.endnotes+xml"),
    "comments": ("application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.comments+xml"),
}

FOOTNOTE_CONTACT = "通讯作者：张三，Email：zs@pku.edu.cn"          # 无单位词,旧逻辑漏删
FOOTNOTE_BIO = "作者简介：张三，研究方向为数字金融，Email：zs@pku.edu.cn"  # 同上
FOOTNOTE_CONTENT = "数据来自北京大学。"                            # 内容注:必须保留


def build_docx(path):
    """四段式合成稿：标题、提要、关键词、正文。返回 Document 便于挂引用标记。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束")
    doc.add_paragraph("正文段落，用于确认正文不会被删。")
    doc.save(str(path))
    return doc


def alt_rules(tmp_path, blocks=("acknowledgement",)):
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["submission"]["anonymous"] = True
    rules["submission"]["anonymize_blocks"] = list(blocks)
    path = tmp_path / "anon-rules.json"
    path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return path


def run_anonymize(tmp_path, docx_path, name="anon.docx"):
    """解析 → 匿名化。anonymize 内部跑 strict 闸门，能产出即闸门已通过。"""
    ir_path = tmp_path / "m.ir.json"
    run_tool("parse_docx.py", docx_path, "-o", ir_path)
    out = tmp_path / name
    stdout = run_tool("anonymize.py", docx_path, ir_path, alt_rules(tmp_path),
                      "-o", out)
    removed = json.loads(stdout.strip().splitlines()[-1])["removed"]
    return out, removed


def add_reference(doc, tag, id_):
    """在正文末段挂一个脚注/尾注/批注引用标记。"""
    doc.paragraphs[3]._p.append(
        parse_xml(f'<w:r {nsdecls("w")}><w:{tag} w:id="{id_}"/></w:r>'))


def add_part(path, name, xml_text, kind, rel_id):
    """往包里塞一个部件（脚注/尾注/批注），补上内容类型与关系。"""
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    ct = items["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>", f'<Override PartName="/{name}" '
                                f'ContentType="{CT[kind]}"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")
    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="{rel_id}" Type="{REL_BASE}{kind}" '
                        f'Target="{Path(name).name}"/></Relationships>')
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    items[name] = xml_text.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)
    return path


def part_texts(path, name, tag):
    """取某部件里各节点（w:footnote / w:endnote / w:comment）的文本。"""
    with zipfile.ZipFile(path) as z:
        if name not in z.namelist():
            return []
        root = docx_meta.parse_part_xml(z.read(name))
    return ["".join(t.text or "" for t in n.iter(qn("w:t")))
            for n in root.iter(qn(f"w:{tag}"))]


def body_has_reference(path, tag):
    doc = Document(str(path))
    return any(True for _ in doc.element.body.iter(qn(f"w:{tag}")))


def story_texts(path):
    """全部页眉/页脚部件的文本，形如 [(部件名, 文本), ...]。"""
    out = []
    with zipfile.ZipFile(path) as z:
        for name in sorted(z.namelist()):
            if re.match(r"^word/(?:header|footer)\d*\.xml$", name):
                root = docx_meta.parse_part_xml(z.read(name))
                out.append((name, "".join(t.text or "" for t in root.iter(qn("w:t")))))
    return out


# ---------------------------------------------------------------- F06 脚注强信号


def footnotes_part(*texts):
    notes = "".join(
        f'<w:footnote w:id="{i + 1}"><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:footnote>'
        for i, t in enumerate(texts))
    return (f'<w:footnotes xmlns:w="{W_NS}">'
            '<w:footnote w:type="separator" w:id="-1">'
            '<w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0">'
            '<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            f'{notes}</w:footnotes>')


def test_strong_signal_footnotes_removed_without_unit_word(tmp_path):
    """F06 证据复现:通讯作者/作者简介脚注没有单位词,单靠双条件漏删。"""
    doc = build_docx(tmp_path / "a.docx")
    add_reference(doc, "footnoteReference", 1)
    add_reference(doc, "footnoteReference", 2)
    doc.save(str(tmp_path / "a.docx"))
    add_part(tmp_path / "a.docx", "word/footnotes.xml",
             footnotes_part(FOOTNOTE_CONTACT, FOOTNOTE_BIO), "footnotes", "rIdFixF")

    out, removed = run_anonymize(tmp_path, tmp_path / "a.docx")

    texts = " ".join(part_texts(out, "word/footnotes.xml", "footnote"))
    assert "通讯作者" not in texts and "zs@pku.edu.cn" not in texts, \
        "无单位词的通讯作者/简介脚注漏删（身份泄露）"
    assert any("脚注" in r and "通讯作者" in r for r in removed), f"删除没进报告：{removed}"
    assert not body_has_reference(out, "footnoteReference"), "正文引用标记没摘掉"


def test_content_note_without_identity_is_kept(tmp_path):
    """F06 反例:「数据来自北京大学」这类内容注不能被强信号误删。"""
    doc = build_docx(tmp_path / "b.docx")
    add_reference(doc, "footnoteReference", 1)
    doc.save(str(tmp_path / "b.docx"))
    add_part(tmp_path / "b.docx", "word/footnotes.xml",
             footnotes_part(FOOTNOTE_CONTENT), "footnotes", "rIdFixF")

    out, removed = run_anonymize(tmp_path, tmp_path / "b.docx")

    assert any("数据来自北京大学" in t
               for t in part_texts(out, "word/footnotes.xml", "footnote")), "内容注被误删"
    assert not any("脚注" in r and "数据来自北京大学" in r for r in removed)


def test_weak_signal_double_condition_still_works(tmp_path):
    """弱信号兜底:单位+身份双命中的老路径仍要删。"""
    doc = build_docx(tmp_path / "w.docx")
    add_reference(doc, "footnoteReference", 1)
    doc.save(str(tmp_path / "w.docx"))
    add_part(tmp_path / "w.docx", "word/footnotes.xml",
             footnotes_part("张三，北京大学经济学院，邮政编码 100871"),
             "footnotes", "rIdFixF")

    out, removed = run_anonymize(tmp_path, tmp_path / "w.docx")

    texts = " ".join(part_texts(out, "word/footnotes.xml", "footnote"))
    assert "北京大学经济学院" not in texts
    assert any("脚注" in r and "北京大学经济学院" in r for r in removed)


def test_endnote_author_bio_removed(tmp_path):
    """F05:尾注里的作者简介此前完全无人处理。"""
    doc = build_docx(tmp_path / "c.docx")
    add_reference(doc, "endnoteReference", 1)
    doc.save(str(tmp_path / "c.docx"))
    endnotes = (f'<w:endnotes xmlns:w="{W_NS}">'
                '<w:separatorNs/>'
                '<w:endnote w:id="1"><w:p><w:r>'
                '<w:t>作者简介：张三，Email：zs@pku.edu.cn</w:t>'
                '</w:r></w:p></w:endnote></w:endnotes>')
    endnotes = endnotes.replace('<w:separatorNs/>',
                                '<w:endnote w:type="separator" w:id="-1">'
                                '<w:p><w:r><w:separator/></w:r></w:p></w:endnote>')
    add_part(tmp_path / "c.docx", "word/endnotes.xml", endnotes, "endnotes", "rIdFixE")

    out, removed = run_anonymize(tmp_path, tmp_path / "c.docx")

    texts = " ".join(part_texts(out, "word/endnotes.xml", "endnote"))
    assert "作者简介" not in texts and "zs@pku.edu.cn" not in texts, "尾注身份没清"
    assert any("尾注" in r and "作者简介" in r for r in removed), f"删除没进报告：{removed}"
    assert not body_has_reference(out, "endnoteReference")


# ---------------------------------------------------------------- F10 前置区连字符行


def test_dash_front_matter_identity_lines_removed(tmp_path):
    """F10 证据复现:「-」打头的单位/地址行被当副标题静默保留。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("- 北京大学经济学院，北京 100871")
    doc.add_paragraph("- 通信地址：北京市海淀区中关村大街59号，100871")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束")
    doc.add_paragraph("正文段落。")
    doc.save(str(tmp_path / "d.docx"))

    out, removed = run_anonymize(tmp_path, tmp_path / "d.docx")

    text = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "北京大学经济学院" not in text, "连字符单位行留在匿名稿"
    assert "中关村大街59号" not in text, "连字符地址行留在匿名稿"
    assert any("北京大学经济学院" in r for r in removed), f"删除没进报告：{removed}"
    assert any("中关村大街59号" in r for r in removed)
    assert "数字金融与企业融资约束" in text, "标题被误删"


def test_real_subtitle_and_clean_dash_line_kept(tmp_path):
    """F10 反例:破折号副标题与不含身份词的连字符行都要保留。"""
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("——来自制造业上市公司的证据")
    doc.add_paragraph("- 基于企业调查数据的研究")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文段落。")
    doc.save(str(tmp_path / "s.docx"))

    out, removed = run_anonymize(tmp_path, tmp_path / "s.docx")

    text = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "来自制造业上市公司的证据" in text, "破折号副标题被误删"
    assert "基于企业调查数据的研究" in text, "无身份词的连字符行被误删"
    assert not any("证据" in r or "调查数据" in r for r in removed)


# ---------------------------------------------------------------- F05 页眉页脚/批注


def test_headers_footers_identity_cleared(tmp_path):
    """F05:页眉、默认页脚、偶数页脚、页脚内表格此前无人清理。"""
    doc = build_docx(tmp_path / "e.docx")
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    sec.header.paragraphs[0].add_run("张三，北京大学经济学院，Email：zs@pku.edu.cn")
    sec.header.add_paragraph("工作论文（请勿引用）")  # 中性内容：必须保留
    sec.footer.paragraphs[0].add_run("通讯作者：李四，zs2@pku.edu.cn")
    tbl = sec.footer.add_table(rows=1, cols=2, width=Cm(15))  # 页脚内表格：旧实现够不着
    tbl.cell(0, 0).text = "通讯作者：李四"
    tbl.cell(0, 1).text = "zs2@pku.edu.cn"
    sec.first_page_footer.paragraphs[0].add_run("基金项目：国家自科基金（72100001）")
    sec.even_page_footer.paragraphs[0].add_run("李四，北京大学光华管理学院，zs2@pku.edu.cn")
    doc.save(str(tmp_path / "e.docx"))

    out, removed = run_anonymize(tmp_path, tmp_path / "e.docx")

    stories = story_texts(out)
    joined = " ".join(t for _, t in stories)
    assert "zs@pku.edu.cn" not in joined, "页眉身份没清"
    assert "zs2@pku.edu.cn" not in joined, "默认/偶数页脚身份没清"
    assert "光华管理学院" not in joined, "偶数页脚身份没清"
    assert "通讯作者" not in joined, "页脚内表格身份没清"
    assert any("工作论文（请勿引用）" in t for _, t in stories), "中性页眉被误删"
    assert any("页眉" in r and "北京大学经济学院" in r for r in removed)
    assert any("首页脚注" in r and "基金项目" in r for r in removed)


def test_comments_deleted_with_marks_and_reported(tmp_path):
    """F05:批注此前只改作者名，正文（常带姓名/联系方式）原样留着。"""
    doc = build_docx(tmp_path / "f.docx")
    p = doc.paragraphs[3]
    p._p.append(parse_xml(f'<w:commentRangeStart {nsdecls("w")} w:id="1"/>'))
    p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:commentReference w:id="1"/></w:r>'))
    p._p.append(parse_xml(f'<w:commentRangeEnd {nsdecls("w")} w:id="1"/>'))
    doc.save(str(tmp_path / "f.docx"))
    comments = (f'<w:comments xmlns:w="{W_NS}">'
                '<w:comment w:id="1" w:author="张三" w:initials="张" '
                'w:date="2026-01-01T00:00:00Z">'
                '<w:p><w:r><w:t>数据找张三要（北京大学经院，zs@pku.edu.cn）</w:t></w:r></w:p>'
                '</w:comment></w:comments>')
    add_part(tmp_path / "f.docx", "word/comments.xml", comments, "comments", "rIdFixC")

    out, removed = run_anonymize(tmp_path, tmp_path / "f.docx")

    assert part_texts(out, "word/comments.xml", "comment") == [], "批注正文没删"
    assert not body_has_reference(out, "commentReference"), "正文批注标记没摘掉"
    assert any("批注" in r and "zs@pku.edu.cn" in r for r in removed), \
        f"删除没进报告：{removed}"


# ---------------------------------------------------------------- 交付闸门（docx_meta）


def test_gate_catches_comment_and_story_residue(tmp_path):
    """F05:闸门要能查到批注正文与页眉页脚联系方式，否则清理漏网照常放行。"""
    path = tmp_path / "leaky.docx"
    doc = build_docx(path)
    doc.sections[0].header.paragraphs[0].add_run("联系邮箱 zs@pku.edu.cn")
    doc.save(str(path))
    comments = (f'<w:comments xmlns:w="{W_NS}">'
                '<w:comment w:id="1" w:author="张三">'
                '<w:p><w:r><w:t>作者是张三</w:t></w:r></w:p>'
                '</w:comment></w:comments>')
    add_part(path, "word/comments.xml", comments, "comments", "rIdFixC")

    codes = {f["code"] for f in docx_meta.scan_docx(path, strict_people=True)}
    assert "comment_body" in codes, "批注残留没被发现"
    assert "part_contact" in codes, "页眉联系方式没被发现"
    with pytest.raises(docx_meta.MetaLeakError) as err:
        docx_meta.assert_clean(path, strict_people=True)
    assert "批注" in str(err.value), f"错误提示要说清问题在哪：{err.value}"

    # 署名稿闸门（非严格）口径不变：仍只查文档属性，批注/页眉不算它的账
    lenient = {f["code"] for f in docx_meta.scan_docx(path)}
    assert "comment_body" not in lenient and "part_contact" not in lenient


def test_anonymized_output_passes_strict_gate(tmp_path):
    """端到端：页眉带联系方式、包里带批注的稿子，匿名稿必须过 strict 闸门。"""
    doc = build_docx(tmp_path / "g.docx")
    doc.sections[0].header.paragraphs[0].add_run("通讯作者：张三，zs@pku.edu.cn")
    doc.paragraphs[3]._p.append(
        parse_xml(f'<w:r {nsdecls("w")}><w:commentReference w:id="1"/></w:r>'))
    doc.save(str(tmp_path / "g.docx"))
    comments = (f'<w:comments xmlns:w="{W_NS}">'
                '<w:comment w:id="1" w:author="张三">'
                '<w:p><w:r><w:t>作者：张三（zs@pku.edu.cn）</w:t></w:r></w:p>'
                '</w:comment></w:comments>')
    add_part(tmp_path / "g.docx", "word/comments.xml", comments, "comments", "rIdFixC")
    before = docx_meta.scan_docx(tmp_path / "g.docx", strict_people=True)
    assert {f["code"] for f in before} >= {"comment_body", "part_contact"}, \
        "夹具没注入成功"

    out, _ = run_anonymize(tmp_path, tmp_path / "g.docx")

    assert docx_meta.scan_docx(out, strict_people=True) == [], \
        "匿名稿没通过 strict 闸门"


def test_people_attrs_scanned_outside_word_parts(tmp_path):
    """F05:作者属性藏在 customXml 里时，旧扫描只看 word/ 目录。"""
    path = tmp_path / "custom.docx"
    build_docx(path)
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    items["customXml/item1.xml"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<data xmlns="http://example.com/x" author="zs@pku.edu.cn"/>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)

    findings = docx_meta.scan_docx(path, strict_people=True)
    assert any(f["part"] == "customXml/item1.xml" for f in findings), \
        "customXml 里的作者属性没被发现"
    docx_meta.sanitize_docx(path, scrub_people=True)
    assert docx_meta.scan_docx(path, strict_people=True) == [], "清完仍有残留"


def test_parse_part_xml_rejects_dtd_and_entities():
    """解析不受信部件：拒绝 DTD/实体，不启用外部实体。"""
    evil = (b'<?xml version="1.0"?>'
            b'<!DOCTYPE r [<!ENTITY leak SYSTEM "file:///C:/secret.txt">]>'
            b'<r>&leak;</r>')
    with pytest.raises(etree.XMLSyntaxError):
        docx_meta.parse_part_xml(evil)
    assert docx_meta.parse_part_xml(
        f'<w:r xmlns:w="{W_NS}"/>'.encode("utf-8")) is not None
