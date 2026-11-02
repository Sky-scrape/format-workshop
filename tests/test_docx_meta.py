"""DOCX 元数据：正文匿名了，属性里还写着作者，等于没匿名。

Word 的“文档属性”和 WPS 的“账号记录”不在正文里，也不在审计报告的改动列表里，
但会跟着文件一起发给编辑部。这里守住两件事：
  1. 渲染产物统一清掉作者、公司、WPS 账号；
  2. 匿名稿连修订/批注作者一起清，并能在交付前检查出来。
"""

import base64
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import docx_meta  # noqa: E402

from conftest import render_into  # noqa: E402

CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"


def _wps_record():
    payload = json.dumps({"hdid": "fixture-template", "userId": "test-user-0001"},
                         separators=(",", ":")).encode()
    return base64.b64encode(payload).decode()


def poison_metadata(path, with_revision=False):
    """给普通 DOCX 装上作者、公司、WPS 账号与可选的修订作者。"""
    path = Path(path)
    if not path.exists():
        doc = Document()
        doc.add_paragraph("数字金融与企业融资约束")
        doc.add_paragraph("内容摘要：本文考察数字金融对企业融资约束的影响。")
        doc.add_paragraph("关键词：数字金融；融资约束")
        doc.add_paragraph("正文段落，用于确认清元数据不会动内容。")
        doc.save(str(path))

    doc = Document(str(path))
    cp = doc.core_properties
    cp.author = "张三"
    cp.last_modified_by = "李四"
    cp.title = "未发表论文的内部标题"
    cp.subject = "作者私有选题"
    cp.keywords = "内部；未公开"
    cp.category = "私人"
    cp.comments = "联系邮箱 private@example.org"

    if with_revision:
        p = doc.add_paragraph("修订段落")
        p._p.append(parse_xml(
            f'<w:ins {nsdecls("w")} w:id="99" w:author="王五" '
            f'w:date="2026-01-01T00:00:00Z"><w:r><w:t>新增内容</w:t></w:r></w:ins>'))

    doc.save(str(path))

    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}

    custom = (
        f'<Properties xmlns="{CUSTOM_NS}" '
        f'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" '
        f'name="KSOTemplateDocerSaveRecord"><vt:lpwstr>{_wps_record()}</vt:lpwstr></property>'
        f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="3" '
        f'name="OwnerEmail"><vt:lpwstr>private@example.org</vt:lpwstr></property>'
        f'</Properties>'
    )
    items["docProps/custom.xml"] = custom.encode("utf-8")

    ct = items["[Content_Types].xml"].decode("utf-8")
    if "/docProps/custom.xml" not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="/docProps/custom.xml" '
                                    f'ContentType="{CUSTOM_CT}"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")

    rels = items["_rels/.rels"].decode("utf-8")
    if CUSTOM_REL not in rels:
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rIdCustomMeta" Type="{CUSTOM_REL}" '
            f'Target="docProps/custom.xml"/></Relationships>')
    items["_rels/.rels"] = rels.encode("utf-8")

    app = items["docProps/app.xml"].decode("utf-8")
    for tag, value in (("Company", "示例公司"), ("Manager", "李四")):
        pat = rf"<{tag}\s*/>|<{tag}>.*?</{tag}>"
        repl = f"<{tag}>{value}</{tag}>"
        if re.search(pat, app, re.S):
            app = re.sub(pat, repl, app, count=1, flags=re.S)
        else:
            app = app.replace("</Properties>", f"{repl}</Properties>")
    items["docProps/app.xml"] = app.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items.items():
            z.writestr(name, data)
    return path


def test_scan_finds_and_sanitize_removes_metadata(tmp_path):
    path = poison_metadata(tmp_path / "private.docx", with_revision=True)

    findings = docx_meta.scan_docx(path)
    codes = {f["code"] for f in findings}
    assert {"core_property", "app_property", "custom_property"} <= codes
    names = {f["field"] for f in findings}
    assert "KSOTemplateDocerSaveRecord" in names
    assert "people_author" in {f["code"] for f in docx_meta.scan_docx(path, strict_people=True)}

    before = "\n".join(p.text for p in Document(str(path)).paragraphs)
    removed = docx_meta.sanitize_docx(path, scrub_people=True)
    after = "\n".join(p.text for p in Document(str(path)).paragraphs)

    assert removed, "有隐私属性却没报告清理项"
    assert after == before, "清元数据不能动正文"
    assert docx_meta.scan_docx(path, strict_people=True) == []


def test_sanitize_is_idempotent(tmp_path):
    path = poison_metadata(tmp_path / "once.docx")
    assert docx_meta.sanitize_docx(path)
    assert docx_meta.sanitize_docx(path) == []
    assert docx_meta.scan_docx(path, strict_people=True) == []


def test_gate_raises_without_leaking_values(tmp_path):
    path = poison_metadata(tmp_path / "gate.docx")
    with pytest.raises(docx_meta.MetaLeakError) as err:
        docx_meta.assert_clean(path)
    msg = str(err.value)
    assert "张三" not in msg and "private@example.org" not in msg
    assert "WPS" in msg, "错误提示要让人知道问题在哪"


def test_app_gate_blocks_dirty_docx(tmp_path):
    """网页交付闸门要把元数据残留翻译成用户看得懂的失败。"""
    import app

    path = poison_metadata(tmp_path / "app-gate.docx")
    with pytest.raises(app.ConvertError) as err:
        app._check_meta(path)
    assert "隐私检查" in err.value.title

    docx_meta.sanitize_docx(path, scrub_people=True)
    assert app._check_meta(path, strict_people=True) is True


def test_renderer_cleans_source_metadata(sample_docx, tmp_path):
    """端到端：带 WPS 账号的源稿渲染后，成品属性必须是干净的。"""
    src = tmp_path / "poisoned.docx"
    shutil.copy(sample_docx, src)
    poison_metadata(src)
    assert docx_meta.scan_docx(src), "夹具没有成功注入隐私属性"

    result = render_into(tmp_path / "run", src, target="tjyj")
    assert docx_meta.scan_docx(result.docx) == [], "渲染成品仍带作者/WPS 元数据"
