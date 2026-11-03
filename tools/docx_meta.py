"""DOCX 元数据隐私：清理与交付闸门。

正文里的作者信息由 anonymize.py 处理；这里管另一条更容易漏掉的通道：
`docProps/` 里的作者、最后修改者、公司，以及 WPS 写进去的账号标识。
这些值在 Word 界面上看不见，但会随文件一起发给编辑部。

清理是“删属性”，不是“另存新文档”：保留原 ZIP 的全部部件与正文，
只改 `docProps/core.xml`、`docProps/app.xml`、`docProps/custom.xml`。
`--people` 额外把修订与批注的作者改成“匿名”，用于匿名稿。

`check --strict-people`（匿名稿交付闸门）在属性之外还查三处：
修订/批注作者、批注部件的正文（只改作者名挡不住正文里的姓名与联系方式）、
页眉页脚正文里的邮箱/电话等联系方式。

用法：
    python tools/docx_meta.py sanitize 成品.docx [--people]
    python tools/docx_meta.py check 成品.docx [--strict-people]
"""

import argparse
import json
import os
import re
import tempfile
import time
import zipfile
from pathlib import Path

from lxml import etree

CORE_PART = "docProps/core.xml"
APP_PART = "docProps/app.xml"
CUSTOM_PART = "docProps/custom.xml"
COMMENTS_PART = "word/comments.xml"
# 页眉/页脚部件：word/header1.xml、word/footer2.xml 这类命名
STORY_PART_RE = re.compile(r"^word/(?:header|footer)\d*\.xml$")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 这些字段描述的是“谁、用什么模板、什么主题做的”，跟正文无关。
# 标题也在内：题名常常含姓名或未发表选题，没必要随文件外发。
CORE_PRIVATE = {
    "title", "subject", "creator", "keywords", "description", "lastModifiedBy",
    "revision", "category", "contentStatus", "identifier", "language",
    "version", "lastPrinted",
}
APP_PRIVATE = {"Company", "Manager", "Template", "HyperlinkBase"}

ANON = "匿名"
ANON_INITIALS = "匿"

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
ID_RE = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
ORCID_RE = re.compile(r"(?<!\d)\d{4}-\d{4}-\d{4}-\d{3}[\dXx](?!\d)")

LABELS = {
    "core_property": "文档属性里的作者/主题等描述信息",
    "app_property": "应用属性里的公司/管理者/模板信息",
    "custom_property": "自定义属性（含 WPS 账号标识）",
    "core_contact": "文档属性里的邮箱/电话/证件号",
    "people_author": "修订或批注的作者名",
    "people_initials": "修订或批注的作者缩写",
    "comment_body": "批注里的正文或作者",
    "part_contact": "页眉页脚里的联系方式",
}


def parse_part_xml(data):
    """解析 DOCX 部件里的 XML。部件来自不受信的上传文件：
    带 DTD/实体的部件一律拒绝（OOXML 部件本来就不该有），
    解析器也不启用外部实体。"""
    head = data[:2048]
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise etree.XMLSyntaxError("OOXML 部件不允许 DTD/ENTITY", 0, 0, 0)
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    return etree.fromstring(data, parser)


class MetaLeakError(RuntimeError):
    """交付文件仍带可识别个人身份或账号的元数据。"""

    def __init__(self, findings, path=""):
        self.findings = list(findings)
        where = f"{Path(path).name}：" if path else ""
        kinds = "；".join(sorted({LABELS.get(f["code"], f["code"]) for f in self.findings}))
        super().__init__(f"{where}元数据仍有隐私残留（{kinds}）")


def _localname(el):
    return etree.QName(el).localname


def _xml_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _text(el):
    return "".join(el.itertext()).strip()


def _attr_local(attr):
    return etree.QName(attr).localname


def _has_contact(text):
    return any(p.search(text or "") for p in (EMAIL_RE, PHONE_RE, ID_RE, ORCID_RE))


def _scan_comments(blob, part):
    """批注部件：正文或作者未清即算残留。只把作者改成「匿名」时，
    批注正文（常带姓名、单位、联系方式）仍会原样发出去。"""
    try:
        root = parse_part_xml(blob)
    except etree.XMLSyntaxError:
        return []
    findings = []
    for c in root.iter(f"{{{W_NS}}}comment"):
        author = ""
        for attr, value in c.attrib.items():
            if _attr_local(attr) == "author":
                author = value.strip()
        if _text(c) or author not in ("", ANON):
            findings.append({"code": "comment_body", "part": part,
                             "field": author or "正文"})
    return findings


def _scan_story_contact(blob, part):
    """页眉/页脚部件正文（w:t）里的邮箱/电话等联系方式。
    属性扫描覆盖不到部件正文，清理漏网时由闸门拦下。"""
    try:
        root = parse_part_xml(blob)
    except etree.XMLSyntaxError:
        return []
    if _has_contact("".join(root.itertext())):
        return [{"code": "part_contact", "part": part, "field": "text"}]
    return []


def scan_docx(path, strict_people=False):
    """返回元数据隐私发现；不修改文件。

    strict_people=False 只查文档属性；True 额外查修订/批注作者、
    批注正文与页眉页脚联系方式，匿名稿交付必须用 True。
    """
    findings = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        if CORE_PART in names:
            root = parse_part_xml(z.read(CORE_PART))
            for el in root:
                field = _localname(el)
                value = _text(el)
                if value and field in CORE_PRIVATE:
                    findings.append({"code": "core_property", "part": CORE_PART, "field": field})
                if value and _has_contact(value):
                    findings.append({"code": "core_contact", "part": CORE_PART, "field": field})

        if APP_PART in names:
            root = parse_part_xml(z.read(APP_PART))
            for el in root:
                field = _localname(el)
                if _text(el) and field in APP_PRIVATE:
                    findings.append({"code": "app_property", "part": APP_PART, "field": field})

        if CUSTOM_PART in names:
            root = parse_part_xml(z.read(CUSTOM_PART))
            for el in root:
                if _localname(el) != "property":
                    continue
                field = el.get("name") or "property"
                findings.append({"code": "custom_property", "part": CUSTOM_PART, "field": field})

        if strict_people:
            for part in sorted(names):
                if not part.endswith(".xml"):
                    continue
                if part == COMMENTS_PART:
                    findings.extend(_scan_comments(z.read(part), part))
                    continue
                if STORY_PART_RE.match(part):
                    findings.extend(_scan_story_contact(z.read(part), part))
                try:
                    root = parse_part_xml(z.read(part))
                except etree.XMLSyntaxError:
                    continue
                for el in root.iter():
                    for attr, value in el.attrib.items():
                        local = _attr_local(attr)
                        if local == "author" and value.strip() not in ("", ANON):
                            findings.append({"code": "people_author", "part": part, "field": local})
                            break
                        if local == "initials" and value.strip() not in ("", ANON_INITIALS):
                            findings.append({"code": "people_initials", "part": part, "field": local})
                            break
    return findings


def assert_clean(path, strict_people=False):
    findings = scan_docx(path, strict_people=strict_people)
    if findings:
        raise MetaLeakError(findings, path)
    return True


def _sanitize_core(blob, removed):
    root = parse_part_xml(blob)
    changed = False
    for el in list(root):
        field = _localname(el)
        if field not in CORE_PRIVATE:
            continue
        if _text(el):
            removed.append(f"core:{field}")
        root.remove(el)
        changed = True
    return _xml_bytes(root) if changed else blob


def _sanitize_app(blob, removed):
    root = parse_part_xml(blob)
    changed = False
    for el in root:
        field = _localname(el)
        if field not in APP_PRIVATE:
            continue
        if _text(el):
            removed.append(f"app:{field}")
        if el.text:
            el.text = ""
        for child in list(el):
            el.remove(child)
        changed = True
    return _xml_bytes(root) if changed else blob


def _sanitize_custom(blob, removed):
    root = parse_part_xml(blob)
    changed = False
    for el in list(root):
        if _localname(el) != "property":
            continue
        removed.append(f"custom:{el.get('name') or 'property'}")
        root.remove(el)
        changed = True
    return _xml_bytes(root) if changed else blob


def _sanitize_people(blob, part, removed):
    """把修订与批注的作者改成固定匿名值，而不是删属性破坏 OOXML。"""
    try:
        root = parse_part_xml(blob)
    except etree.XMLSyntaxError:
        return blob
    changed = False
    for el in root.iter():
        for attr, value in list(el.attrib.items()):
            local = _attr_local(attr)
            if local == "author" and value.strip() not in ("", ANON):
                el.set(attr, ANON)
                changed = True
            elif local == "initials" and value.strip() not in ("", ANON_INITIALS):
                el.set(attr, ANON_INITIALS)
                changed = True
    if changed:
        removed.append(f"people:{part}")
        return _xml_bytes(root)
    return blob


def sanitize_docx(path, scrub_people=False):
    """就地清理元数据，保留全部正文与版式。返回被清理的字段名。"""
    path = Path(path)
    removed = []
    fd, tmp_name = tempfile.mkstemp(prefix=path.stem + "-", suffix=".docx",
                                    dir=str(path.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == CORE_PART:
                    data = _sanitize_core(data, removed)
                elif info.filename == APP_PART:
                    data = _sanitize_app(data, removed)
                elif info.filename == CUSTOM_PART:
                    data = _sanitize_custom(data, removed)
                elif scrub_people and info.filename.endswith(".xml"):
                    # 属性扫描覆盖所有 .xml 部件（含 customXml），清理口径必须一致
                    data = _sanitize_people(data, info.filename, removed)
                dst.writestr(info, data)
        # Windows 瞬时文件占用（杀毒/索引/预览句柄，WinError 5）会让 replace
        # 偶发 PermissionError——这不是真失败，稍候重试即可（FX1）。同款姿势见
        # app.py details.json 的原子写：4 次退避共约 0.24s，最后一次直接抛出。
        for attempt in range(5):
            try:
                os.replace(tmp, path)
                break
            except PermissionError:
                if attempt == 4:
                    raise
                time.sleep((0.02, 0.04, 0.08, 0.1, 0.2)[attempt])
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
    return removed


def main():
    ap = argparse.ArgumentParser(description="DOCX 元数据清理与检查")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_san = sub.add_parser("sanitize", help="清理文档属性与 WPS 账号标识")
    p_san.add_argument("docx")
    p_san.add_argument("--people", action="store_true", help="同时匿名化修订/批注作者")

    p_chk = sub.add_parser("check", help="检查元数据隐私，命中时退出码为 1")
    p_chk.add_argument("docx")
    p_chk.add_argument("--strict-people", action="store_true",
                       help="同时检查修订/批注作者")

    args = ap.parse_args()
    if args.cmd == "sanitize":
        removed = sanitize_docx(args.docx, scrub_people=args.people)
        print(json.dumps({"removed": removed}, ensure_ascii=False))
        return 0

    findings = scan_docx(args.docx, strict_people=args.strict_people)
    print(json.dumps({"findings": findings}, ensure_ascii=False, indent=2))
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
