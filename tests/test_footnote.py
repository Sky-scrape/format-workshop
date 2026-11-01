"""脚注：正文套版碰不到它。

脚注部件独立于 document.xml，python-docx 的对象模型里也没有它的位置。
于是出过一个很隐蔽的问题：正文按规则集排好了，脚注还是稿件原来的字号——
成品正文合规、脚注不合规，报告却写着"已按规则集处理"。

python-docx 不能创建脚注，这里用一次 zip 手术拼出来。
"""

import json
import zipfile

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from conftest import RULES, TARGET_JOURNAL, render_into, run_tool

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FOOTNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
FOOTNOTES_CT = ("application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.footnotes+xml")


def build_docx_with_footnote(path, footnote_text="本文数据来自 Wind 数据库。", size_pt=12):
    """拼一份带真实脚注的 docx。

    步骤：先让 python-docx 生成本体骨架，再往 zip 里补 footnotes.xml、
    关系项与内容类型，最后在正文插 footnoteReference。
    """
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    p = doc.add_paragraph("本文的样本涵盖 2011-2020 年 A 股上市公司，"
                          "数据来源见注，处理方式与既有文献保持一致。" * 2)
    p.add_run()._r.append(parse_xml(f'<w:footnoteReference {nsdecls("w")} w:id="1"/>'))
    doc.add_paragraph("正文继续，用于确认脚注引用不会中断段落。")

    entry = path.with_suffix(".tmp.docx")
    doc.save(str(entry))

    half = int(round(size_pt * 2))
    footnotes = (
        f'<w:footnotes {nsdecls("w")}>'
        f'<w:footnote w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        f'<w:footnote w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        f'<w:footnote w:id="1"><w:p><w:r><w:rPr><w:sz w:val="{half}"/></w:rPr>'
        f'<w:t>{footnote_text}</w:t></w:r></w:p></w:footnote>'
        f'</w:footnotes>'
    )

    with zipfile.ZipFile(entry) as src:
        items = {n: src.read(n) for n in src.namelist()}

    ct = items["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>",
                    f'<Override PartName="/word/footnotes.xml" ContentType="{FOOTNOTES_CT}"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")

    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="rIdFn1" Type="{FOOTNOTES_REL}" '
                        f'Target="footnotes.xml"/></Relationships>')
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")

    items["word/footnotes.xml"] = footnotes.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for n, blob in items.items():
            out.writestr(n, blob)
    entry.unlink(missing_ok=True)
    return path


def read_footnote_sizes(docx_path):
    """读出脚注正文里所有字号（半磅值）。"""
    with zipfile.ZipFile(docx_path) as z:
        if "word/footnotes.xml" not in z.namelist():
            return []
        xml = z.read("word/footnotes.xml").decode("utf-8")
    from lxml import etree
    root = etree.fromstring(xml.encode("utf-8"))
    sizes = []
    for fn in root.findall(qn("w:footnote")):
        if int(fn.get(qn("w:id")) or "0") < 1:
            continue
        for el in fn.iter(qn("w:sz")):
            sizes.append(int(el.get(qn("w:val"))))
    return sizes


def test_footnote_font_follows_ruleset(tmp_path):
    """脚注字号必须按规则集改，而不是保留稿件原样。"""
    src = build_docx_with_footnote(tmp_path / "fn.docx", size_pt=12)
    assert read_footnote_sizes(src) == [24], "夹具本身没造出 12pt 的脚注"

    target_size = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8")
                             )["typography"]["footnote"]["size_pt"]

    result = render_into(tmp_path / "run", src, target="tjyj")
    sizes = read_footnote_sizes(result.docx)
    assert sizes, "成品里脚注不见了"
    assert set(sizes) == {int(round(target_size * 2))}, (
        f"脚注字号没跟着规则集变：{sizes}，期望 {target_size}pt")


def test_footnote_restyling_is_audited(tmp_path):
    """脚注套版要进报告，否则用户不知道这里动过。"""
    src = build_docx_with_footnote(tmp_path / "fn2.docx")
    result = render_into(tmp_path / "run2", src, target="tjyj")

    items = [it for it in result.items if it["action"] == "脚注正文套版"]
    assert items, f"没有脚注套版记录，实际动作：{sorted({i['action'] for i in result.items})}"
    assert items[0]["rule"] == "typography.footnote"
    assert "条脚注" in items[0]["note"]


def test_footnote_reference_in_body_is_kept(tmp_path):
    """正文里的脚注引用标记不能消失——否则脚注成了孤儿。"""
    src = build_docx_with_footnote(tmp_path / "fn3.docx")
    result = render_into(tmp_path / "run3", src, target="tjyj")

    with zipfile.ZipFile(result.docx) as z:
        body = z.read("word/document.xml").decode("utf-8")
    assert "<w:footnoteReference" in body, "正文里的脚注引用被删掉了"


def test_footnote_text_is_not_altered(tmp_path):
    """脚注里的文字一个字都不能改。"""
    text = "本文数据来自 Wind 数据库，经作者整理。"
    src = build_docx_with_footnote(tmp_path / "fn4.docx", footnote_text=text)
    result = render_into(tmp_path / "run4", src, target="tjyj")

    with zipfile.ZipFile(result.docx) as z:
        xml = z.read("word/footnotes.xml").decode("utf-8")
    assert text in xml, "脚注文字被改动了"


def test_journal_without_footnote_rules_is_untouched(tmp_path):
    """规则集没写脚注规格时，不该硬塞一个默认字号。"""
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    rules["typography"].pop("footnote", None)
    alt = tmp_path / "no_footnote.json"
    alt.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")

    src = build_docx_with_footnote(tmp_path / "fn5.docx", size_pt=12)
    result = render_into(tmp_path / "run5", src, rules_path=alt)

    assert read_footnote_sizes(result.docx) == [24], "没有脚注规则却改了字号"
    actions = {it["action"] for it in result.items}
    assert "脚注正文套版" not in actions, "没有脚注规则却记了一条改动"


def test_already_compliant_footnote_is_not_reported(tmp_path):
    """脚注本来就合规时，报告里不该出现「改为 X」。

    否则一份什么都没变的稿子，报告里却写着「N 条脚注改为 7.5pt」，
    等于自己给自己编了一条改动。
    """
    target = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8")
                        )["typography"]["footnote"]
    size = target["size_pt"]
    cn = target["font_cn"]
    latin = target.get("font_latin") or "Times New Roman"

    # 造一条字号、中西文字体都已合规的脚注：此时不该再有任何改动记录
    src = build_docx_with_footnote(tmp_path / "fn6.docx", size_pt=size)
    _patch_footnote_font(src, cn, latin)

    result = render_into(tmp_path / "run6", src, target="tjyj")
    items = [it for it in result.items if it["action"] == "脚注正文套版"]
    assert not items, f"脚注已合规却报了改动：{[i['note'] for i in items]}"


def _patch_footnote_font(docx_path, cn, latin):
    """把脚注正文的字体改成指定值，用于构造「已合规」的样本。"""
    with zipfile.ZipFile(docx_path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    xml = items["word/footnotes.xml"].decode("utf-8")
    fonts = f'<w:rFonts w:eastAsia="{cn}" w:ascii="{latin}" w:hAnsi="{latin}"/>'
    xml = xml.replace("<w:sz ", f"{fonts}<w:sz ", 1)
    items["word/footnotes.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as out:
        for n, blob in items.items():
            out.writestr(n, blob)
