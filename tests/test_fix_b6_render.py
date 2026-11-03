"""第六批渲染收尾的回归钉：F13DEEP/JRYJNOTE/EXTFOOT/SDTNEST/SETRUN0/FIGACT。

每个用例对应问题单里的证据复现：

- F13DEEP  必填键校验只查键存在性：line_spacing={"mode":"multiple"} 缺 value
           子键照样渲染中途裸 KeyError。现在 REQUIRED_RULE_KEYS 带深一层断言
           （dict 形态 / 条件子键 / list 条目子键），缺什么一次报全
- JRYJNOTE  jryj 声明 cn-first-by-alphabet-then-en-by-alphabet，note 描述的
           排序没有发生（中文实际按码位序，「周一」排在「张三」前）。现在
           中文「字母序」用 GB 序实现、英文用 str.lower 键序，SORT_NOTE 与
           实现一致
- EXTFOOT  TargetMode="External" 的 footnotes 关系访问 target_part 抛未捕获
           ValueError。现在跳过外部关系并在汇总里说明 N 条未处理
- SDTNEST  sdt 只下钻一层：外层 sdt 已有直接段落时，内嵌 sdt 的段落静默
           不进 IR。现在递归展开直到取尽段落/表格
- SETRUN0  set_run 用 if letter_spacing: 判断，声明 0pt 时残留旧字间距。
           现在 None 才不动，数值（含 0）都生效
- FIGACT   图片段动作名固定为「图片段解除固定行距」：仅段距不合规时命名
           不实。现在仅确有行距差异才用现名，否则「图片段套版」，note 按
           fmt_delta 列实际差异

夹具姿势与 test_fix_render_core.py / test_citation_convert.py 相同：
python-docx 造稿、子进程跑 parse/render（参数列表直传，不经 shell）。
"""

import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

from conftest import RULES, TOOLS, run_tool

import sys
sys.path.insert(0, str(TOOLS))
import render_from_ir as rf  # noqa: E402

NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FOOTNOTES_REL = ("http://schemas.openxmlformats.org/officeDocument"
                 "/2006/relationships/footnotes")

TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
    "en_journal": "{authors}，{year}，“{title}”，{journal}{volume_part}{page_part}.",
    "en_book": "{authors}，{year}，“{title}”，{publisher}.",
    "en_other": "{authors}，{year}，“{title}”，{source}.",
}

REFS_4 = [
    "[1] 周一. 融资约束的度量[J]. 金融研究, 2021, 76(2): 20-40.",
    "[2] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
    "[3] Lee M. Firm investment and financing[J]. Journal of Finance, 2018, 1(1): 1-9.",
    "[4] Chen Y. Digital finance and firms[J]. Journal of Banking, 2019, 2(2): 10-30.",
]


# ---------------------------------------------------------------- 夹具

def load_rules(tmp_path, journal="tjyj", name=None, mutate=None):
    """读仓库自带规则集的副本；mutate(rules) 可在落盘前改字段。"""
    rules = json.loads((RULES / f"{journal}.json").read_text(encoding="utf-8"))
    if mutate is not None:
        mutate(rules)
    p = tmp_path / (name or f"{journal}.json")
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return p


def make_src(tmp_path, name="src", refs=(), cite_para=None):
    """最小可渲染稿件；refs 逐条挂进文献表，cite_para 是正文引用句。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph(cite_para or "正文第一段，引用了相关文献。")
    doc.add_paragraph("参考文献")
    for t in refs:
        doc.add_paragraph(t)
    p = tmp_path / f"{name}.docx"
    doc.save(str(p))
    return p


def parse(tmp_path, docx):
    ir = docx.with_suffix(".ir.json")
    run_tool("parse_docx.py", docx, "-o", ir)
    return ir


def render(tmp_path, ir, rules_path, citation=False):
    out = Path(ir).with_suffix(".out.docx")
    args = [ir, rules_path, "-o", out]
    if citation:
        args.append("--citation")
    run_tool("render_from_ir.py", *args)
    return out


def audit_items(out, action=None):
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    items = audit["items"]
    return [i for i in items if i["action"] == action] if action else items


def paragraph_with_spacing(doc, text, twips):
    """带 w:spacing（字间距，twips）run 的段落。"""
    p = doc.add_paragraph(text)
    rpr = p.runs[0]._element.get_or_add_rPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), str(twips))
    rpr.append(sp)
    return p


def run_spacing_vals(p):
    """段落各 run 的 w:spacing 值（无则空列表）。"""
    out = []
    for r in p.runs:
        rpr = r._element.find(qn("w:rPr"))
        sp = rpr.find(qn("w:spacing")) if rpr is not None else None
        out.append(sp.get(qn("w:val")) if sp is not None else None)
    return out


def add_external_footnote_rel(docx_path):
    """给 docx 补一条 TargetMode="External" 的 footnotes 关系（EXTFOOT 复现）。

    外部关系不指向包内部件，无需 Content_Types 条目——这正是 target_part
    无定义、访问即抛 ValueError 的那种关系。
    """
    with zipfile.ZipFile(str(docx_path)) as src:
        items = {n: src.read(n) for n in src.namelist()}
    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rIdExtFn" Type="{FOOTNOTES_REL}" '
        f'Target="https://example.org/footnotes.xml" TargetMode="External"/>'
        f"</Relationships>")
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    with zipfile.ZipFile(str(docx_path), "w", zipfile.ZIP_DEFLATED) as out:
        for n, blob in items.items():
            out.writestr(n, blob)
    return docx_path


def sdt_with(paras):
    """w:sdt，sdtContent 下依次挂给定文字的段落。"""
    sdt = parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>")
    content = sdt.find(qn("w:sdtContent"))
    for text in paras:
        content.append(parse_xml(
            f'<w:p {nsdecls("w")}><w:r><w:t>{text}</w:t></w:r></w:p>'))
    return sdt


def docx_xml_text(docx_path):
    """成品 word/document.xml 的全部文字（含 sdt 内段落，.paragraphs 取不到）。

    解析的是渲染器自己刚产出的 XML，仍按仓库口径先拒 DOCTYPE/ENTITY
    （同 conftest.layout_signature 的姿势）。
    """
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(str(docx_path)) as z:
        raw = z.read("word/document.xml")
    if b"<!DOCTYPE" in raw or b"<!ENTITY" in raw:
        raise ValueError("word/document.xml 带有 DOCTYPE/ENTITY 声明，拒绝解析")
    root = ET.fromstring(raw)
    return "".join(t.text or "" for t in root.iter() if t.tag.endswith("}t"))


# ---------------------------------------------------------------- F13DEEP

def test_missing_line_spacing_value_fails_clean(tmp_path):
    """证据复现：mode=multiple 缺 value 子键 → 入口干净报错，不裸 KeyError。"""
    src = make_src(tmp_path)

    def mutate(rules):
        rules["typography"]["body"]["line_spacing"] = {"mode": "multiple"}
    rp = load_rules(tmp_path, name="no_value.json", mutate=mutate)

    ir = parse(tmp_path, src)
    import subprocess
    r = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir), str(rp),
         "-o", str(Path(ir).with_suffix(".out.docx"))],
        capture_output=True, text=True)
    out = Path(ir).with_suffix(".out.docx")

    assert r.returncode == 1, f"应干净退出，实际 exit {r.returncode}"
    assert "[RULESET-INCOMPLETE] 规则集 tjyj 缺少字段 " \
        "typography.body.line_spacing.value" in r.stderr, r.stderr[-500:]
    assert "（章节 版式）" in r.stderr, "错误信息保持「章节」说明风格"
    assert "Traceback" not in r.stderr and "KeyError" not in r.stderr
    assert not out.exists(), "校验失败不应产出成品文件"


def test_deep_subkeys_reported_in_one_pass(tmp_path):
    """note_layers 条目缺 kind、blocks 条目缺 type：与存在性缺口一次报全。"""
    src = make_src(tmp_path)

    def mutate(rules):
        rules["tables"]["note_layers"] = [{"size_pt": 7.5}]
        rules["blocks"].append({"lang": "zh", "label": "内容摘要"})
    rp = load_rules(tmp_path, name="deep.json", mutate=mutate)

    ir = parse(tmp_path, src)
    import subprocess
    r = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir), str(rp),
         "-o", str(Path(ir).with_suffix(".out.docx"))],
        capture_output=True, text=True)

    assert r.returncode == 1
    assert "规则集 tjyj 缺少字段 tables.note_layers[].kind" in r.stderr, r.stderr[-500:]
    assert "规则集 tjyj 缺少字段 blocks[].type" in r.stderr, "缺什么要一次报全"
    assert "Traceback" not in r.stderr


def test_shipped_rulesets_all_pass_validation():
    """全键齐全的本地规则集全部通过校验（exact 模式无 value 不得误杀）。"""
    for j in ("cjyj", "gjjryj", "tjyj"):
        rules = json.loads((RULES / f"{j}.json").read_text(encoding="utf-8"))
        assert rf.missing_required_keys(rules) == [], f"{j} 被误报缺字段"


def test_keys_when_only_fires_for_multiple_mode():
    """条件断言只管 mode=multiple：exact/未声明模式不要求 value 子键。"""
    base = {
        "typography": {"body": {"line_spacing": {}}},
        "tables": {"note_layers": []}, "figures": {"note_layers": []},
        "blocks": [],
    }
    ls = base["typography"]["body"]["line_spacing"]
    ls["mode"], ls["value_pt"] = "exact", 15.5
    assert "typography.body.line_spacing.value" not in rf.missing_required_keys(base)
    ls["mode"] = "multiple"
    del ls["value_pt"]
    assert "typography.body.line_spacing.value" in rf.missing_required_keys(base)
    ls["value"] = 1.5
    assert "typography.body.line_spacing.value" not in rf.missing_required_keys(base)


# ---------------------------------------------------------------- JRYJNOTE

def test_alphabet_scheme_key_order_matches_note():
    """jryj scheme 的比较键：中文组按 GB 序（张三<周一），英文组 str.lower 键序。"""
    e_zhou = {"authors": ["周一"], "year": "2021", "lang": "cn", "raw": REFS_4[0]}
    e_zhang = {"authors": ["张三"], "year": "2020", "lang": "cn", "raw": REFS_4[1]}
    e_lee = {"authors": ["Lee M"], "year": "2018", "lang": "en", "raw": REFS_4[2]}
    e_chen = {"authors": ["Chen Y"], "year": "2019", "lang": "en", "raw": REFS_4[3]}
    scheme = "cn-first-by-alphabet-then-en-by-alphabet"
    keys = [rf.sort_key(e, scheme) for e in (e_zhou, e_zhang, e_lee, e_chen)]
    assert keys[0][0] == keys[1][0] == 0, "中文组在前"
    assert keys[2][0] == keys[3][0] == 1, "英文组在后"
    assert keys[1] < keys[0], "中文按 GB 序：张三(zhang) 在 周一(zhou) 前"
    assert keys[3] < keys[2], "英文按 str.lower 键序：Chen 在 Lee 前"


def test_alphabet_scheme_end_to_end_matches_note(tmp_path):
    """端到端：张三/周一/Chen/Lee 的成品顺序与 SORT_NOTE 描述一致。"""
    cite = "正文第一段，引用了文献一[1]、文献二[2]、文献三[3]和文献四[4]。"
    src = make_src(tmp_path, name="jryj", refs=REFS_4, cite_para=cite)
    ir = parse(tmp_path, src)

    def mutate(rules):
        rules["citation"]["system"] = "author-year"
        rl = rules["citation"].setdefault("reference_list", {})
        rl["templates"] = TEMPLATES
        rl["sort"] = "cn-first-by-alphabet-then-en-by-alphabet"
    rp = load_rules(tmp_path, name="jryj_rules.json", mutate=mutate)
    out = render(tmp_path, ir, rp, citation=True)

    texts = [p.text for p in Document(str(out)).paragraphs]
    pos = {}
    for i, t in enumerate(texts):
        for who in ("张三", "周一", "Chen", "Lee"):
            if t.startswith(who):
                pos[who] = i
    assert set(pos) == {"张三", "周一", "Chen", "Lee"}, f"条目缺失：{pos}"
    assert pos["张三"] < pos["周一"], "中文按 GB 序（≈拼音）"
    assert pos["Chen"] < pos["Lee"], "英文按字母序"
    assert pos["周一"] < pos["Chen"], "中文组整体在前"

    items = audit_items(out, "文献表排序")
    assert items and items[0]["rule"] == "citation.reference_list.sort"
    note = items[0]["note"]
    assert "GB 序" in note, f"note 必须写明中文按 GB 序：{note}"
    assert "字母序" in note, f"note 对应声明的「字母序」措辞：{note}"
    assert "码位序" not in note, "实现不再是码位序，note 不得再说码位序"


# ---------------------------------------------------------------- EXTFOOT

def test_external_footnote_rel_renders_and_is_disclosed(tmp_path):
    """证据复现：外部脚注关系 → 渲染不崩，汇总说明 N 条未处理。"""
    src = add_external_footnote_rel(make_src(tmp_path, name="extfn"))
    rp = load_rules(tmp_path, name="extfn_rules.json")
    ir = parse(tmp_path, src)
    stdout = run_tool("render_from_ir.py", ir, rp, "-o",
                      Path(src).with_suffix(".out.docx"))

    out = Path(src).with_suffix(".out.docx")
    assert out.exists(), "带外部脚注关系的稿子渲染不应崩"
    assert "1 条外部脚注关系未处理" in stdout, stdout[-500:]
    assert "外部超链接式脚注" in stdout


def test_external_footnote_rel_with_citation_gate(tmp_path):
    """--citation 路径（footnote_texts 体检）同样不崩：无脚注正文可报，正常转换。"""
    cite = "正文第一段，引用了文献一[1]和文献二[2]。"
    refs = REFS_4[:2]
    src = add_external_footnote_rel(
        make_src(tmp_path, name="extfn_cite", refs=refs, cite_para=cite))
    ir = parse(tmp_path, src)

    def mutate(rules):
        rules["citation"]["system"] = "author-year"
        rl = rules["citation"].setdefault("reference_list", {})
        rl["templates"] = TEMPLATES
    rp = load_rules(tmp_path, name="extfn_cite_rules.json", mutate=mutate)
    out = render(tmp_path, ir, rp, citation=True)

    joined = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "周一" in joined and "张三" in joined, "转换应正常完成"
    assert "缺少字段" not in joined


# ---------------------------------------------------------------- SDTNEST

def test_nested_sdt_paragraphs_enter_ir_and_get_typed(tmp_path):
    """证据复现：外层 sdt 含段落+内嵌 sdt（内嵌再含段落）→ 两层都进 IR。"""
    src = tmp_path / "nested.docx"
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    outer = sdt_with(["外层内容控件里的段落"])
    outer.find(qn("w:sdtContent")).append(
        sdt_with(["内嵌内容控件里的段落"]))
    doc.element.body.append(outer)
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))

    ir = json.loads(Path(parse(tmp_path, src)).read_text(encoding="utf-8"))
    texts = [b.get("text", "") for b in ir["blocks"]]
    assert any("外层内容控件里的段落" in t for t in texts), "外层段落要进 IR"
    assert any("内嵌内容控件里的段落" in t for t in texts), \
        "内嵌 sdt 的段落要递归进 IR，不能静默丢弃"
    assert not ir.get("unsupported_content_blocks"), \
        "两层段落都处理了，不该再记未处理块"
    sdt_blocks = [b for b in ir["blocks"] if "sdt_pos" in b]
    assert [b["sdt_pos"] for b in sdt_blocks] == [0, 1], \
        "槽位内序按文档序递增，渲染端按它取元素"

    rp = load_rules(tmp_path, name="nested_rules.json")
    out = render(tmp_path, src.with_suffix(".ir.json"), rp)
    ids = {i["id"] for i in audit_items(out)}
    slot = sdt_blocks[0]["idx"]
    assert f"{slot}.0|正文套版" in ids and f"{slot}.1|正文套版" in ids, \
        f"两层段落都要套版：{sorted(ids)[:8]}"
    texts_out = docx_xml_text(out)
    assert "外层内容控件里的段落" in texts_out
    assert "内嵌内容控件里的段落" in texts_out


def test_nested_empty_sdt_still_counted(tmp_path):
    """边界钉：取空的内嵌内容控件仍计数；相邻段落照常进 IR（F24 口径不变）。"""
    src = tmp_path / "nested_empty.docx"
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    outer = sdt_with(["外层段落"])
    outer.find(qn("w:sdtContent")).append(
        parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>"))
    doc.element.body.append(outer)
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))

    ir = json.loads(Path(parse(tmp_path, src)).read_text(encoding="utf-8"))
    assert ir.get("unsupported_content_blocks") == 1, "取空的内嵌控件要计数"
    assert any("外层段落" in b.get("text", "") for b in ir["blocks"])


def test_empty_sdt_still_counted_alone(tmp_path):
    """空内容控件整块无段落：照旧计 1（批五 F24 行为不回退）。"""
    src = tmp_path / "empty_sdt.docx"
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.element.body.append(
        parse_xml(f"<w:sdt {nsdecls('w')}><w:sdtPr/><w:sdtContent/></w:sdt>"))
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))

    ir = json.loads(Path(parse(tmp_path, src)).read_text(encoding="utf-8"))
    assert ir.get("unsupported_content_blocks") == 1


def test_nested_sdt_render_is_deterministic(tmp_path):
    """同一输入两次渲染结果相同（承诺③）：递归下钻不引入随机性。"""
    src = tmp_path / "det.docx"
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    outer = sdt_with(["外层内容控件里的段落"])
    outer.find(qn("w:sdtContent")).append(sdt_with(["内嵌内容控件里的段落"]))
    doc.element.body.append(outer)
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))

    from conftest import layout_signature
    ir = parse(tmp_path, src)
    rp = load_rules(tmp_path, name="det_rules.json")
    out1 = render(tmp_path, ir, rp)
    out2 = Path(src).with_suffix(".out2.docx")
    run_tool("render_from_ir.py", ir, rp, "-o", out2)
    assert layout_signature(out1) == layout_signature(out2)


# ---------------------------------------------------------------- SETRUN0

def test_zero_letter_spacing_clears_existing(tmp_path):
    """证据复现：规则集声明 0pt → 原带 w:spacing 的标题 run 渲染后无残留。"""
    src = tmp_path / "spacing0.docx"
    doc = Document()
    title = paragraph_with_spacing(doc, "数字金融发展与企业融资约束", 26)
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))
    assert run_spacing_vals(title) == ["26"], "夹具本身要带字间距"

    def mutate(rules):
        rules["typography"]["title"]["letter_spacing_pt"] = 0
    rp = load_rules(tmp_path, name="spacing0_rules.json", mutate=mutate)
    out = render(tmp_path, parse(tmp_path, src), rp)

    doc_out = Document(str(out))
    first = doc_out.paragraphs[0]
    assert first.text.startswith("数字金融"), "定位的是标题段"
    assert run_spacing_vals(first) == [None], \
        f"0pt 声明必须清掉旧字间距，实际 {run_spacing_vals(first)}"

    items = audit_items(out, "文章标题套版")
    assert items, "0pt 字间距的清除要出现在标题套版条目里"
    assert "字间距" in items[0]["note"], \
        f"字间距变化要进审计：{items[0]['note']}"


def test_undeclared_letter_spacing_leaves_run_alone(tmp_path):
    """未声明 letter_spacing_pt（None）→ 不动 run 上已有的字间距。"""
    src = tmp_path / "spacing_none.docx"
    doc = Document()
    title = paragraph_with_spacing(doc, "数字金融发展与企业融资约束", 26)
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(src))

    def mutate(rules):
        rules["typography"]["title"].pop("letter_spacing_pt", None)
    rp = load_rules(tmp_path, name="spacing_none_rules.json", mutate=mutate)
    out = render(tmp_path, parse(tmp_path, src), rp)

    doc_out = Document(str(out))
    assert run_spacing_vals(doc_out.paragraphs[0]) == ["26"], \
        "规则集未声明时不得清除/改写已有的字间距"


# ---------------------------------------------------------------- FIGACT

def _fig_doc(path, *, center, line_pt, before_after=None):
    from io import BytesIO
    import base64
    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    fig = doc.add_paragraph()
    fig.add_run().add_picture(BytesIO(png))
    if center:
        fig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if line_pt:
        fig.paragraph_format.line_spacing = Pt(line_pt)
    else:
        fig.paragraph_format.line_spacing = 1.0
    if before_after:
        fig.paragraph_format.space_before = Pt(before_after)
        fig.paragraph_format.space_after = Pt(before_after)
    doc.add_paragraph("图1 企业创新产出趋势")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REFS_4[1])
    doc.save(str(path))
    return path


def test_figure_spacing_only_gets_generic_action(tmp_path):
    """回归：仅段前/段后距不合规（行距本就单倍且居中）→「图片段套版」+实际差异。"""
    src = _fig_doc(tmp_path / "fig_sa.docx", center=True, line_pt=None,
                   before_after=6)
    rp = load_rules(tmp_path, name="fig_sa_rules.json")
    out = render(tmp_path, parse(tmp_path, src), rp)

    assert audit_items(out, "图片段解除固定行距") == [], \
        "行距本就合规，不得再叫「解除固定行距」"
    items = audit_items(out, "图片段套版")
    assert items, "仅段距不合规的图片段要有套版条目"
    note = items[0]["note"]
    assert "段前距" in note and "段后距" in note, \
        f"note 要按实际差异列项：{note}"
    assert items[0]["rule"] == "figures"


def test_figure_fixed_line_spacing_keeps_name(tmp_path):
    """确有行距差异的图片段仍叫「图片段解除固定行距」，note 说明裁切风险。"""
    src = _fig_doc(tmp_path / "fig_fix.docx", center=False, line_pt=18)
    rp = load_rules(tmp_path, name="fig_fix_rules.json")
    out = render(tmp_path, parse(tmp_path, src), rp)

    items = audit_items(out, "图片段解除固定行距")
    assert items, "固定行距图片段的条目名不应改变"
    assert "固定行距会裁切内嵌图片" in items[0]["note"]
    assert "行距" in items[0]["note"], f"note 按实际差异列项：{items[0]['note']}"
    assert audit_items(out, "图片段套版") == []
