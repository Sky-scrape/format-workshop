"""三个已确认问题的回归钉：段首缩进（F29）、审计漏报（F25）、审计虚报（F28）。

- F29  手工敲段首全角空格缩进的稿，原文空格保留、渲染器又无条件设首行缩进
       ——全稿双重缩进。修法：N-01 段首版＋N-08 落地，段首缩进空白删除并逐块
       进审计（B 档、可驳回、幂等）。
- F25  keywords／表注／文献表标题／作者简介／文献条目／标题等 restyle 分支不出
       审计条目；fmt_snapshot 缺对齐与缩进，「仅对齐/缩进变了」不产生条目。
       修法：快照扩为字号/字体/加粗/行距(值+规则)/对齐/缩进/段距，统一走
       「先拍照、套版、再拍照、变了才记」的助手。
- F28  三线表、图片段解除固定行距、参考文献悬挂缩进无条件出条目——本来就
       合规的对象重跑仍列「改动」。修法：三处都先取 before 快照、应用后取
       after，无差异不出条目。
"""

import importlib.util
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt

from conftest import RULES, render_into, run_tool

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"


def _load(name, filename):
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, TOOLS / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


render = _load("render_from_ir_audit_indent_under_test", "render_from_ir.py")
sys.path.insert(0, str(TOOLS))
from rules_ref import resolvable  # noqa: E402

STRIP_ACTION = "段首缩进空格删除"

# 所有「版式套版」类动作：合规对象重跑时一条都不该再出（F25+F28）
# 「图片段套版」是批五拆出来的新动作名（FIGACT：图片段只动段前/段后距、
# 未涉行距时不再叫「解除固定行距」）——第二遍零套版断言必须对新名设防
RESTYLE_FAMILY = {
    "文章标题套版", "副标题套版", "作者署名套版", "正文套版", "标题套版",
    "摘要套版", "关键词套版", "题注套版", "表注套版", "文献表标题套版",
    "作者简介套版", "参考文献条目套版", "补录块套版",
    "表格重设为三线表", "图片段解除固定行距", "图片段套版", "公式段行距",
    "参考文献悬挂缩进", "参考文献缩进", STRIP_ACTION,
}

INDENT_WS = "\u3000 \t"


def build_indent_doc(path):
    """一份手工敲段首缩进的稿：摘要、关键词、正文都带全角空格起头。

    关键词分隔符直接写目标刊的「；」，让分隔符归一保持沉默——这条用例只
    核对段首空格这一件事，别让别的白名单归一混进逐字比对。
    """
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("\u3000\u3000内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("\u3000\u3000关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("\u3000\u3000本文考察数字金融对企业融资约束的影响，样本为A股上市公司。")
    doc.add_paragraph("　 \t稳健性检验表明上述结论在更换度量方式后依然成立。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("\u3000\u3000[1] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.save(str(path))
    return path


def leading_ws_blocks(result):
    """源稿里带段首缩进空白的正文类块（按 IR 的 idx）。

    只看 body/abstract/keywords——文献条目的段首空格不在动作范围内（见
    test_leading_indent_spaces_stripped_and_audited 的对照组）。"""
    return {i for i, t in result.source_texts.items()
            if t[:1] in ("\u3000", " ", "\t")
            and result.block_type.get(i) in ("body", "abstract", "keywords")}


def first_line_chars(docx_path, idx):
    """读成品某段的 firstLineChars 缩进设置。"""
    doc = Document(str(docx_path))
    child = list(doc.element.body)[idx]
    ppr = child.find(qn("w:pPr"))
    if ppr is None:
        return None
    ind = ppr.find(qn("w:ind"))
    return None if ind is None else ind.get(qn("w:firstLineChars"))


def items_of(result, action, idx=None):
    return [it for it in result.items
            if it["action"] == action and (idx is None or it["idx"] == idx)]


# ---------- F29：段首缩进空格删除 ----------

def test_leading_indent_spaces_stripped_and_audited(tmp_path):
    """段首全角空格被删、其余逐字不变；逐块出审计条目且依据可解析。"""
    src = build_indent_doc(tmp_path / "indent.docx")
    result = render_into(tmp_path / "run", src, target="tjyj")
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))

    stripped = leading_ws_blocks(result)
    assert stripped, "夹具里没有带段首缩进的块"
    for idx in stripped:
        btype = result.block_type.get(idx, "")
        assert btype in ("body", "abstract", "keywords"), \
            f"块 {idx} 被判成 {btype}，夹具布局变了"
        # 只少了段首缩进空白，其余逐字不变
        assert result.output_texts[idx] == result.source_texts[idx].lstrip(INDENT_WS), \
            f"块 {idx} 除段首空格外发生了其他文字改动"
        # 首行缩进改由段落设置接管，不再叠手工空格
        assert first_line_chars(result.docx, idx) == "200", \
            f"块 {idx} 没有写入首行缩进设置"

        entries = items_of(result, STRIP_ACTION, idx)
        assert len(entries) == 1, f"块 {idx} 缺段首缩进审计条目"
        it = entries[0]
        assert it["before"].startswith("\u3000"), "条目 before 没记下原空格"
        assert it["after"] == render._frag(result.output_texts[idx]), \
            "条目 after 与成品文字不符"
        assert it["rule"] == "typography.body.first_line_indent_chars"
        assert resolvable(rules, it["rule"]), "段首缩进条目的依据解析不回规则集"

    # 非正文块（文献条目）的段首空格不在白名单动作之内，原样保留
    ref_idx = next(i for i, t in result.block_type.items()
                   if t == "reference_item")
    assert result.output_texts[ref_idx] == result.source_texts[ref_idx], \
        "文献条目的段首空格不该被动"
    assert not items_of(result, STRIP_ACTION, ref_idx)


def test_rejected_strip_keeps_text_but_still_sets_indent(tmp_path):
    """驳回段首缩进删除：保留原空格，缩进设置仍写入，条目如实记「已驳回」。"""
    src = build_indent_doc(tmp_path / "indent2.docx")
    d = tmp_path / "probe"
    d.mkdir()
    run_tool("parse_docx.py", src, "-o", d / "m.ir.json")
    ir = json.loads((d / "m.ir.json").read_text(encoding="utf-8"))
    # IR 的 text 已剥过首尾空白（parse_docx.para_text），按正文内容定位；
    # 「稳健性」那段在 docx 里带三种缩进空白混合的段首
    body_idx = next(b["idx"] for b in ir["blocks"]
                    if b["type"] == "body" and b["text"].startswith("稳健性检验"))

    result = render_into(tmp_path / "run", src, target="tjyj",
                         rejects=[f"{body_idx}|{STRIP_ACTION}"])
    entries = items_of(result, STRIP_ACTION, body_idx)
    assert len(entries) == 1 and entries[0]["rejected"], "驳回没有进审计"
    assert "已驳回" in entries[0]["note"]
    assert result.output_texts[body_idx] == result.source_texts[body_idx], \
        "被驳回的段首空格没有被保留"
    assert first_line_chars(result.docx, body_idx) == "200", \
        "驳回文字改动时，首行缩进设置也应照写"


def test_second_render_no_new_indent_entries(tmp_path):
    """幂等：第二遍渲染不再出段首条目，文本保持不变。"""
    src = build_indent_doc(tmp_path / "indent3.docx")
    first = render_into(tmp_path / "pass1", src, target="tjyj")
    second = render_into(tmp_path / "pass2", first.docx, target="tjyj")

    assert not items_of(second, STRIP_ACTION), \
        f"第二遍渲染又出了段首条目：{[i['id'] for i in items_of(second, STRIP_ACTION)]}"
    for idx in leading_ws_blocks(first):
        assert second.output_texts[idx] == first.output_texts[idx], \
            f"块 {idx} 的文字在第二遍渲染中变了"


# ---------- F25：restyle 分支的审计不再漏报 ----------

def test_abstract_and_keywords_restyle_are_audited(tmp_path):
    """改 keywords／abstract 的字体字号，审计必须出现对应条目（原先零记录）。"""
    src = tmp_path / "plain.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；货币政策")
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    kw = [it for it in result.items if it["action"] == "关键词套版"]
    ab = [it for it in result.items if it["action"] == "摘要套版"]
    assert kw and ab, f"摘要/关键词套版没有进审计：{sorted({i['action'] for i in result.items})}"
    assert kw[0]["rule"] == "typography.keywords"
    assert ab[0]["rule"] == "typography.abstract"
    assert "9" in kw[0]["note"] and "pt" in kw[0]["note"], \
        f"条目没写目标值：{kw[0]['note']}"


def test_alignment_only_change_is_audited(tmp_path):
    """字体字号行距缩进都已合规、只有对齐不合规：也要出条目（快照缺对齐时漏报）。"""
    src = tmp_path / "align.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    p = doc.add_paragraph("关键词：数字金融；融资约束；货币政策")
    # 按 tjyj 的 keywords 规则先排好：黑体 9pt、Times、固定行距 15pt、缩进 2 字，
    # 只留左对齐不合规
    render.set_run(p.runs[0], cn="黑体", latin="Times New Roman", size=9)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing = Pt(15)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    render.set_chars_indent(p, 2)
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    kw_idx = next(i for i, t in result.block_type.items() if t == "keywords")
    entries = items_of(result, "关键词套版", kw_idx)
    assert len(entries) == 1, f"仅对齐变化没有出条目：{[i['action'] for i in result.items]}"
    assert entries[0]["note"].startswith("改对齐 →"), \
        f"条目没说清只改了对齐：{entries[0]['note']}"


def test_fmt_snapshot_covers_alignment_indent_spacing_and_bold():
    """快照必须覆盖对齐、缩进、段距、行距规则与加粗，fmt_delta 要能说出改了什么。"""
    p = Document().add_paragraph("样本段落")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.line_spacing = Pt(17)
    render.set_chars_indent(p, 2)

    snap = render.fmt_snapshot(p)
    assert snap["jc"] == WD_ALIGN_PARAGRAPH.CENTER
    assert snap["sb"] == 6
    assert snap["rule"] == WD_LINE_SPACING.EXACTLY
    assert any(k.endswith("firstLineChars") and v == "200" for k, v in snap["ind"])

    before = render.fmt_snapshot(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    assert "对齐" in render.fmt_delta(before, render.fmt_snapshot(p), "目标")
    render.set_chars_indent(p, 4)
    assert "缩进" in render.fmt_delta(before, render.fmt_snapshot(p), "目标")

    p2 = Document().add_paragraph()
    p2.add_run("加粗核验").font.bold = True
    assert render.fmt_snapshot(p2)["bold"] == {True}
    render.restyle(p2, cn="宋体", latin="Times New Roman", size=10.5, align=None, bold=False)
    assert render.fmt_snapshot(p2)["bold"] == {False}


def test_second_pass_emits_no_restyle_entries(sample_docx, tmp_path):
    """同刊重跑：版式已合规，第二遍审计里不再重出任何套版类条目（F25+F28 汇总）。"""
    first = render_into(tmp_path / "pass1", sample_docx, target="tjyj")
    second = render_into(tmp_path / "pass2", first.docx, target="tjyj")

    again = [it["action"] for it in second.items if it["action"] in RESTYLE_FAMILY]
    assert not again, f"第二遍渲染重出了 {len(again)} 条套版条目：{sorted(set(again))}"
    # 第一遍确实有套版条目（否则这条用例测了个寂寞）
    assert [it for it in first.items if it["action"] in RESTYLE_FAMILY], \
        "第一遍渲染没有产生任何套版条目，夹具选刊可能过于接近"


# ---------- F28：三线表／图片段／悬挂缩进不再虚报 ----------

def test_plain_table_gets_exactly_one_table_entry(tmp_path):
    """真需要改的表：恰好一条「表格重设为三线表」，说明写清边框与表内字号。"""
    src = tmp_path / "table.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "变量"
    table.cell(0, 1).text = "定义"
    table.cell(1, 0).text = "Size"
    table.cell(1, 1).text = "企业规模"
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    tbl_idx = next(i for i, t in result.block_type.items() if t == "table")
    entries = items_of(result, "表格重设为三线表", tbl_idx)
    assert len(entries) == 1, f"表格条目应恰好一条：{len(entries)}"
    assert "边框重设为三线表" in entries[0]["note"]
    assert "表内文字改为" in entries[0]["note"]
    assert entries[0]["rule"] == "tables.border_pt"


def test_compliant_table_second_pass_has_no_table_entry(sample_docx, tmp_path):
    """合规三线表二次套版：零表格条目（原先每张表每次重跑都重报）。"""
    first = render_into(tmp_path / "pass1", sample_docx, target="tjyj")
    assert items_of(first, "表格重设为三线表"), "夹具里的表第一遍没有任何改动，检查选刊"
    second = render_into(tmp_path / "pass2", first.docx, target="tjyj")
    assert not items_of(second, "表格重设为三线表"), \
        f"已合规的表重跑又报了改动：{[i['id'] for i in items_of(second, '表格重设为三线表')]}"


def _build_ref_doc(path, hang_on_first=False):
    """文献表：第一条可选地带上目标刊的悬挂缩进（用渲染器自己的函数写）。"""
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("参考文献")
    for k, text in enumerate([
        "[1] 张三. 数字金融研究[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] 李四. 融资约束研究[J]. 金融研究, 2019, 40(5): 1-9.",
        "[3] 王五. 银行风险研究[J]. 财贸经济, 2018, 39(6): 1-8.",
    ]):
        p = doc.add_paragraph(text)
        if hang_on_first and k == 0:
            render.set_hanging(p, 2.3)   # tjyj 的 hanging_indent_chars
    doc.save(str(path))
    return path


def test_hanging_indent_reported_only_when_changed(tmp_path):
    """悬挂缩进：已合规的条目不再报；真缺缩进的条目恰一条。"""
    src = _build_ref_doc(tmp_path / "refs.docx", hang_on_first=True)
    result = render_into(tmp_path / "run", src, target="tjyj")

    ref_idxs = sorted(i for i, t in result.block_type.items() if t == "reference_item")
    assert len(ref_idxs) == 3, "夹具里没有三条文献"
    audited = {it["idx"] for it in items_of(result, "参考文献悬挂缩进")}
    assert audited == set(ref_idxs[1:]), \
        f"悬挂缩进条目该只出现在真缺缩进的条目上：{sorted(audited)}"
    for it in items_of(result, "参考文献悬挂缩进"):
        assert it["rule"] == "citation.reference_list.hanging_indent_chars"
        assert it["after"] == "2.3 字符"

    # 对照组：全都不带悬挂缩进时，每条恰一条
    src2 = _build_ref_doc(tmp_path / "refs2.docx", hang_on_first=False)
    result2 = render_into(tmp_path / "run2", src2, target="tjyj")
    assert len(items_of(result2, "参考文献悬挂缩进")) == len(ref_idxs)


def test_figure_paragraph_not_reported_when_already_compliant(tmp_path):
    """已是单倍行距且居中的图片段：不再报「解除固定行距」。
    （真带固定行距、需要解除的情形由 test_figure_safety.py 钉着。）"""
    from io import BytesIO

    import base64
    PNG_1PX = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
    src = tmp_path / "fig.docx"
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    fig = doc.add_paragraph()
    fig.add_run().add_picture(BytesIO(PNG_1PX))
    fig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fig.paragraph_format.line_spacing = 1.0
    doc.save(str(src))

    result = render_into(tmp_path / "run", src, target="tjyj")
    assert not items_of(result, "图片段解除固定行距"), \
        f"已合规的图片段又报了改动：{[i['id'] for i in items_of(result, '图片段解除固定行距')]}"
