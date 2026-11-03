"""第七批渲染侧回归钉（IDMPOT / SDTCAP / EXTFOOTAUD / F13STR）。

- IDMPOT  渲染器对自家成品二次渲染时 sjjj/cwe 版式指纹漂移（楷体→宋体、
         jc center→both，文字逐字不变）。根因全在解析端「读不回渲染器按
         规则集写出的形态」：sjjj 的无冒号标签（「内容提要　」）、空白分隔
         一级编号（「一　」）、cwe 的「1.」一级模板（被「1.」歧义机制判成
         二级）与「References」文献表标题。修法在 parse_docx 的识别模式与
         render_from_ir 的编号剥除（NUM_PREFIX[3] 防半截剥除）。全库双渲染
         幂等守卫落在 test_fix_full_library_matrix（每本渲染两次、版式签名
         一致）；本文件钉 sjjj/cwe 这两本点名刊与各识别模式的单元行为。
- SDTCAP  sdt 递归下钻无深度上限：超长嵌套链把 _sdt_blocks 打穿
         （RecursionError）。修法：depth 上限 8，超深不再下钻、整体计数进
         unsupported_content_blocks。
- EXTFOOTAUD  「N 条外部脚注关系未处理」此前只进 stdout 汇总，下游只读
         *.audit.json 看不到。修法：计数非零时 audit.add 一条说明型条目。
- F13STR  规则集深校验对字符串叶子静默放行：typography.body.line_spacing
         该是 dict 却给了字符串，渲染中途才 AttributeError。修法：必填键
         条目声明期望形态，形态不符计入缺失清单。

端到端用例走子进程（与 test_citation_convert.py / test_fix_page_setup_sections.py
同一夹具姿势）；纯逻辑判定直接 import 工具模块断言。
"""

import json
import re
import subprocess
import sys

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from conftest import RULES, TOOLS, layout_signature, mutated_rules, render_into

sys.path.insert(0, str(TOOLS))
import parse_docx as pd  # noqa: E402
import render_from_ir as rf  # noqa: E402

W = pd.W
# XML 字面量里的命名空间要用裸 URI（xmlns:w="…"），不是 ET 的花括号形态
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 点名刊只在私有库盘上才有（不入库）：文件存在才参数化，CI 只有入库 3 本
# 时这两个用例自动不生成——与全库矩阵「文件存在才跑」同一约定
SJJJ_CWE = [j for j in ("sjjj", "cwe") if (RULES / f"{j}.json").exists()]


# ---------------------------------------------------------------- IDMPOT

def test_parse_reads_back_renderer_label_forms():
    """渲染器按规则集写出的标签形态必须能被解析读回（IDMPOT 根因单元钉）。"""
    # sjjj：标签后是全角空格、没有冒号
    assert pd.PAT["abstract"].match("内容提要　本文以美联储数据为样本。")
    assert pd.PAT["keywords"].match("关  键  词　宽松货币政策　美联储")
    # kjyj：括号体例，闭合括号本身就是分隔
    assert pd.PAT["abstract"].match("【摘要】本文考察数字金融。")
    assert pd.PAT["keywords"].match("【关键词】数字金融；融资约束")
    # 无分隔的串仍不当摘要（既无冒号也无空白/括号）
    assert not pd.PAT["abstract"].match("内容提要本文以美联储数据为样本。")
    # 文献表标题：中文冒号体例 / 纯词体例 / 括号体例 / 英文刊
    for head in ("参考文献：", "参 考 文 献", "﹝参考文献﹞", "References",
                 "REFERENCES"):
        assert pd.PAT["ref_head"].match(head), head
    assert not pd.PAT["ref_head"].match("参考文献综述如下")


def test_parse_reads_back_whitespace_separated_h1():
    """「一　引言」（sjjj 一级体例）读回一级标题；正文短句不误判。"""
    assert pd.classify_heading("一　引  言") == (1, "high")
    assert pd.classify_heading("十二　文献综述与理论框架") == (1, "high")
    # 护栏：带句读 / 超长的不当标题（与数字编号同一口径）
    assert pd.classify_heading("一 些研究表明上述结论在更换度量方式后依然成立，样本为A股。") is None


def _para(text):
    p = pd.ET.Element(W + "p")
    r = pd.ET.SubElement(p, W + "r")
    t = pd.ET.SubElement(r, W + "t")
    t.text = text
    return (0, p, None)


def test_detect_arabic_level_resolves_1dot_and_paren_by_context():
    """「1.」「（1）」的层级按全文编号体系判定（IDMPOT：cwe 与经济管理系）。"""
    # 阿拉伯层级体例（1／1.1，无任何中文编号）：「1.」是一级
    assert pd.detect_arabic_level([_para("1. 引言"), _para("2.1 文献综述")]) \
        == (1, True, 4)
    # 一、／1.／（1）体系（有「一、」无「（一）」）：1. 是二级，（1）顺延三级
    assert pd.detect_arabic_level(
        [_para("一、引言"), _para("1. 理论基础"), _para("（1）货币政策传导机制")]) \
        == (2, True, 3)
    # （一）在场：「1.」是三级，「（1）」保持缺省四级（既有行为）
    assert pd.detect_arabic_level([_para("（一）理论基础"), _para("1. 小节")]) \
        == (3, True, 4)
    # 「（1）」在场但全文没有「一、」：「1.」仍判一级——「（1）」开头的不
    # 一定是标题（正文枚举句同形），不能作为判二级的依据；（1）保持缺省四级
    assert pd.detect_arabic_level([_para("（1）某点"), _para("1. 某条")]) \
        == (1, True, 4)


@pytest.mark.parametrize("jid", SJJJ_CWE, ids=lambda j: j)
def test_double_render_layout_stable(sample_docx, tmp_path, jid):
    """点名刊双渲染：二次渲染的版式签名与第一次逐字一致（承诺③）。"""
    first = render_into(tmp_path / f"b7-idem1-{jid}", sample_docx,
                        target=jid, rules_path=RULES / f"{jid}.json")
    second = render_into(tmp_path / f"b7-idem2-{jid}", first.docx,
                         target=jid, rules_path=RULES / f"{jid}.json")
    assert layout_signature(second.docx) == layout_signature(first.docx), \
        f"{jid} 对自家成品二次渲染版式指纹漂移"


def test_level3_numeric_prefix_not_half_stripped(sample_docx, tmp_path):
    """「2.1.1 标题」重渲染不得剥成「1.1 标题」再叠编号（cwe 实测漂移点）。"""
    rules = RULES / "cwe.json"
    if not rules.exists():
        pytest.skip("cwe 不在本机规则库")
    first = render_into(tmp_path / "b7-cwe1", sample_docx, target="cwe",
                        rules_path=rules)
    second = render_into(tmp_path / "b7-cwe2", first.docx, target="cwe",
                         rules_path=rules)
    # 二次渲染的标题里不得出现叠加编号（「2.1.1 1.1 …」——链式前缀剥一半的残骸）
    stacked = [b["text"] for b in second.ir["blocks"]
               if b["type"].startswith("heading")
               and re.search(r"\d+\.\d+\.\d+\s+\d+\.\d+", b["text"])]
    assert not stacked, f"编号叠加：{stacked}"


# ---------------------------------------------------------------- SDTCAP

SDT_CHAIN = (
    f'<w:sdt xmlns:w="{W_NS}"><w:sdtPr><w:id w:val="1"/></w:sdtPr>'
    '<w:sdtContent>{inner}</w:sdtContent></w:sdt>'
)


def build_deep_sdt_docx(path, depth, inner_text="深层正文段落"):
    """body 级嵌套 sdt 链：depth 层，最里层一个段落。"""
    doc = Document()
    doc.add_paragraph("文章标题：数字金融研究")
    doc.add_paragraph("正文段落，考察数字金融对企业融资约束的影响，样本为A股上市公司。")
    inner = f'<w:p xmlns:w="{W_NS}"><w:r><w:t>{inner_text}</w:t></w:r></w:p>'
    xml = inner
    for _ in range(depth):
        xml = SDT_CHAIN.format(inner=xml)
    doc.element.body.append(parse_xml(xml))
    doc.save(str(path))
    return path


def parse_ir(tmp_path, docx):
    ir = tmp_path / "deep.ir.json"
    r = subprocess.run(
        [sys.executable, str(TOOLS / "parse_docx.py"), str(docx), "-o", str(ir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    return r, (json.loads(ir.read_text(encoding="utf-8")) if ir.exists() else None)


def test_deep_sdt_chain_does_not_crash_and_counts_unsupported(tmp_path):
    """depth=60 嵌套链：解析不崩，超深部分计 unsupported，深层文字不进 IR。"""
    docx = build_deep_sdt_docx(tmp_path / "deep.docx", 60)
    r, ir = parse_ir(tmp_path, docx)
    assert r.returncode == 0, (r.stdout[-400:] + r.stderr[-400:])
    assert ir is not None
    assert ir.get("unsupported_content_blocks", 0) >= 1, "超深嵌套要计数进 IR"
    texts = " ".join(b.get("text", "") for b in ir["blocks"])
    assert "深层正文段落" not in texts, "超过深度上限的内容不应被下钻取出"
    # 解析汇总要把这笔账说出口（stdout 说明）
    assert "未处理内容块" in r.stdout


def test_moderate_sdt_nesting_still_descends(tmp_path):
    """上限内的嵌套（3 层）照常下钻：段落进 IR、不计 unsupported。"""
    docx = build_deep_sdt_docx(tmp_path / "deep3.docx", 3)
    r, ir = parse_ir(tmp_path, docx)
    assert r.returncode == 0, (r.stdout[-400:] + r.stderr[-400:])
    assert not ir.get("unsupported_content_blocks")
    texts = " ".join(b.get("text", "") for b in ir["blocks"])
    assert "深层正文段落" in texts, "上限内的内容控件正文不得丢失"


def test_render_side_sdt_depth_cap_matches_parse(tmp_path):
    """渲染端 _sdt_blocks 与解析端同一深度上限：超深 sdt 不下钻、不崩。"""
    docx = build_deep_sdt_docx(tmp_path / "deep.docx", 60)
    r, _ir = parse_ir(tmp_path, docx)
    assert r.returncode == 0
    # 渲染：IR 与源稿同名配套（render_from_ir 按 IR 路径找 docx）
    out = tmp_path / "deep.out.docx"
    rr = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(tmp_path / "deep.ir.json"),
         str(RULES / "tjyj.json"), "-o", str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    assert rr.returncode == 0, rr.stderr[-500:]
    assert "未处理内容块" in rr.stdout


# ------------------------------------------------------------ EXTFOOTAUD

FOOTNOTES_RELTYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                     "relationships/footnotes")


def build_external_footnote_docx(path):
    """带 TargetMode=External 脚注关系的稿件（外部超链接式脚注，无可解析部件）。"""
    doc = Document()
    doc.add_paragraph("正文一段，考察数字金融对企业融资约束的影响。")
    doc.part.rels.add_relationship(
        FOOTNOTES_RELTYPE, "http://example.com/footnotes.xml", "rIdFnExt",
        is_external=True)
    doc.save(str(path))
    return path


def test_external_footnote_relations_land_in_audit_file(tmp_path):
    """外部脚注关系计数必须进 *.audit.json（EXTFOOTAUD），不能只留 stdout。"""
    d = tmp_path / "extfoot"
    d.mkdir()
    src = build_external_footnote_docx(d / "manuscript.docx")
    r = subprocess.run(
        [sys.executable, str(TOOLS / "parse_docx.py"), str(src), "-o",
         str(d / "manuscript.ir.json")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    assert r.returncode == 0, r.stderr[-400:]
    out = d / "manuscript.tjyj.docx"
    rr = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"),
         str(d / "manuscript.ir.json"), str(RULES / "tjyj.json"), "-o", str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    assert rr.returncode == 0, rr.stderr[-500:]
    # stdout 汇总保留（既有行为），审计文件里必须有同一条说明
    assert "外部脚注关系未处理" in rr.stdout
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    hits = [it for it in audit["items"] if it["action"] == "外部脚注关系未处理"]
    assert hits, f"audit.json 缺少外部脚注关系说明条目：{[i['action'] for i in audit['items']]}"
    it = hits[0]
    assert it["before"] == "" and it["after"] == "", "说明型条目不携带文字快照"
    assert it.get("changed") is False, "未处理不是改动，不得计入改动数"
    assert "1 条" in it["note"], f"note 要带计数与含义：{it['note']}"


# ---------------------------------------------------------------- F13STR

def test_shape_mismatch_reported_before_rendering(tmp_path):
    """line_spacing 给了字符串：入口干净报错（形态不符），不再渲染中途崩。"""
    base = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    base["typography"]["body"]["line_spacing"] = "1.5倍"
    rules = tmp_path / "b7-shape.json"
    rules.write_text(json.dumps(base, ensure_ascii=False), encoding="utf-8")

    doc = Document()
    doc.add_paragraph("正文一段。")
    doc.save(str(tmp_path / "manuscript.docx"))
    ir = tmp_path / "manuscript.ir.json"
    pr = subprocess.run(
        [sys.executable, str(TOOLS / "parse_docx.py"), str(tmp_path / "manuscript.docx"),
         "-o", str(ir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    assert pr.returncode == 0

    rr = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir), str(rules),
         "-o", str(tmp_path / "out.docx")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(TOOLS.parent))
    assert rr.returncode == 1, rr.stdout[-300:]
    assert "Traceback" not in rr.stderr, "形态错误不得漏到渲染中途才崩"
    assert rf.RULESET_ERR_PREFIX in rr.stderr
    assert "typography.body.line_spacing" in rr.stderr
    assert "形态不符：期望 dict" in rr.stderr


def test_shape_checks_unit():
    """形态断言的单元口径：数值/字符串/list 各自报对期望名，bool 不算数值。"""
    rules = {
        "journal": {"name_zh": "统计研究"},
        "typography": {"body": {"font_cn": "宋体", "size_pt": 10.5,
                                "line_spacing": {"mode": "exact", "value_pt": 15.5},
                                "first_line_indent_chars": 2},
                       "title": {}},
        "headings": {"level1": {}, "levels": ["一、"]},
        "blocks": [{"type": "body"}],
        "page": {"size": "A4", "size_mm": {"width": 210, "height": 297},
                 "margin_mm": {"top": 25, "bottom": 20, "left": 25, "right": 24}},
        "tables": {"border_pt": {"top": 1.5, "header_bottom": 0.75, "bottom": 1.5},
                   "note_layers": []},
        "figures": {"note_layers": []},
        "citation": {},
    }
    assert rf.missing_required_keys(rules) == [], "齐全的规则集不得误报"

    rules["typography"]["body"]["size_pt"] = "10.5"      # 数值键给了字符串
    assert "typography.body.size_pt 形态不符：期望 数值" in rf.missing_required_keys(rules)
    rules["typography"]["body"]["size_pt"] = True        # bool 不当数值
    assert "typography.body.size_pt 形态不符：期望 数值" in rf.missing_required_keys(rules)
    rules["typography"]["body"]["size_pt"] = 10.5
    rules["headings"]["levels"] = "一、"                  # 数组键给了字符串
    assert rf.missing_required_keys(rules) == \
        ["headings.levels 形态不符：期望 list"]
    rules["headings"]["levels"] = ["一、"]
    rules["blocks"] = {"type": "body"}                    # 数组键给了 dict
    assert "blocks 形态不符：期望 list" in rf.missing_required_keys(rules)


def test_shape_declarations_match_whole_library():
    """形态声明不得误杀任何一本实有规则集（本机 23 本全过，CI 3 本）。"""
    for p in sorted(RULES.glob("*.json")):
        rules = json.loads(p.read_text(encoding="utf-8"))
        assert rf.missing_required_keys(rules) == [], \
            f"{p.stem} 被形态校验误报：{rf.missing_required_keys(rules)}"
