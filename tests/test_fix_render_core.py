"""渲染与引用六处修复的回归钉：F13/F21/F22/F23/F24/F30。

每个用例对应问题单里的证据复现：

- F13  规则集缺字段 → 渲染中途裸 KeyError 崩溃、traceback 外泄；
       现在入口一次报全、带专属前缀干净退出
- F21  排序行为从不读 citation.reference_list.sort 却冒用其名出审计；
       现在按声明排，appearance 不排，未知声明保持原序 + 如实 note
- F22  pypinyin 未安装，码位序冒充拼音序（「周一」排在「张三」前）；
       现在用 GB 序（一级字库即拼音序），note 如实写「GB 序（近似拼音）」
- F23  非编号条目与脚注编号引用两头不管；现在 fail-closed 整批不转并点名
- F24  w:sdt 整块跳过、文本框文字混进段落文本；现在 sdt 内段落进 IR 且
       套版，文本框文字不参与分类且原样保留
- F30  内置缺省值构成第二真源；现在集中在 DEFAULTS，用到的审计条目带
       '~' 可缺省标注与「缺省值」说明（推定值专指规则集 measured:false 的
       取值，代码缺省不冒称推定），数值行为不变

夹具姿势与 test_citation_convert.py 相同：python-docx 造稿、子进程跑
parse/render（参数列表直传，不经 shell），不 import 渲染器进程内的可变状态。
"""

import json
import subprocess
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from conftest import RULES, TOOLS

sys.path.insert(0, str(TOOLS))
import citation_ref as cr  # noqa: E402
import render_from_ir as rf  # noqa: E402

TEMPLATES = {
    "cn_journal": "{authors}：《{title}》，《{journal}》，{year}年{issue_part}。",
    "cn_book": "{authors}：《{title}》，{publisher}，{year}年。",
    "cn_other": "{authors}：《{title}》，{source}，{year}年。",
    "en_journal": "{authors}，{year}，“{title}”，{journal}{volume_part}{page_part}.",
    "en_book": "{authors}，{year}，“{title}”，{publisher}.",
    "en_other": "{authors}，{year}，“{title}”，{source}.",
}

# 证据复现的名字：拼音序里 张三(zhang) 在 周一(zhou) 前；
# 修复前的码位序恰好相反（周一 U+5468 < 张三 U+5F20）。
REF_ZHOU = "[1] 周一. 融资约束的度量[J]. 金融研究, 2021, 76(2): 20-40."
REF_ZHANG = "[2] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10."


# ---------------------------------------------------------------- 夹具

def load_rules(tmp_path, journal="tjyj", name=None, mutate=None):
    """读仓库自带规则集的副本；mutate(rules) 可在落盘前改字段。"""
    rules = json.loads((RULES / f"{journal}.json").read_text(encoding="utf-8"))
    if mutate is not None:
        mutate(rules)
    p = tmp_path / (name or f"{journal}.json")
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return p


def citation_rules(tmp_path, sort=None, base="tjyj", name="cite_rules.json"):
    """作者-年份制规则集 + 条目模板；sort 覆盖排序声明。"""
    def mutate(rules):
        rules["citation"]["system"] = "author-year"
        rl = rules["citation"].setdefault("reference_list", {})
        rl["templates"] = TEMPLATES
        rl["must_require_bidirectional_match"] = True
        if sort is not None:
            rl["sort"] = sort
    return load_rules(tmp_path, base, name, mutate)


def make_src(tmp_path, name="src", extra_refs=()):
    """双条目夹具：正文引 [1][2]，文献表 周一 在前、张三 在后——
    任何「按拼音/字母」的真排序都必须把 张三 挪到前面。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHOU)
    doc.add_paragraph(REF_ZHANG)
    for t in extra_refs:
        doc.add_paragraph(t)
    p = tmp_path / f"{name}.docx"
    doc.save(str(p))
    return p


def parse_render(tmp_path, docx, rules_path, citation=True):
    """解析 + 渲染一遍，返回渲染子进程结果与成品路径。

    两条分支的参数都是写死的参数列表（不经 shell，参考
    test_citation_convert.run_render 的姿势）。
    """
    ir = docx.with_suffix(".ir.json")
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(docx),
                    "-o", str(ir)], capture_output=True, text=True, check=True)
    out = docx.with_name("out.docx")
    if citation:
        r = subprocess.run(
            [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir),
             str(rules_path), "-o", str(out), "--citation"],
            capture_output=True, text=True)
    else:
        r = subprocess.run(
            [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir),
             str(rules_path), "-o", str(out)],
            capture_output=True, text=True)
    return r, out


def out_texts(out):
    return [p.text for p in Document(str(out)).paragraphs]


def audit_items(out, action=None):
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    items = audit["items"]
    return [i for i in items if i["action"] == action] if action else items


def ref_positions(out):
    """成品里两条文献条目的位置：张三 / 周一 谁在前。"""
    return {t[:2]: i for i, t in enumerate(out_texts(out))
            if t.startswith(("张三", "周一"))}


# ---------------------------------------------------------------- F13

def test_missing_ruleset_keys_fail_clean_and_complete(tmp_path):
    """删掉两个键：一次报全、指名缺键、无 traceback、不产出成品。"""
    src = make_src(tmp_path)

    def mutate(rules):
        del rules["figures"]["note_layers"]
        del rules["typography"]["body"]["first_line_indent_chars"]
    rp = load_rules(tmp_path, name="broken.json", mutate=mutate)
    r, out = parse_render(tmp_path, src, rp, citation=False)

    assert r.returncode == 1, f"应干净退出，实际 exit {r.returncode}"
    err = r.stderr
    assert "[RULESET-INCOMPLETE]" in err, "缺专属前缀，上层无法识别这是规则集问题"
    assert "规则集 tjyj 缺少字段 typography.body.first_line_indent_chars" in err
    assert "规则集 tjyj 缺少字段 figures.note_layers" in err, "缺哪些要一次报全"
    assert "（章节" in err, "错误信息要带章节说明"
    assert "Traceback" not in err and "KeyError" not in err, "traceback 不外泄"
    assert not out.exists(), "校验失败不应产出成品文件"


# ---------------------------------------------------------------- F21/F22

def test_gb_sort_key_puts_zhang_before_zhou():
    """F22 证据复现：拼音序应 张三 < 周一，码位序恰好相反。"""
    assert "周一" < "张三"                      # 码位序（旧实现的实际行为）
    assert cr.gb_sort_key("张三") < cr.gb_sort_key("周一")  # GB 序（新实现）


def test_pinyin_scheme_reorders_and_note_is_honest(tmp_path):
    """F21+F22 端到端：声明拼音类 scheme → 真重排，note 写 GB 序。"""
    src = make_src(tmp_path)
    rp = citation_rules(tmp_path, sort="cn-first-by-pinyin-then-en-by-alphabet",
                        name="pinyin.json")
    r, out = parse_render(tmp_path, src, rp)
    assert r.returncode == 0, r.stderr[-500:]

    pos = ref_positions(out)
    assert pos["张三"] < pos["周一"], "GB 序（≈拼音）应把 张三 排在 周一 前"

    items = audit_items(out, "文献表排序")
    assert items, "排序发生了却没有审计条目"
    assert items[0]["rule"] == "citation.reference_list.sort"
    assert "GB 序（近似拼音）" in items[0]["note"], items[0]["note"]
    assert "拼音在前" not in items[0]["note"], "不得谎称真拼音序"


def test_appearance_scheme_keeps_original_order(tmp_path):
    """F21：appearance 类声明不重排。"""
    src = make_src(tmp_path)
    rp = citation_rules(tmp_path, sort="by_appearance_order", name="appear.json")
    r, out = parse_render(tmp_path, src, rp)
    assert r.returncode == 0, r.stderr[-500:]

    pos = ref_positions(out)
    assert pos["周一"] < pos["张三"], "声明按出现序的刊不该被重排"
    assert audit_items(out, "文献表排序") == [], "顺序没变就不该有排序条目"


def test_unknown_sort_scheme_keeps_order_with_honest_note(tmp_path):
    """F21：声明了渲染器不认识的 scheme → 保持原序 + note 如实说明。"""
    src = make_src(tmp_path)
    rp = citation_rules(tmp_path, sort="by-magic", name="unknown.json")
    r, out = parse_render(tmp_path, src, rp)
    assert r.returncode == 0, r.stderr[-500:]

    pos = ref_positions(out)
    assert pos["周一"] < pos["张三"], "不认识的排序声明不该瞎排"

    items = audit_items(out, "文献表排序")
    assert items and "未支持的排序声明「by-magic」" in items[0]["note"]
    assert "保持原序" in items[0]["note"]
    assert items[0]["rule"] == "citation.reference_list.sort"


def test_alphabetical_scheme_sorts_by_full_text():
    """alphabetical 按全文键名排：剥掉 [n] 后整条比。"""
    e_adams = {"authors": ["Adams J"], "year": "2019", "lang": "en",
               "raw": "[2] Adams J. Beta[J]. Journal B, 2019, 1(1): 1-2."}
    e_wang = {"authors": ["Wang H"], "year": "2020", "lang": "en",
              "raw": "[1] Wang H. Alpha[J]. Journal A, 2020, 1(1): 1-2."}
    assert rf.sort_key(e_adams, "alphabetical")[0] < rf.sort_key(e_wang, "alphabetical")[0]


def test_cn_alphabet_scheme_uses_name_codepoint_order():
    """cn-first-by-alphabet 类：中文组(0)在前、组内按码位字母序；英文组(1)在后。"""
    e_zhou = {"authors": ["周一"], "year": "2021", "lang": "cn", "raw": REF_ZHOU}
    e_zhang = {"authors": ["张三"], "year": "2020", "lang": "cn", "raw": REF_ZHANG}
    e_en = {"authors": ["Adams J"], "year": "2019", "lang": "en", "raw": ""}
    keys = [rf.sort_key(e, "cn-first-by-alphabet-then-en-by-alphabet")
            for e in (e_zhou, e_zhang, e_en)]
    assert keys[0] < keys[1]               # 周 U+5468 < 张 U+5F20
    assert keys[0][0] == 0 and keys[2][0] == 1


# ---------------------------------------------------------------- F23

def test_mixed_reference_list_aborts_and_names_entries(tmp_path):
    """文献表混着无编号条目 → 整批不转，审计点名该段落。"""
    stray = "张三，李四. 无编号条目的研究[J]. 经济研究, 2019, 1(1): 1-5."
    src = make_src(tmp_path, extra_refs=[stray])
    rp = citation_rules(tmp_path, name="mixed.json")
    r, out = parse_render(tmp_path, src, rp)
    assert r.returncode == 0, r.stderr[-500:]

    joined = "\n".join(out_texts(out))
    assert "[1]" in joined and "[2]" in joined, "fail-closed：混合文献表仍被改了正文"
    assert "张三. 数字金融与企业融资[J]" in joined, "条目不应被模板重写"

    items = audit_items(out, "引用体系转换（已跳过）")
    assert items, "整批放弃必须留审计"
    assert "不带 [n] 编号" in items[0]["note"]
    assert "无编号条目的研究" in items[0]["note"], "要点名这些段落"


def test_footnote_citation_aborts_and_names_footnote(tmp_path):
    """脚注里的编号引用 → 整批不转，审计点名脚注。"""
    from test_footnote import build_docx_with_footnote

    src = build_docx_with_footnote(tmp_path / "fn.docx", footnote_text="方法设计参见[1]。")
    doc = Document(str(src))
    doc.add_paragraph("参考文献")
    doc.add_paragraph(REF_ZHOU)
    doc.add_paragraph(REF_ZHANG)
    doc.save(str(src))

    rp = citation_rules(tmp_path, name="fn.json")
    r, out = parse_render(tmp_path, src, rp)
    assert r.returncode == 0, r.stderr[-500:]

    joined = "\n".join(out_texts(out))
    assert "[1]" in joined, "fail-closed：脚注有编号引用时正文不得转换"

    items = audit_items(out, "引用体系转换（已跳过）")
    assert items and "脚注" in items[0]["note"] and "[1]" in items[0]["note"]


def test_unnumbered_helper_and_footnote_hits():
    """citation_ref 侧的两个判定器：口径与正文 classify_spans 一致。"""
    assert cr.unnumbered_entries([REF_ZHOU.strip(), "纯文字段落"]) == ["纯文字段落"]
    assert cr.unnumbered_entries([""]) == []
    # 数量语境的样文本不算脚注引用，不误杀转换
    assert cr.footnote_citation_hits(["样本取值区间为[90, 100]家企业。"]) == []
    assert len(cr.footnote_citation_hits(["方法参见[2]。"])) == 1


# ---------------------------------------------------------------- F24

NS_FULL = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')

TEXTBOX_TEXT = "文本框里的字"
SDT_BODY_TEXT = "这是内容控件里的正文段落。"


def build_sdt_textbox_doc(path):
    """内容控件包正文段 + AlternateContent 双份同文文本框 + 空内容控件。"""
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    body_p = doc.add_paragraph(SDT_BODY_TEXT)
    tb_p = doc.add_paragraph("正常段落，后面跟一个文本框。")
    textbox = (f'<w:r {NS_FULL}><mc:AlternateContent>'
               f'<mc:Choice Requires="v">'
               f'<w:pict><v:shape style="width:80pt;height:30pt"><v:textbox>'
               f'<w:txbxContent><w:p><w:r><w:t>{TEXTBOX_TEXT}</w:t></w:r></w:p>'
               f'</w:txbxContent></v:textbox></v:shape></w:pict></mc:Choice>'
               f'<mc:Fallback>'
               f'<w:pict><v:shape style="width:80pt;height:30pt"><v:textbox>'
               f'<w:txbxContent><w:p><w:r><w:t>{TEXTBOX_TEXT}</w:t></w:r></w:p>'
               f'</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>'
               f'</mc:AlternateContent></w:r>')
    tb_p._p.append(parse_xml(textbox))

    sdt = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
                    f'<w:sdtContent/></w:sdt>')
    body_p._p.addprevious(sdt)
    sdt.find(qn("w:sdtContent")).append(body_p._p)     # 段落搬进内容控件
    tb_p._p.addprevious(parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr/>'
                                  f'<w:sdtContent/></w:sdt>'))   # 无可处理内容
    doc.save(str(path))
    return path


def test_sdt_paragraph_enters_ir_and_typeset(tmp_path):
    """F24：sdt 内段落进 IR（带 sdt_pos）、参与套版；空内容控件记数。"""
    src = build_sdt_textbox_doc(tmp_path / "sdt.docx")
    ir = src.with_suffix(".ir.json")
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(src),
                    "-o", str(ir)], capture_output=True, text=True, check=True)
    parsed = json.loads(ir.read_text(encoding="utf-8"))

    assert parsed.get("unsupported_content_blocks") == 1, "空内容控件要记数进 IR"
    sdt_blocks = [b for b in parsed["blocks"] if "sdt_pos" in b]
    assert len(sdt_blocks) == 1 and sdt_blocks[0]["sdt_pos"] == 0
    assert sdt_blocks[0]["text"] == SDT_BODY_TEXT
    assert sdt_blocks[0]["type"] == "body", "sdt 内段落不参与前置区位置推断"
    assert all(TEXTBOX_TEXT not in b.get("text", "") for b in parsed["blocks"]), \
        "文本框文字不得混进任何块文本"

    rp = load_rules(tmp_path)
    r, out = parse_render(tmp_path, src, rp, citation=False)
    assert r.returncode == 0, r.stderr[-500:]
    assert "未处理内容块" in r.stdout, "渲染汇总要报「N 个未处理内容块」"

    slot = sdt_blocks[0]["idx"]
    ids = {i["id"] for i in audit_items(out)}
    assert f"{slot}.0|正文套版" in ids, "sdt 内段落没有套版审计记录"

    # 成品：sdt 内段落真的被套了正文字号（tjyj 正文 10.5pt → w:sz 21）
    doc = Document(str(out))
    szs = [e.get(qn("w:val")) for child in doc.element.body
           if child.tag == qn("w:sdt") for e in child.iter(qn("w:sz"))]
    assert "21" in szs, f"sdt 内段落没被套版，实际字号 {szs}"


def test_textbox_text_stays_in_object_and_out_of_blocks(tmp_path):
    """F24：文本框文字不出现在任何块文本里，Choice/Fallback 双份原样保留。"""
    src = build_sdt_textbox_doc(tmp_path / "tb.docx")
    rp = load_rules(tmp_path)
    r, out = parse_render(tmp_path, src, rp, citation=False)
    assert r.returncode == 0, r.stderr[-500:]

    assert all(TEXTBOX_TEXT not in t for t in out_texts(out)), \
        "文本框文字混进了某个段落"
    for p in (src, out):     # 源稿与成品都原样保留（Choice + Fallback 各一份）
        raw = zipfile.ZipFile(str(p)).read("word/document.xml").decode("utf-8")
        assert raw.count(TEXTBOX_TEXT) == 2, f"{p.name} 里文本框对象被改动了"


def test_sdt_note_reorder_survives_elem_lookup(tmp_path):
    """sdt 内的表注也走 elem_of 搬运：重排不崩、层序真的换了。

    修.elem_of 迁移漏网：删除循环按 elem_of 取 sdt 内块，搬运循环却
    仍用 elems[idx] 直取——sdt 槽位是列表，搬一条 sdt 内的表注就
    AttributeError。夹具：普通表格 + sdt 包「资料来源→注」（与声明层序
    相反，必须触发重排）。
    """
    src = tmp_path / "notes.docx"
    doc = Document()
    doc.add_paragraph("数字金融发展与企业融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_table(rows=2, cols=2).cell(0, 0).paragraphs[0].add_run("指标")
    sdt = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr/>'
                    f'<w:sdtContent>'
                    f'<w:p><w:r><w:t>资料来源：某统计年鉴。</w:t></w:r></w:p>'
                    f'<w:p><w:r><w:t>注：显著性水平为 5%。</w:t></w:r></w:p>'
                    f'</w:sdtContent></w:sdt>')
    doc.element.body.append(sdt)
    doc.save(str(src))

    def mutate(rules):
        rules["tables"]["note_layers"] = [{"kind": "general"}, {"kind": "source"}]
    rp = load_rules(tmp_path, name="layers.json", mutate=mutate)

    ir = src.with_suffix(".ir.json")
    subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(src),
                    "-o", str(ir)], capture_output=True, text=True, check=True)
    r = subprocess.run(
        [sys.executable, str(TOOLS / "render_from_ir.py"), str(ir), str(rp),
         "-o", str(src.with_name("out.docx"))], capture_output=True, text=True)
    assert r.returncode == 0, f"表注重排崩了：{r.stderr[-400:]}"

    # 成品里「注」必须排在「资料来源」前面（声明的层序 general → source）
    texts = [p.text for p in Document(str(src.with_name("out.docx"))).paragraphs]
    assert texts.index(next(t for t in texts if t.startswith("注："))) < \
        texts.index(next(t for t in texts if t.startswith("资料来源"))), \
        f"层序没有按声明重排：{texts}"


# ---------------------------------------------------------------- F30

def test_subtitle_default_is_marked_presumed(tmp_path):
    """F30：删 subtitle 规则的刊 → 审计条目带 ~ 与「缺省值」说明（黑体 16）。

    说明用「缺省值」不用「推定值」：推定值专指规则集标了 measured:false 的
    取值（test_audit_traceable 守着这条线），代码缺省自报家门但不冒称推定。
    """
    from test_manuscript_matrix import build_messy_manuscript

    src = build_messy_manuscript(tmp_path / "m.docx")
    rp = load_rules(tmp_path, "cjyj")                             # cjyj 无 subtitle
    r, out = parse_render(tmp_path, src, rp, citation=False)
    assert r.returncode == 0, r.stderr[-500:]

    items = audit_items(out, "副标题套版")
    assert items, "副标题套版条目缺失"
    it = items[0]
    assert it["rule"].endswith("~"), f"用到缺省值的条目应带 ~ 标注：{it['rule']}"
    assert "缺省值" in it["note"], it["note"]
    assert "推定" not in it["note"], f"代码缺省不得冒称推定：{it['note']}"
    assert "16" in it["note"] and "黑体" in it["note"], it["note"]


def test_subtitle_declared_values_unchanged(tmp_path):
    """F30：规则集有值时数值不变（tjyj 副标题 26pt 落到成品）。"""
    from test_manuscript_matrix import build_messy_manuscript

    src = build_messy_manuscript(tmp_path / "m.docx")
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    declared = rules["typography"]["subtitle"]["size_pt"]
    rp = tmp_path / "declared.json"
    rp.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    r, out = parse_render(tmp_path, src, rp, citation=False)
    assert r.returncode == 0, r.stderr[-500:]

    items = audit_items(out, "副标题套版")
    assert items and f"{declared:g}pt" in items[0]["note"], \
        f"声明值 {declared} 被改了：{items[0]['note'] if items else '缺条目'}"
    sub = next(p for p in Document(str(out)).paragraphs if p.text.startswith("——"))
    sizes = [run.font.size.pt for run in sub.runs if run.font.size]
    assert sizes and sizes[0] == declared, f"成品字号 {sizes} ≠ 声明值 {declared}"


def test_body_line_default_marked_only_when_fallback(tmp_path):
    """F30：正文行距缺省 20pt 只在规则集没给值时标注。

    用乱排稿件做夹具——它有真正的正文块（make_src 的短稿前三段会被
    归为前置区，没有正文套版条目可比）。"""
    from test_manuscript_matrix import build_messy_manuscript

    src = build_messy_manuscript(tmp_path / "m.docx")

    def no_line_pt(rules):
        rules["typography"]["body"]["line_spacing"].pop("value_pt", None)
        rules["typography"]["body"]["line_spacing"]["equivalent_exact_pt"] = None
    rp = load_rules(tmp_path, name="nolinespacing.json", mutate=no_line_pt)
    r, out = parse_render(tmp_path, src, rp, citation=False)
    assert r.returncode == 0, r.stderr[-500:]

    items = audit_items(out, "正文套版")
    assert items, "正文套版条目缺失"
    assert all("缺省值" in i["note"] and "20" in i["note"] for i in items), \
        "行距落到缺省值 20pt 时必须标出「缺省值」"
    assert all(i["rule"].endswith("~") for i in items), \
        f"用缺省值的条目 rule 应带 ~：{items[0]['rule']}"

    # 对照组：规则集有值时不标缺省（基准行为不变）
    src2 = build_messy_manuscript(tmp_path / "m2.docx")
    rp2 = load_rules(tmp_path, name="plain.json")
    r2, out2 = parse_render(tmp_path, src2, rp2, citation=False)
    assert r2.returncode == 0, r2.stderr[-500:]
    items2 = audit_items(out2, "正文套版")
    assert items2 and all("缺省值" not in i["note"] for i in items2), \
        "规则集已声明行距时不应标缺省值"
