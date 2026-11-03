"""Word 域：文献管理软件插入的文字不归本工具改。

EndNote / Zotero / NoteExpress 插入的引文与文献表，在 docx 里是 Word 域。
可见文字是域的「结果」，不是作者敲的 run。而解析层只收 `w:t`——指令文本
`w:instrText` 天然读不到——于是域在这条管道里是**隐形**的：IR 里只有纯文本，
重建时域结构凭空蒸发，全流程没有一处留下痕迹。

这条链断掉的后果不是「格式不对」，而是「返修时改不了」：作者想补一篇文献，
文献库已经找不到对应关系了。所以这里守三件事：

1. 解析层认得出来，并把统计写进 IR（不许静默）
2. 域块里的文字不参与任何改写（改了也守不住，F9 一更新就回滚）
3. 报告里明确写出「这里看过、但没动」，且不计入改动数
"""

import json

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from conftest import NS_W, minimal_doc, mutated_rules, render_into, run_tool

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ENDNOTE_CITE = (" ADDIN EN.CITE <EndNote><Cite><Author>张三</Author>"
                "<Year>2020</Year><RecNum>12</RecNum></Cite></EndNote> ")
ZOTERO_CITE = (' ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"abc","properties":{}} ')


def _elem(tag, text=None, **attrs):
    n = OxmlElement(tag)
    for k, v in attrs.items():
        n.set(qn("w:" + k), v)
    if text is not None:
        n.text = text
    return n


def _run(*children):
    r = OxmlElement("w:r")
    for c in children:
        r.append(c)
    return r


def add_complex_field(par, instr, result):
    """在段落末尾追加一个完整的复杂域：begin → instrText → separate → 结果 → end。"""
    it = _elem("w:instrText", instr)
    it.set(XML_SPACE, "preserve")
    for node in (
        _run(_elem("w:fldChar", None, fldCharType="begin")),
        _run(it),
        _run(_elem("w:fldChar", None, fldCharType="separate")),
        _run(_elem("w:t", result)),
        _run(_elem("w:fldChar", None, fldCharType="end")),
    ):
        par._p.append(node)


def add_simple_field(par, instr, result):
    """简化域 w:fldSimple：指令挂在属性上，没有三段结构。"""
    fld = _elem("w:fldSimple", None, instr=instr)
    fld.append(_run(_elem("w:t", result)))
    par._p.append(fld)


def build_manuscript(path, with_field):
    """构造一份最小稿件：正文一段 + 参考文献表一条。

    with_field=False 的对照组除了没有域，其余逐字相同——两组审计的差异
    只能来自域本身，不能来自稿件内容。
    """
    doc = minimal_doc(path, title="企业创新的影响因素研究",
                      abstract="内容提要：本文考察企业创新的影响因素。",
                      keywords="关键词：企业创新；研发投入")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("已有研究对企业创新的驱动因素做过大量讨论。")
    doc.add_paragraph("参考文献")

    entry = doc.add_paragraph("[1] 张三. 2020. 企业创新的驱动因素[J]. 管理世界, (3): 45-58.")
    if with_field:
        add_complex_field(entry, ENDNOTE_CITE,
                          "[1] 张三. 2020. 企业创新的驱动因素[J]. 管理世界, (3): 45-58.")
        # 域结果才是可见文字，上面那句是占位；清掉它模拟真实结构
        entry._p.remove(entry._p.findall(qn("w:r"))[0])

    doc.save(str(path))
    return path


def parse_ir(docx_path, workdir):
    out = workdir / "m.ir.json"
    run_tool("parse_docx.py", docx_path, "-o", out)
    return json.loads(out.read_text(encoding="utf-8"))


# ---------------------------------------------------------------- 解析层

def test_scan_fields_reads_endnote_citation(tmp_path):
    """EndNote 复杂域要认成 citation，且读出指令与域结果。"""
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
    import parse_docx as pdx

    doc = minimal_doc(tmp_path / "a.docx")
    p = doc.add_paragraph("已有研究表明")
    add_complex_field(p, ENDNOTE_CITE, "（张三，2020）")

    fields = pdx.scan_fields(p._p)
    assert len(fields) == 1, fields
    f = fields[0]
    assert f["kind"] == "citation"
    assert "ADDIN EN.CITE" in f["instr"]
    assert f["result"] == "（张三，2020）"


def test_scan_fields_covers_zotero_and_simple(tmp_path):
    """Zotero 走 ADDIN ZOTERO_ITEM；fldSimple 也要认。"""
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
    import parse_docx as pdx

    doc = minimal_doc(tmp_path / "b.docx")
    p = doc.add_paragraph("甲")
    add_complex_field(p, ZOTERO_CITE, "(Li, 2021)")
    q = doc.add_paragraph("乙")
    add_simple_field(q, " PAGE ", "3")

    assert [f["kind"] for f in pdx.scan_fields(p._p)] == ["citation"]
    assert [f["kind"] for f in pdx.scan_fields(q._p)] == ["page"]


def test_scan_fields_ignores_plain_text(tmp_path):
    """没有域的段落不能凭空造出域。"""
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
    import parse_docx as pdx

    doc = minimal_doc(tmp_path / "c.docx")
    p = doc.add_paragraph("普通正文，没有任何域。")
    assert pdx.scan_fields(p._p) == []


def test_ir_reports_field_stats(tmp_path):
    """域统计必须进 IR：不记录的话，「文献表被拆成死文本」全流程无人知晓。"""
    src = build_manuscript(tmp_path / "f.docx", with_field=True)
    ir = parse_ir(src, tmp_path)

    fs = ir.get("fields") or {}
    assert fs.get("total") == 1, fs
    assert fs["by_kind"].get("citation") == 1
    assert fs["blocks_with_fields"] == 1

    refs = [b for b in ir["blocks"] if b["type"] == "reference_item"]
    assert refs and refs[0].get("fields"), "文献条目块上应挂着域明细"


def test_ir_without_fields_has_no_key(tmp_path):
    """无域稿件不写 fields 键，免得每份 IR 都多一个空对象。"""
    src = build_manuscript(tmp_path / "g.docx", with_field=False)
    assert "fields" not in parse_ir(src, tmp_path)


# ---------------------------------------------------------------- 渲染层

def test_field_block_text_is_not_rewritten(tmp_path):
    """含域的文献条目不改编号：改了 F9 一更新就回滚，白改还留隐患。

    对照组（无域）必须被改——否则这条测试可能只是「什么都没发生」。
    """
    rules = mutated_rules(tmp_path, "tjyj", **{"citation.reference_list.number_format": "〔1〕"})

    plain = build_manuscript(tmp_path / "plain.docx", with_field=False)
    pipe = render_into(tmp_path / "w-plain", plain, target="tjyj", rules_path=rules)
    acts = [it["action"] for it in pipe.audit["items"]]
    assert "文献条目编号归一" in acts, acts

    locked = build_manuscript(tmp_path / "locked.docx", with_field=True)
    pipe2 = render_into(tmp_path / "w-locked", locked, target="tjyj", rules_path=rules)
    acts2 = [it["action"] for it in pipe2.audit["items"]]
    assert "文献条目编号归一" not in acts2, acts2
    assert "含域，文字未改" in acts2, acts2

    text = " ".join(p.text for p in Document(str(pipe2.docx)).paragraphs)
    assert "[1] 张三" in text, "域结果必须原样保留"
    assert "〔1〕" not in text, "含域的条目不该被改编号"


def test_locked_note_is_not_counted_as_change(tmp_path):
    """说明条目必须进报告、但不计入改动数——否则报告自己就在虚报。"""
    rules = mutated_rules(tmp_path, "tjyj", **{"citation.reference_list.number_format": "〔1〕"})
    src = build_manuscript(tmp_path / "h.docx", with_field=True)
    pipe = render_into(tmp_path / "w-count", src, target="tjyj", rules_path=rules)

    notes = [it for it in pipe.audit["items"] if it["action"] == "含域，文字未改"]
    assert notes, "锁定必须留下说明条目，藏起来等于当没发生过"
    assert all(it["changed"] is False for it in notes)

    changed_blocks = pipe.audit.get("changed_blocks")
    assert changed_blocks is not None, "审计要报改动块数，报告才谈得上覆盖率"
    # 同一块可以既出「文字未改」（N 档）又出「参考文献条目套版」（A 档）：
    # 版式照改、文字不动，这正是设计意图。要守的是该块没有出文字改动条目。
    locked_idx = {str(it["idx"]) for it in notes}
    text_changes = {str(it["idx"]) for it in pipe.audit["items"]
                    if it["action"] in ("文献条目编号归一", "条目按目标刊体例重排")}
    assert not (locked_idx & text_changes), locked_idx & text_changes


def test_audit_reports_coverage(tmp_path):
    """报告要能回答「没改的块是看过还是漏了」。"""
    plain = build_manuscript(tmp_path / "i.docx", with_field=False)
    pipe = render_into(tmp_path / "w-cov", plain, target="tjyj")
    assert pipe.audit["checked_blocks"] > 0
    assert pipe.audit["checked_blocks"] >= pipe.audit["changed_blocks"]


def test_rerun_on_own_output_is_idempotent(tmp_path):
    """把成品再喂回来跑同刊，审计里不该再有改动。

    与 test_idempotent 的「同一输入渲染两次」不是一回事：那条守的是确定性
    （同进同出），这条守的是**自洽**——渲染器写出的东西，自己的解析与重排
    逻辑读得回来吗。实测过两个真 bug 都卡在这里：

      《经济管理》二级标题写作「1.」，硬编码成三级后被降级保护打成正文，
      整篇二级标题套成正文样式；剥编号的正则只认「1.1」不认「1.」，
      重跑一次叠一层，`1.标题` → `1.1.标题`。

    所以这条不只是「条目数应为 0」，它守的是「每一条都对应一次真实改动」——
    审计报告可信度的底线。覆盖全部入库刊：编号体系不同的刊各有各的坑。
    """
    from conftest import tracked_rulesets

    journals = sorted(tracked_rulesets() or ["tjyj"])
    for jid in journals:
        src = build_manuscript(tmp_path / f"k-{jid}.docx", with_field=False)
        first = render_into(tmp_path / f"idem1-{jid}", src, target=jid)
        assert first.audit["items"], f"{jid}：首轮应当有改动，否则测试没有起点"

        second = render_into(tmp_path / f"idem2-{jid}", first.docx, target=jid)
        leftover = [it for it in second.audit["items"] if it.get("changed", True)]
        sample = "、".join(f"{it['action']}@{it['idx']}" for it in leftover[:5])
        assert leftover == [], (
            f"{jid}：成品重跑仍报 {len(leftover)} 条改动，说明渲染不幂等：{sample}")
