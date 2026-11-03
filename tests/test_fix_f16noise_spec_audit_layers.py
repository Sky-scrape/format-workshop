"""F16NOISE 回炉：规格页与审计页的推定项展示分层。

批五 F16 把继承披露扩到全量后，AI 建库草稿的 assumptions 从几条涨到 48+ 条。
复核页（journal_draft.html）已分组折叠，但独立验证证实规格页（render_spec.py）
与审计页（render_audit.py）仍是平铺 <ul>。本文件钉住回炉后的两层展示：

- 规格：≥12 条按区块 <details> 分组折叠、默认收起、计数徽标、逐条可达；
  少量推定项（手工规则集常态）保持平铺。
- 审计：render_from_ir 往 audit JSON 写 assumption_sections（新键，与
  assumptions 同源同口径），审计页据此分组折叠；旧 JSON 无键时兜底自算；
  assumptions 数组本身逐条原样落盘（披露不删）。
- 共享对号逻辑 tools/assume_groups.py：同输入同输出（承诺③）；规则集
  留痕（verification.extraction.inherited）是唯一分组依据，不内置字段表
  （承诺④）。

夹具姿势照 tests/test_fix_security_pages.py（render_spec/render_audit 走
run_tool 子进程）与 tests/test_citation_convert.py（render_from_ir 子进程，
python-docx 造最小稿）。
"""

import json
import sys
from pathlib import Path

from docx import Document

from conftest import RULES, TOOLS, minimal_doc, run_tool

sys.path.insert(0, str(TOOLS))
import assume_groups as ag  # noqa: E402


def _inherited(n=45):
    """仿 extract_journal 的 inherited 留痕：typography/limits 两个区块。"""
    items = []
    for i in range(30):
        items.append({"key": f"typography.body.f{i}", "label": f"正文字段{i}",
                      "value": i, "section": "typography", "fact": True})
    for i in range(n - 30):
        items.append({"key": f"limits.f{i}", "label": f"限制字段{i}",
                      "value": i, "section": "limits", "fact": False})
    return items


def _derived(inherited):
    """按 extract_journal 的生成口径造派生假设：'label = value（尾注）'。"""
    return [f"{it['label']} = {it['value']}（未从稿约抽到，沿用基线值，须人工核对）"
            for it in inherited]


def _rules_with_draft_assumptions(tmp_path, n=45, name="draft_rules.json"):
    """带 AI 建库留痕的规则集：assumptions 全量 + extraction.inherited。"""
    inherited = _inherited(n)
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["assumptions"] = _derived(inherited) + [
        "另有 3 个结构字段（blocks 区块清单、statements、personal_fields、"
        "tables 边框等）未经稿约逐项抽取，逐字段沿用基线《tjyj》原值；"
        "逐字段清单见 verification.extraction.inherited，发布前须人工核对"]
    rules["verification"]["extraction"] = {"inherited": inherited}
    p = tmp_path / name
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return p


def _spec_html(tmp_path, rules_path, out="spec.html"):
    run_tool("render_spec.py", rules_path, "-o", tmp_path / out)
    return (tmp_path / out).read_text(encoding="utf-8")


def _audit_html(audit, tmp_path, out="audit.html", job=None):
    p = tmp_path / "audit.json"
    p.write_text(json.dumps(audit, ensure_ascii=False), encoding="utf-8")
    args = [p, "-o", tmp_path / out]
    if job:
        args += ["--job", job]
    run_tool("render_audit.py", *args)
    return (tmp_path / out).read_text(encoding="utf-8")


def _minimal_audit(assumptions, sections=None):
    d = {"target": "统计研究", "source": "manuscript.docx",
         "items": [], "assumptions": assumptions}
    if sections is not None:
        d["assumption_sections"] = sections
    return d


# ---------------------------------------------------------------- 共享对号

def test_group_assumptions_maps_by_inherited_labels():
    """派生假设按 inherited 的 label 对回区块；对不上号进「其他推定」。"""
    inherited = _inherited()
    assumptions = _derived(inherited) + ["手写假设：表注层序按惯例处理"]
    groups = ag.group_assumptions(assumptions, {"inherited": inherited})
    names = [s for s, _ in groups]
    assert "typography" in names and "limits" in names and ag.OTHER in names
    d = dict(groups)
    assert len(d["typography"]) == 30 and len(d["limits"]) == 15
    assert d[ag.OTHER] == ["手写假设：表注层序按惯例处理"]
    # 逐条原文保留：分组不改一字（承诺①④）
    assert sum((["".join(rows)] and len(rows)) for rows in d.values()) == len(assumptions)
    assert "".join(d["typography"][0]) == assumptions[0]


def test_group_assumptions_is_pure():
    """同一输入两次结果相同（承诺③），且不改入参。"""
    inherited = _inherited()
    assumptions = _derived(inherited)
    snap = list(assumptions)
    g1 = ag.group_assumptions(assumptions, {"inherited": inherited})
    g2 = ag.group_assumptions(assumptions, {"inherited": inherited})
    assert g1 == g2 and assumptions == snap


def test_group_assumptions_without_extraction_all_to_other():
    """手工规则集无 extraction 留痕：全部进「其他推定」，一条不丢。"""
    assumptions = ["推定甲", "推定乙"]
    groups = ag.group_assumptions(assumptions, None)
    assert dict(groups) == {ag.OTHER: assumptions}


def test_section_map_roundtrip():
    """section_map 与 group_assumptions 同口径：每条都能查到组名。"""
    inherited = _inherited()
    assumptions = _derived(inherited)
    m = ag.section_map(assumptions, {"inherited": inherited})
    assert len(m) == len(assumptions)
    assert m[assumptions[0]] == "typography"
    assert m[assumptions[-1]] == "limits"


# ---------------------------------------------------------------- 规格页

def test_spec_page_groups_big_assumptions(tmp_path):
    """复现 F16NOISE③：46 条草稿假设在规格页按区块折叠、默认收起、徽标、逐条可达。"""
    p = _rules_with_draft_assumptions(tmp_path)
    page = _spec_html(tmp_path, p)
    assert "推定项（官方未规定，非期刊要求）" in page, "推定项节标题不因分层而变"
    assert "共 46 条，按区块折叠" in page, "KPI 总数保留（45 派生 + 1 汇总）"
    assert page.count("<details>") == 3, "typography/limits/其他推定 三个分组，默认收起"
    assert "<details open" not in page, "大批量时分组默认收起"
    assert page.count('class="abadge"') == 3, "每组带计数徽标"
    assert "typography" in page and "limits" in page, "区块名随组名亮出"
    assert "其他推定" in page, "结构字段汇总条对不上号，进其他推定组"
    # 逐条仍可达：第一条与最后一条派生假设 + 汇总条都要在页面里
    for a in (_derived(_inherited())[0], _derived(_inherited())[-1]):
        assert a in page
    assert "逐字段清单见 verification.extraction.inherited" in page


def test_spec_page_small_assumptions_stay_flat(tmp_path):
    """手工规则集常态（<12 条）：平铺照旧，不引入折叠结构。"""
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    assert 0 < len(rules["assumptions"]) < ag.GROUP_THRESHOLD, \
        "夹具前提：tjyj 推定项少于阈值"
    p = tmp_path / "tjyj_copy.json"
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    page = _spec_html(tmp_path, p)
    assert "<details" not in page, "少量推定项不该折叠"
    assert rules["assumptions"][0] in page, "平铺逐条展示"


# ---------------------------------------------------------------- 审计页

def test_render_from_ir_writes_assumption_sections(tmp_path):
    """render_from_ir 的 audit JSON 带 assumption_sections，且 assumptions 原样。"""
    minimal_doc(tmp_path / "m.docx")
    run_tool("parse_docx.py", tmp_path / "m.docx", "-o", tmp_path / "m.ir.json")
    rules_path = _rules_with_draft_assumptions(tmp_path, name="r.json")
    out = tmp_path / "m.out.docx"
    run_tool("render_from_ir.py", tmp_path / "m.ir.json", rules_path, "-o", out)
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    assert audit["assumptions"] == json.loads(rules_path.read_text(encoding="utf-8"))["assumptions"], \
        "assumptions 数组必须逐条原样落盘（展示分层不删披露）"
    sec = audit["assumption_sections"]
    assert len(sec) == len(audit["assumptions"]), "映射与数组一一对应"
    derived = [a for a in audit["assumptions"] if a.endswith("（未从稿约抽到，沿用基线值，须人工核对）")]
    assert sec[derived[0]] == "typography" and sec[derived[-1]] == "limits"
    assert all(sec[a] == ag.OTHER for a in audit["assumptions"] if a not in derived), \
        "结构字段汇总条不在 inherited 里，应落其他推定"


def test_audit_page_groups_big_assumptions(tmp_path):
    """复现 F16NOISE④：审计页大批量推定项按区块折叠、默认收起、徽标、逐条可达。"""
    inherited = _inherited()
    assumptions = _derived(inherited)
    audit = _minimal_audit(assumptions, ag.section_map(assumptions, {"inherited": inherited}))
    page = _audit_html(audit, tmp_path)
    assert "本刊的推定项（45 条）" in page, "总数保留"
    assert "按区块折叠，逐条可达" in page
    assert page.count('<details class="agroup">') == 2, "typography/limits 两组"
    assert "<details open" not in page, "默认收起"
    assert "typography" in page and "limits" in page
    assert assumptions[0] in page and assumptions[-1] in page, "逐条仍可达"


def test_audit_page_small_assumptions_stay_flat(tmp_path):
    """手工规则集（3 条，tjyj 常态）：审计页平铺照旧——既有姿势不受影响。"""
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    page = _audit_html(_minimal_audit(rules["assumptions"]), tmp_path)
    assert '<details class="agroup">' not in page
    assert rules["assumptions"][0] in page, "平铺逐条展示"
    assert "本刊的推定项（3 条）" in page


def test_audit_page_falls_back_without_sections_key(tmp_path):
    """旧 JSON（render_from_ir 早期产物）没有 assumption_sections：兜底自算不崩。"""
    inherited = _inherited()
    assumptions = _derived(inherited)
    audit = _minimal_audit(assumptions, sections=None)
    page = _audit_html(audit, tmp_path)
    assert '<details class="agroup">' in page, "兜底口径仍能分组"
    assert assumptions[0] in page
    # 乱值也兜底：非 dict 的 assumption_sections 按缺键处理
    audit2 = _minimal_audit(assumptions, sections="garbage")
    page2 = _audit_html(audit2, tmp_path, out="audit2.html")
    assert '<details class="agroup">' in page2 and assumptions[0] in page2
