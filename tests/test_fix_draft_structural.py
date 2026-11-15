"""建库复核页的结构开关（2026-11-13）。

匿名、作者信息页、区块必备/可选、声明必备/可选原先只能手改 draft.json。
现在复核页有可视开关，保存仍走 apply_manual 同一事务：改动计入
verification.extraction.manual，并从「沿用基线」披露里移出。
"""

import sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parent.parent / "tools"
sys.path.insert(0, str(TOOLS))

import extract_journal as ej  # noqa: E402


def _rules():
    base, _bid = ej.baseline_rules()
    # 深拷贝一份并造出可开关的结构
    rules = ej.deepcopy(base)
    rules["submission"] = dict(rules.get("submission") or {})
    rules["submission"]["anonymous"] = False
    rules["statements"] = {"conflict_of_interest": {"required": True, "note": ""},
                           "ai_declaration": {"required": False, "note": ""}}
    return rules


def test_structural_fields_lists_blocks_and_statements():
    s = ej.structural_fields(_rules())
    assert s["anonymous"] is False
    assert s["blocks"], "区块清单为空"
    assert any(b["type"] == "abstract" for b in s["blocks"])
    assert {x["key"] for x in s["statements"]} == {"conflict_of_interest",
                                                   "ai_declaration"}


def test_toggle_structural_values_and_record_manual():
    rules = _rules()
    s = ej.structural_fields(rules)
    abstract = next(b for b in s["blocks"] if b["type"] == "abstract")
    out, report, errors = ej.apply_manual(
        rules, {},
        structural={"anonymous": True, "author_page": False,
                    "blocks": {abstract["index"]: True},
                    "statements": {"conflict_of_interest": False,
                                   "ai_declaration": True}})
    assert errors == []
    assert out["submission"]["anonymous"] is True
    assert out["blocks"][abstract["index"]]["required"] is True
    assert out["statements"]["conflict_of_interest"]["required"] is False
    assert out["statements"]["ai_declaration"]["required"] is True

    manual = out["verification"]["extraction"]["manual"]
    assert "structural:anonymous" in manual
    assert "structural:statements:ai_declaration" in manual
    assert manual["structural:anonymous"]["path"] == ["submission", "anonymous"]

    # 继承披露重算后，被改过的路径不能继续以「沿用基线值」出现
    inherited = out["verification"]["extraction"]["inherited"]
    assert not any(it["path"] if "path" in it else False for it in inherited)


def test_structural_save_is_idempotent():
    rules = _rules()
    s = ej.structural_fields(rules)
    structural = {"anonymous": True, "author_page": False,
                  "blocks": {b["index"]: b["required"] for b in s["blocks"]},
                  "statements": {x["key"]: x["required"] for x in s["statements"]}}
    once, _r, _e = ej.apply_manual(rules, {}, structural=structural)
    twice, _r2, _e2 = ej.apply_manual(once, {}, structural=structural)
    assert twice["verification"]["extraction"]["manual"] == \
        once["verification"]["extraction"]["manual"], \
        "同样的结构开关保存两次又新增了手工修正记录"


def test_manual_paths_includes_structural_paths():
    manual = {"structural:blocks:3": {"label": "摘要·必备", "value": True,
                                      "previous": False,
                                      "path": ["blocks", 3, "required"]}}
    assert ("blocks", 3, "required") in ej.manual_paths(manual)


def test_english_section_toggle_moves_the_group():
    """英文区放正文前还是文末，由复核页开关决定（不再手改 draft.json）。"""
    rules = _rules()
    types = [b.get("type") for b in rules["blocks"]]
    assert ej.structural_fields(rules)["english_section"] == "back"

    out, _report, errors = ej.apply_manual(
        rules, {}, structural={"english_section": "front"})
    assert errors == []
    after = [b.get("type") for b in out["blocks"]]
    body_at = after.index("body")
    en_at = [i for i, b in enumerate(out["blocks"]) if b.get("lang") == "en"]
    assert en_at and max(en_at) < body_at, "英文区没有整组移到正文前"
    assert "structural:english_section" in \
        out["verification"]["extraction"]["manual"]

    # 再存一次同一选择：区块顺序与手工修正记录都不再变化（幂等）
    twice, _r2, _e2 = ej.apply_manual(
        out, {}, structural={"english_section": "front"})
    assert [b.get("type") for b in twice["blocks"]] == after
    assert twice["verification"]["extraction"]["manual"]["structural:english_section"] == \
        out["verification"]["extraction"]["manual"]["structural:english_section"]


def test_english_section_back_places_after_references():
    rules = _rules()
    out, _report, errors = ej.apply_manual(
        rules, {}, structural={"english_section": "back"})
    assert errors == []
    after = [b.get("type") for b in out["blocks"]]
    assert max(i for i, t in enumerate(after) if t == "abstract") > \
        after.index("references")
