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
