"""第二承诺：交付必带审计，且每条都能被追到底。

「无报告视为交付失败」是 DESIGN §6 的硬判定。这里守的是报告本身的质量：
条目唯一、指向真实存在的块、改动有据可查。
"""

import json

import pytest

from conftest import RULES, TARGET_JOURNAL


def test_audit_is_always_produced(baseline):
    """成品必须伴随审计文件，且不是空壳。"""
    assert baseline.docx.exists()
    assert baseline.items, "审计条目为空：报告存在但没内容，等同于失败"


def test_item_ids_are_unique(baseline):
    ids = [it["id"] for it in baseline.items]
    assert len(ids) == len(set(ids)), "审计 id 重复，驳回时会误伤"


def test_every_item_has_action_and_reason(baseline):
    """每条都得说清「做了什么」和「改成什么」。

    说明可以在 note 里，也可以直接放在 before/after（如悬挂缩进记「2.3 字符」），
    但三者不能全空——那就成了一条什么也没说的记录。
    """
    bad = [it for it in baseline.items
           if not (it.get("action") or "").strip()
           or not (it.get("note") or it.get("before") or it.get("after"))]
    assert not bad, f"条目既没 note 也没 before/after：{bad[:3]}"


def test_items_point_to_real_blocks(baseline):
    """条目指向的块必须在源稿里存在（-1 是页面设置，不对应任何块）。"""
    known = set(baseline.source_texts) | {-1}
    ghosts = sorted({it["idx"] for it in baseline.items} - known)
    assert not ghosts, f"审计指向不存在的块：{ghosts}"


def test_audit_records_source_and_target(baseline):
    """报告要能说明「谁被排成了哪本刊」，否则事后没法复核。"""
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    assert baseline.audit.get("source")
    assert baseline.audit.get("target") == rules["journal"]["name_zh"]


@pytest.mark.xfail(strict=False, reason=(
    "DESIGN §6/§8 要求每条 B 档改动带「规则 ID」（如 N-03）并可回溯到规则集。"
    "当前审计项只有中文 action（如「标签归一」），没有机器可读的规则 ID。"
    "这条差距由本用例记录，修好后会自动转为 XPASS。"))
def test_b_tier_items_carry_a_rule_id(baseline):
    """B 档改动要能回溯到规则集里的具体条目，而不只是一个中文动作名。"""
    missing = [it["id"] for it in baseline.text_editing_items
               if not (it.get("rule") or it.get("rule_id"))]
    assert not missing, f"这些改动没有规则 ID：{missing[:5]}"
