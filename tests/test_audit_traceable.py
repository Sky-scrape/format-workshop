"""第二承诺：交付必带审计，且每条都能被追到底。

「无报告视为交付失败」是 DESIGN §6 的硬判定。这里守的是报告本身的质量：
条目唯一、指向真实存在的块、改动有据可查。
"""

import json

import pytest

from conftest import RULES, TARGET_JOURNAL, render_into, run_sizes
from rules_ref import resolvable

SHIPPED_JOURNALS = ["cjyj", "tjyj", "gjjryj"]


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


def test_b_tier_items_carry_a_rule_id(baseline):
    """B 档改动（改文字的）必须带规则集引用，不能只有一个中文动作名。"""
    missing = [it["id"] for it in baseline.text_editing_items if not it.get("rule")]
    assert not missing, f"这些改动没有规则引用：{missing[:5]}"


def test_every_rule_reference_resolves(baseline):
    """审计里的引用必须真能落到规则集的一个值上。

    写个好看的名字不算回溯；`rules_ref.resolve` 指不到就抛 KeyError。
    带 `~` 的引用表示该键在部分期刊可缺省，此时只要求其父路径存在。
    """
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    broken = [(it["action"], it["rule"]) for it in baseline.items
              if it["rule"] and not resolvable(rules, it["rule"])]
    assert not broken, f"这些引用在规则集里指不到实处：{broken[:5]}"


@pytest.mark.parametrize("target", SHIPPED_JOURNALS)
def test_rule_refs_resolve_across_shipped_journals(sample_docx, tmp_path, target):
    """换几本刊跑，所有引用都要能落到各自的规则集里。

    单刊通过不代表全通过——不同规则集的键不一样，缺哪个键只有跑到了才知道。
    """
    p = render_into(tmp_path / target, sample_docx, target=target)
    rules = json.loads((RULES / f"{target}.json").read_text(encoding="utf-8"))

    broken = [(it["action"], it["rule"]) for it in p.items
              if it["rule"] and not resolvable(rules, it["rule"])]
    assert not broken, f"{target} 有不落地的引用：{broken[:5]}"


def test_body_restyling_is_audited(baseline):
    """正文字号变了，就必须有「正文套版」条目。

    这是原先的空白：标签、编号这些小事都记了，最显眼的正文重排反倒一条没有。
    """
    src = run_sizes(baseline.dir / "manuscript.docx")
    out = run_sizes(baseline.docx)
    body = {i for i, t in baseline.block_type.items() if t == "body"}
    changed = {i for i in body if src.get(i) != out.get(i)}
    assert changed, "正文没有发生变化，这条用例失去意义——检查夹具选刊"

    audited = {it["idx"] for it in baseline.items if it["action"] == "正文套版"}
    missing = sorted(changed - audited)
    assert not missing, f"这些正文段落改了字号却没有审计记录：{missing[:5]}"


def test_body_restyling_records_what_changed(baseline):
    """正文套版的说明要写清改成什么，否则用户没法核对。"""
    items = [it for it in baseline.items if it["action"] == "正文套版"]
    assert items, "没有任何正文套版条目"
    for it in items:
        assert "pt" in it["note"], f"条目没说清目标字号：{it['note']}"
        assert it["before"] == "" and it["after"] == "", "正文套版不该携带文字快照"


def test_audit_carries_ruleset_assumptions(baseline):
    """报告要带上本刊的推定项。

    规则集里写着「样本无副标题实例，按与主标题同号推定」，报告却只报 26pt，
    用户会以为是期刊规定。推定值必须自报家门。
    """
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    assert baseline.audit.get("assumptions") == (rules.get("assumptions") or [])
    assert baseline.audit["assumptions"], "统计研究有 3 条推定项，不该为空"


def test_assumed_typography_is_labelled_in_report(tmp_path):
    """标了 measured:false 的字号，条目里必须写「推定值」。"""
    from test_manuscript_matrix import build_messy_manuscript

    src = build_messy_manuscript(tmp_path / "m.docx")

    tjyj = render_into(tmp_path / "tjyj", src, target="tjyj")
    notes = [it["note"] for it in tjyj.items if it["action"] == "副标题套版"]
    assert notes, "稿件里有副标题，却没有副标题套版条目"
    assert any("推定" in n for n in notes), f"统计研究的副标题是推定值，却没标注：{notes}"

    # 没写 subtitle 规则的刊走代码默认值，不该冒称推定
    cjyj = render_into(tmp_path / "cjyj", src, target="cjyj")
    defaults = [it["note"] for it in cjyj.items if it["action"] == "副标题套版"]
    assert all("推定" not in n for n in defaults), f"无该规则时不应标推定：{defaults}"


def test_report_page_lists_assumptions(baseline, tmp_path):
    """审计页要真的把这些推定项列出来，不能只存在 JSON 里。"""
    from conftest import run_tool

    out = tmp_path / "audit.html"
    run_tool("render_audit.py", baseline.docx.with_suffix(".audit.json"), "-o", out, "--job", "t1")
    page = out.read_text(encoding="utf-8")
    assert "推定项" in page, "审计页没有列出推定项"
    assert baseline.audit["assumptions"][0][:12] in page, "推定项内容没有渲染进页面"
