"""第三承诺：可重复。同一输入同一结果；驳回一条只回退那一条。

用户在界面上逐条驳回改动，如果他驳回了第 3 条，第 7 条跟着变了，
这个工具就没法用了——他无法确认自己看到的就是最终交付的样子。
"""

import json

from conftest import RULES, TARGET_JOURNAL, layout_signature, render_into

# 驳回这些动作不会引发连锁（标题重新编号会改变后续编号，不适合做单条回退的样本）
SAFE_ACTIONS = {"标签归一", "关键词分隔符归一", "题注归一", "文献表标题归一"}


def test_rendering_twice_gives_identical_result(baseline, sample_docx, tmp_path):
    """同样的输入渲染两次，成品与审计必须逐字相同。"""
    again = render_into(tmp_path / "again", sample_docx)

    assert again.output_texts == baseline.output_texts, "两次渲染的正文不一致"

    def shape(p):
        return [(i["id"], i["action"], i["note"], i["before"], i["after"]) for i in p.items]

    assert shape(again) == shape(baseline), "两次渲染的审计条目不一致"


def test_layout_layer_is_reproducible(baseline, sample_docx, tmp_path):
    """版式层同样可重复（F40）：document.xml 与 styles.xml 的结构指纹一致。

    只比正文与审计条目形状时，「仅版式不同」的两次渲染照样全绿——第三承诺
    必须落到版式层。docx 是 zip、条目头带时间戳，裸比字节必挂，所以比
    layout_signature：解析后按文档顺序重新序列化、剔除 rsid 修订会话属性。
    """
    again = render_into(tmp_path / "layout-again", sample_docx)
    assert layout_signature(again.docx) == layout_signature(baseline.docx), \
        "同一输入两次渲染的版式结构（document.xml/styles.xml）不一致"


def test_ruleset_change_moves_the_layout(baseline, sample_docx, tmp_path):
    """「规则集是唯一真源」的反面：改了规则集，版式指纹必须漂移。

    若改了字号版式指纹还不动，说明版式值没来自规则集（硬编码或缓存）。
    """
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    rules["typography"]["body"]["size_pt"] = rules["typography"]["body"]["size_pt"] + 1.5
    alt_rules = tmp_path / "alt.json"
    alt_rules.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")

    alt = render_into(tmp_path / "layout-alt", sample_docx, rules_path=alt_rules)
    assert layout_signature(alt.docx) != layout_signature(baseline.docx), \
        "规则集改了正文号，版式指纹却没动——版式值没有来自规则集"


def test_rejecting_one_item_reverts_only_that_block(baseline, sample_docx, tmp_path):
    """驳回某条后，只有那一块回到原样，其余改动必须原封不动。"""
    target = next((i for i in baseline.text_editing_items if i["action"] in SAFE_ACTIONS), None)
    assert target is not None, "夹具里没有可安全驳回的条目，检查源稿与目标刊是否真的存在差异"

    idx = target["idx"]
    after = render_into(tmp_path / "rejected", sample_docx, rejects=[target["id"]])

    reverted = after.item_by_id(target["id"])
    assert reverted is not None and reverted["rejected"] is True, "驳回没被记录进审计"

    assert after.output_texts[idx] == baseline.source_texts[idx], "被驳回的块没有回到原样"

    drifted = {i for i in after.output_texts
               if after.output_texts[i] != baseline.output_texts[i]}
    assert drifted == {idx}, f"驳回一条却连带改变了其他块：{sorted(drifted - {idx})}"


def test_rejected_ids_are_echoed_in_audit(baseline, sample_docx, tmp_path):
    """驳回清单要落到报告里，事后才能解释「为什么这处没改」。"""
    target = next((i for i in baseline.text_editing_items if i["action"] in SAFE_ACTIONS), None)
    after = render_into(tmp_path / "echo", sample_docx, rejects=[target["id"]])
    assert target["id"] in (after.audit.get("rejected") or [])
