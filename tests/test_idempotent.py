"""第三承诺：可重复。同一输入同一结果；驳回一条只回退那一条。

用户在界面上逐条驳回改动，如果他驳回了第 3 条，第 7 条跟着变了，
这个工具就没法用了——他无法确认自己看到的就是最终交付的样子。
"""

from conftest import render_into

# 驳回这些动作不会引发连锁（标题重新编号会改变后续编号，不适合做单条回退的样本）
SAFE_ACTIONS = {"标签归一", "关键词分隔符归一", "题注归一", "文献表标题归一"}


def test_rendering_twice_gives_identical_result(baseline, sample_docx, tmp_path):
    """同样的输入渲染两次，成品与审计必须逐字相同。"""
    again = render_into(tmp_path / "again", sample_docx)

    assert again.output_texts == baseline.output_texts, "两次渲染的正文不一致"

    def shape(p):
        return [(i["id"], i["action"], i["note"], i["before"], i["after"]) for i in p.items]

    assert shape(again) == shape(baseline), "两次渲染的审计条目不一致"


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
