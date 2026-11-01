"""第一承诺：正文字词零改动。

这是产品最贵的一句话——「不改你的字」。它必须由机器守，不能靠人点页面看。
"""

from conftest import TEXT_ACTIONS, note_reorder_blocks, signature


def test_every_changed_block_has_an_audit_record(baseline):
    """文字变了的块，审计里必须有对应条目。没有就是「偷偷改了」。

    纯换位的题注/表注除外——它们由 `表注／图注层序调整` 统一记录，那条记在
    表格上而不是记在段落上，因此单独要求「存在这样一条记录」。
    """
    reordered = note_reorder_blocks(baseline)
    recorded = {it["idx"] for it in baseline.items}

    missing = sorted(baseline.changed_idx - recorded - reordered)
    assert not missing, f"这些块文字变了却没有审计记录：{missing}"

    if reordered:
        actions = {it["action"] for it in baseline.items}
        assert "表注／图注层序调整" in actions, (
            f"块 {sorted(reordered)} 的题注被换了位置，但审计里没有层序调整记录")


def test_audit_before_and_after_match_the_files(baseline):
    """改文字类条目的 before/after 必须与源稿、成品逐字一致，报告不能说谎。

    其余条目（页面设置、套版、悬挂缩进）的 before/after 存的是格式值，
    不属于文字快照，不在此列。
    """
    text_items = [i for i in baseline.items if i["action"] in TEXT_ACTIONS]
    assert text_items, "没有采集到改文字的条目，夹具选刊可能过于接近"
    for it in text_items:
        idx = it["idx"]
        assert it["before"] == baseline.source_texts.get(idx), f"块 {idx} 的 before 与源稿不符"
        assert it["after"] == baseline.output_texts.get(idx), f"块 {idx} 的 after 与成品不符"


def test_text_changes_stay_within_the_whitelist(baseline):
    """抹掉白名单允许的归一后，改前改后必须完全相同。

    题注的纯换位（层序调整）不参与逐块比对——那两段文字本身没变，只是换了位置。
    """
    reordered = note_reorder_blocks(baseline)
    offenders = []
    for idx in sorted(baseline.changed_idx - reordered):
        btype = baseline.block_type.get(idx, "")
        before = signature(baseline.source_texts[idx], btype)
        after = signature(baseline.output_texts[idx], btype)
        if before != after:
            offenders.append((idx, btype,
                              baseline.source_texts[idx][:40],
                              baseline.output_texts[idx][:40]))
    assert not offenders, "白名单之外出现了文字改动：\n" + "\n".join(map(str, offenders))


def test_no_text_is_created_or_destroyed(baseline):
    """全局守恒：源稿里的每个字词，在成品里都必须还在（换位不算丢）。

    比的是归一后的指纹，这样标签归一、分隔符归并不会被误判成「凭空出现」，
    而真正被吞掉或塞进去的内容一定对不上。
    """
    from collections import Counter

    def fingerprints(texts, block_type):
        return Counter(signature(t, block_type.get(i, "")) for i, t in texts.items())

    before = fingerprints(baseline.source_texts, baseline.block_type)
    after = fingerprints(baseline.output_texts, baseline.block_type)

    created = ["".join(s)[:30] for s, n in (after - before).items() for _ in range(n)]
    lost = ["".join(s)[:30] for s, n in (before - after).items() for _ in range(n)]
    assert not created, f"成品里凭空出现了文字：{created[:3]}"
    assert not lost, f"成品里丢了文字：{lost[:3]}"


def test_block_structure_is_preserved(baseline):
    """不增删段落：源稿与成品的 body 子元素一一对应。"""
    assert set(baseline.output_texts) == set(baseline.source_texts)


def test_layout_items_do_not_touch_text(baseline):
    """A 档（纯版式）条目的 before/after 必须都是空——它们本就不该碰文字。"""
    layout = [it for it in baseline.items if not (it["before"] or it["after"])]
    assert layout, "没有采集到任何纯版式条目，夹具可能没跑到渲染"
    for it in layout:
        assert it["before"] == "" and it["after"] == ""
