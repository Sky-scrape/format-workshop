"""第一承诺：正文字词零改动。

这是产品最贵的一句话——「不改你的字」。它必须由机器守，不能靠人点页面看。
"""

from conftest import TEXT_ACTIONS, signature


def test_every_changed_block_has_an_audit_record(baseline):
    """文字变了的块，审计里必须有对应条目。没有就是「偷偷改了」。"""
    recorded = {it["idx"] for it in baseline.items}
    missing = sorted(baseline.changed_idx - recorded)
    assert not missing, f"这些块文字变了却没有审计记录：{missing}"


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
    """抹掉白名单允许的归一后，改前改后必须完全相同。"""
    offenders = []
    for idx in sorted(baseline.changed_idx):
        btype = baseline.block_type.get(idx, "")
        before = signature(baseline.source_texts[idx], btype)
        after = signature(baseline.output_texts[idx], btype)
        if before != after:
            offenders.append((idx, btype,
                              baseline.source_texts[idx][:40],
                              baseline.output_texts[idx][:40]))
    assert not offenders, "白名单之外出现了文字改动：\n" + "\n".join(map(str, offenders))


def test_block_structure_is_preserved(baseline):
    """不增删段落：源稿与成品的 body 子元素一一对应。"""
    assert set(baseline.output_texts) == set(baseline.source_texts)


def test_layout_items_do_not_touch_text(baseline):
    """A 档（纯版式）条目的 before/after 必须都是空——它们本就不该碰文字。"""
    layout = [it for it in baseline.items if not (it["before"] or it["after"])]
    assert layout, "没有采集到任何纯版式条目，夹具可能没跑到渲染"
    for it in layout:
        assert it["before"] == "" and it["after"] == ""
