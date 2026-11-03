"""第一承诺：正文字词零改动。

这是产品最贵的一句话——「不改你的字」。它必须由机器守，不能靠人点页面看。
"""

from conftest import (TEXT_ACTIONS, block_count_mismatch, fingerprint_delta,
                      note_reorder_blocks, text_snapshot_violations,
                      unaudited_changes, whitelist_violations)


def test_every_changed_block_has_an_audit_record(baseline):
    """文字变了的块，审计里必须有对应条目。没有就是「偷偷改了」。

    纯换位的题注/表注除外——它们由 `表注／图注层序调整` 统一记录，那条记在
    表格上而不是记在段落上，因此单独要求「存在这样一条记录」。

    判定与矩阵测试的 invariant_failures 同源（conftest.unaudited_changes）。
    """
    reordered = note_reorder_blocks(baseline)
    missing = unaudited_changes(baseline)
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

    题注的纯换位（层序调整）不参与逐块比对——那两段文字本身没变，只是换了
    位置。判定与矩阵测试同源（conftest.whitelist_violations，含引用改写豁免）。
    """
    offenders = whitelist_violations(baseline)
    assert not offenders, "白名单之外出现了文字改动：\n" + "\n".join(map(str, offenders))


def test_no_text_is_created_or_destroyed(baseline):
    """全局守恒：源稿里的每个字词，在成品里都必须还在（换位不算丢）。

    比的是归一后的指纹，这样标签归一、分隔符归并不会被误判成「凭空出现」，
    而真正被吞掉或塞进去的内容一定对不上。判定与矩阵测试同源
    （conftest.fingerprint_delta）。
    """
    created_c, lost_c = fingerprint_delta(baseline, exclude=note_reorder_blocks(baseline))
    created = ["".join(s)[:30] for s, n in created_c.items() for _ in range(n)]
    lost = ["".join(s)[:30] for s, n in lost_c.items() for _ in range(n)]
    assert not created, f"成品里凭空出现了文字：{created[:3]}"
    assert not lost, f"成品里丢了文字：{lost[:3]}"


def test_block_structure_is_preserved(baseline):
    """不增删段落：源稿与成品的 body 子元素一一对应（判定与矩阵测试同源）。"""
    assert not block_count_mismatch(baseline), \
        "段落数量变了：源稿与成品的 body 子元素不再一一对应"


def test_layout_items_do_not_touch_text(baseline):
    """A 档（纯版式）条目不得携带文字快照——双向对账（F54）。

    旧版用 before/after 为空筛出「版式条目」、再断言它们的 before/after 为空——
    同义反复，验证力≈0。真正的静默缺口在反方向：渲染器新增了改文字动作却没
    登记进 TEXT_ACTIONS 时，上面的逐字校验会整类漏掉它。所以这里按
    conftest.text_snapshot_violations 双向对账：
    ① 登记为文字动作的条目必须带非空 before/after；
    ② 未登记的条目不得携带与源稿/成品逐字一致的快照对（未登记即红，逼它
       先进 TEXT_ACTIONS 再谈逐字）。
    版式条目记的格式目标值（悬挂缩进的「2.3 字符」）与层序名词（表注层序
    调整的「source / general」）不是文字快照，不在 ② 之列。
    """
    layout = [it for it in baseline.items if not (it["before"] or it["after"])]
    assert layout, "没有采集到任何纯版式条目，夹具可能没跑到渲染"
    bad = text_snapshot_violations(baseline)
    assert not bad, "审计快照与动作登记对不上：\n" + "\n".join(map(str, bad))


# F54 的三条复现用例（未登记快照红／登记缺快照红／格式值不算快照）在
# tests/test_fix_f54_snapshot_reconciliation.py——回归测试放新建独立文件。
