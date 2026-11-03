"""F54 回归钉：审计「文字快照」登记的双向对账（conftest.text_snapshot_violations）。

原评审 N5-6：test_layout_items_do_not_touch_text 旧版用 before/after 为空筛出
「版式条目」、再断言它们的 before/after 为空——同义反复，验证力≈0。真正的静默
缺口是渲染器新增改文字动作而未加入 TEXT_ACTIONS 时，逐字校验（按动作名筛条目）
整类漏掉。

主修在 tests/test_text_fingerprint.py 的 test_layout_items_do_not_touch_text
（基线上跑双向对账）；本文件是证据里的复现用例——临时构造「动作名新但带快照
且未登记」的审计条目，旧实现视而不见，对账必须红。用例直接喂假流水线
（conftest.text_snapshot_violations 只读 items 与两侧文字表），不必真跑渲染。
"""


class _FakePipeline:
    """text_snapshot_violations 只需要 items 与两侧文字表，不必真跑流水线。"""

    def __init__(self, items, source, output):
        self.items = items
        self.source_texts = source
        self.output_texts = output


def test_unregistered_text_action_carrying_snapshot_is_red():
    """回归（F54）：新文字动作带了逐字快照却未登记进 TEXT_ACTIONS，必须红。

    临时构造一条「动作名是新的、before/after 与源稿/成品逐字一致、又不在
    TEXT_ACTIONS」的审计条目——旧版测试对它视而不见，双向对账必须抓出来。
    """
    from conftest import text_snapshot_violations

    p = _FakePipeline(
        [{"id": "7|神秘新动作", "idx": 7, "action": "神秘新动作",
          "before": "甲、引　言", "after": "一、引言"}],
        {7: "甲、引　言"}, {7: "一、引言"})
    bad = text_snapshot_violations(p)
    assert bad and bad[0][0] == "7|神秘新动作", f"未登记的文字快照没被抓到：{bad}"


def test_registered_text_action_without_snapshot_is_red():
    """登记为文字动作的条目缺了快照，同样算违规（双向的另一侧）。"""
    from conftest import text_snapshot_violations

    p = _FakePipeline(
        [{"id": "3|标签归一", "idx": 3, "action": "标签归一",
          "before": "", "after": ""}],
        {3: "内容摘要：甲"}, {3: "内容提要：甲"})
    bad = text_snapshot_violations(p)
    assert bad and bad[0][0] == "3|标签归一", f"缺快照的登记动作没被抓到：{bad}"


def test_format_values_and_layer_names_are_not_snapshots():
    """悬挂缩进的格式值与层序名词不是文字快照，未登记也不红（现状护栏）。

    「参考文献悬挂缩进」的 after 是「2.3 字符」、「表注／图注层序调整」的
    before/after 是层序名词，都不是快照——双向对账若把它们误判成未登记的
    文字动作，基线渲染就会假红。
    """
    from conftest import text_snapshot_violations

    p = _FakePipeline(
        [{"id": "-1|参考文献悬挂缩进", "idx": -1, "action": "参考文献悬挂缩进",
          "before": "", "after": "2.3 字符"},
         {"id": "9|表注／图注层序调整", "idx": 9, "action": "表注／图注层序调整",
          "before": "source / general", "after": "general / source"}],
        {-1: "", 9: "注：括号内为标准差。"}, {-1: "", 9: "资料来源：作者整理。"})
    assert text_snapshot_violations(p) == []
