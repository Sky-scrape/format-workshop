"""F31 / F39 / F41 的回归钉——守卫对齐规范而不是对齐实现。

F40（版式层可重复性）的两条用例钉在归属地 tests/test_idempotent.py，其余
证据复现集中在本文。

- F31  conftest 的 normalize/signature 是 DESIGN 白名单与渲染器正则的手工副本，
       且漂移得比白名单更宽松：无条件删全部空白——拉丁与中文间的空格删除
       不在任何白名单条目里，invariant_failures 对这类违规返回 []；标题编号
       剥离还作用于 body 块；它与 test_text_fingerprint 又是同一不变式的两份
       实现。→ 收紧到 DESIGN §4.1 实现口径（按块类型），四条不变式的判定
       抽到 conftest 共享函数，两处调用同一批函数。
- F39  --citation 渲染路径整体游离在四承诺测试之外：harness 从未跑过
       citation 模式。→ render_into 加 citation 开关；invariant_failures 对
       引用改写块豁免「白名单外改动」，但豁免凭证是审计条目本身（action+
       rule 对得上），改了字没记录照样违规；端到端用 mwm 规则集 + 带编号
       引用的最小稿验证。顺带钉 app 勾选闸（P4-1 阶段 D 一行）：与渲染层
       硬闸同一判据，编号实引刊不再 offered。
- F41  A 档条目的 rule 引用无一般性强制：可解析性测试静默放过空 rule。
       → conftest.change_items_without_rule 做一般判定（改动型 = 文字快照
       变了 / 动作以「套版」结尾 / 有格式目标值；说明型豁免），
       test_audit_traceable.py 落断言并对三本入库刊各跑一遍；本文用证据
       复现钉守卫本身。
"""

import json
from types import SimpleNamespace

import pytest

from conftest import (RULES, change_items_without_rule, citation_rewrite_blocks,
                      invariant_failures, layout_signature, normalize, render_into,
                      signature, whitelist_violations)


# ---------------------------------------------------------------- F31 守卫口径

def test_latin_cjk_space_is_no_longer_normalized_away():
    """证据复现：中文与拉丁之间的空格删除不在任何白名单条目。

    修前 normalize 无条件删全部空白，normalize('数字金融 A 股') ==
    normalize('数字金融A股')，渲染器真做出这类改动时 invariant_failures
    返回 []。
    """
    assert normalize("数字金融 A 股") != normalize("数字金融A股"), \
        "正文内空格仍被守卫抹平——N-02/N-06 未实现，不属于白名单"


def test_fullwidth_halfwidth_folding_is_gone():
    """全半角**标点**折叠（N-03/N-04）仍未实现，守卫不得替它开脱。

    N-05 只收全角字母数字（2026-11-13 落地）；全角逗号属于语境判断，
    不在白名单，逐字比。
    """
    assert normalize("研究发现，增长", "body") != normalize("研究发现,增长", "body")
    assert normalize("２０２０年", "body") == normalize("2020年", "body")


def test_heading_number_prefix_strip_is_scoped_to_headings():
    """证据复现：标题编号剥离只许作用于标题块，body 逐字比。"""
    assert normalize("二、研究假设", "heading1") == "研究假设"
    assert normalize("二、研究假设", "body") == "二、研究假设", \
        "body 块的「二、」前缀被守卫剥掉了——渲染器对 body 没有这个授权"


def test_body_trailing_whitespace_is_now_whitelisted():
    """段尾空白删除已实现（DESIGN §4.1 N-01 段尾，2026-11-13）。

    守卫必须按同一口径放行；正文内部空格仍不被抹平（见上一条），
    「无条件删全部空白」的宽松口径没有回来。
    """
    assert normalize("正文结尾 ", "body") == normalize("正文结尾", "body")
    assert normalize("正文 内部", "body") != normalize("正文内部", "body")
    # 非正文块不吃段尾删除：文献条目仍逐字比
    assert normalize("条目  ", "reference_item") != normalize("条目", "reference_item")


def test_keyword_terms_keep_inner_word_space():
    """关键词词条逐字比：「digital finance」的词间空格不许被抹掉。

    修前守卫按 `[；;、,，\\s\\u3000]+` 过切词条：渲染器错把一个关键词拆成
    两个（「digital finance」→「digital；finance」）时，修前修后的词条元组
    一样，这类词条数变化完全不可见。现在默认按非空白分隔符的保守口径切
    （与渲染器 split_keywords 同一），并把目标刊分隔符透传给 signature——
    分隔符本身是空白的刊（gjjryj zh 是两个半角空格）仍按渲染器口径切。
    """
    assert signature("关键词：digital finance；融资约束", "keywords") \
        == ("digital finance", "融资约束")
    assert signature("关键词：digitalfinance；融资约束", "keywords") \
        != ("digital finance", "融资约束")
    # 目标分隔符本身是空白（gjjryj zh="  "）：按渲染器的空白分支切
    assert signature("关键词：digital finance；融资约束", "keywords", kw_sep="  ") \
        == ("digital", "finance", "融资约束")


def test_invariant_failures_assemble_shared_predicates(baseline, monkeypatch):
    """invariant_failures 必须装配共享判定，不许留一份私有抄本。"""
    import conftest

    seen = []

    def fake_whitelist(p):
        seen.append("whitelist")
        return [("哨兵块", "body", "改前", "改后")]

    def fake_unaudited(p):
        seen.append("unaudited")
        return [7]

    monkeypatch.setattr(conftest, "whitelist_violations", fake_whitelist)
    monkeypatch.setattr(conftest, "unaudited_changes", fake_unaudited)

    problems = conftest.invariant_failures(baseline)
    assert any("哨兵" in s for s in problems), "invariant_failures 没走共享的白名单判定"
    assert sorted(seen) == ["unaudited", "whitelist"]


def test_fingerprint_tests_import_the_shared_predicates():
    """test_text_fingerprint 与 invariant_failures 必须同一实现（F31 抽到一处）。"""
    import conftest
    import test_text_fingerprint as tf

    for name in ("unaudited_changes", "whitelist_violations",
                 "fingerprint_delta", "block_count_mismatch"):
        assert getattr(tf, name) is getattr(conftest, name), \
            f"test_text_fingerprint 的 {name} 不是共享实现——同一不变式又抄了一份"


# ---------------------------------------------------------------- F39 勾选闸与端到端

def test_citation_gate_matches_renderer_hard_gate(tmp_path):
    """app 勾选闸与渲染层 apply_citation 硬闸同一判据（P4-1 阶段 D 一行）。

    cjyj 有条目模板（描述式）但声明编号实引：渲染层会整批拒绝，UI 若仍
    offered，勾选框永远落空。
    """
    import app

    d = tmp_path
    ir = {"blocks": [{"idx": 0, "type": "body", "text": "引用了文献[1]和[2]。"}]}
    (d / "manuscript.ir.json").write_text(json.dumps(ir, ensure_ascii=False),
                                          encoding="utf-8")

    assert app.details_context(d, {"journal": "tjyj"})["citation_offered"] is True, \
        "作者-年份刊（tjyj）的勾选闸不该被误关"
    assert app.details_context(d, {"journal": "cjyj"})["citation_offered"] is False, \
        "编号实引刊（cjyj）勾了也会被硬闸拒绝，UI 不该开勾选框"
    assert app.details_context(d, {"journal": "gjjryj"})["citation_offered"] is False, \
        "无模板刊的闸保持关闭"


MWM_RULES = RULES / "mwm.json"


def build_numbered_citation_manuscript(path):
    """带编号引用的最小稿：正文三处引用点 + 两条可双向对上的文献。"""
    from docx import Document

    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("正文第二段，用合并写法引用了[1-2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.add_paragraph("[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.")
    doc.save(str(path))
    return path


@pytest.mark.skipif(not MWM_RULES.is_file(),
                    reason="mwm.json 不入库（私有规则集），本机没有则跳过")
def test_citation_render_passes_invariant_harness(tmp_path):
    """端到端（F39）：mwm（作者-年份 + 占位符模板）+ 带编号引用的最小稿。

    转换后文字差异必须全部来自引用改写——白名单口径解释不了的差异，逐块
    都有 action+rule 对得上的审计条目；invariant_failures（含引用豁免）全绿；
    同输入二次渲染幂等（文字、审计形状、版式指纹三层）。
    """
    src = build_numbered_citation_manuscript(tmp_path / "cite-src.docx")
    p = render_into(tmp_path / "cite", src, target="mwm", rules_path=MWM_RULES,
                    citation=True)

    # 审计完整性：每个正文改写点都有 rule=citation.in_text 的条目
    in_text = [it for it in p.items if it["action"] == "正文引用改作者-年份"]
    assert in_text, "编号引用没有被改写——夹具或 mwm 规则集不满足端到端条件"
    assert all(it["rule"] == "citation.in_text" for it in in_text), \
        "正文引用改写点的审计条目 rule 指向不对"
    entries = [it for it in p.items if it["action"] == "条目按目标刊体例重排"]
    assert entries and all(it["rule"] == "citation.reference_list.templates"
                           for it in entries), "条目重排缺审计或 rule 不对"

    # 文字差异全部来自引用改写
    body = "\n".join(p.output_texts.values())
    assert "（张三，2020）" in body and "（Smith，2021）" in body
    unexplained = [idx for idx in sorted(p.changed_idx)
                   if signature(p.source_texts[idx], p.block_type.get(idx, ""))
                   != signature(p.output_texts[idx], p.block_type.get(idx, ""))
                   and idx not in citation_rewrite_blocks(p)]
    assert not unexplained, f"这些块的差异没有引用改写条目背书：{unexplained}"

    # 四承诺 harness（含引用豁免）全绿
    assert invariant_failures(p) == [], str(invariant_failures(p))

    # 二次幂等：文字、审计形状、版式指纹
    q = render_into(tmp_path / "cite-again", src, target="mwm", rules_path=MWM_RULES,
                    citation=True)
    assert q.output_texts == p.output_texts, "citation 渲染两次的正文不一致"

    def shape(r):
        return [(i["id"], i["action"], i["note"], i["before"], i["after"]) for i in r.items]

    assert shape(q) == shape(p), "citation 渲染两次的审计条目不一致"
    assert layout_signature(q.docx) == layout_signature(p.docx), \
        "citation 渲染两次的版式结构不一致"


@pytest.mark.skipif(not MWM_RULES.is_file(),
                    reason="mwm.json 不入库（私有规则集），本机没有则跳过")
def test_citation_exemption_needs_the_audit_item(tmp_path, monkeypatch):
    """豁免不豁免审计：把改写条目的 rule 改错，对应块必须重新算违规。"""
    src = build_numbered_citation_manuscript(tmp_path / "cite-src.docx")
    p = render_into(tmp_path / "cite", src, target="mwm", rules_path=MWM_RULES,
                    citation=True)
    assert invariant_failures(p) == []

    import copy

    import conftest

    broken = copy.deepcopy(p.items)
    for it in broken:
        if it["action"] == "正文引用改作者-年份":
            it["rule"] = "citation.reference_list.templates"   # 指错位置
    tampered = copy.copy(p)
    tampered.items = broken

    offenders = {idx for idx, *_r in whitelist_violations(tampered)}
    assert offenders, "rule 指错的引用改写仍被豁免——豁免凭证没有校验条目"


# ---------------------------------------------------------------- F41 守卫本身

def test_change_item_without_rule_is_flagged():
    """证据复现：改动型条目漏掉 rule= 必须被抓到（修前静默放过）。"""
    rules = json.loads((RULES / "cjyj.json").read_text(encoding="utf-8"))

    items = [
        # 文字快照变了（B 档）却没带 rule——修前 text_editing_items 之外无人管，
        # 且可解析性循环 if it["rule"] and ... 静默放过空值
        {"id": "41|标签归一", "action": "标签归一", "before": "内容摘要：x", "after": "摘 要：x"},
        # A 档套版（before/after 为空、说明在 note）没带 rule——修前 B 档测试不覆盖
        {"id": "9|标题套版", "action": "标题套版", "note": "黑体 15pt，居中"},
        # 记录格式目标值的条目没带 rule
        {"id": "41|参考文献悬挂缩进", "action": "参考文献悬挂缩进", "after": "2.3 字符"},
        # rule 指不到实处
        {"id": "3|正文套版", "action": "正文套版", "note": "宋体 10.5pt", "rule": "typography.nope"},
    ]
    bad = change_items_without_rule(SimpleNamespace(items=items), rules)
    assert len(bad) == 4, f"改动型条目缺依据没被全部抓到：{bad}"


def test_informational_items_are_exempt():
    """说明型条目（before/after 全空、非套版）豁免：它们只解释为什么没动。"""
    rules = json.loads((RULES / "cjyj.json").read_text(encoding="utf-8"))
    items = [
        {"id": "-3|citation", "action": "引用体系转换（已跳过）",
         "note": "该刊声明编号实引，不转", "rule": "citation.system"},
        {"id": "12|citation", "action": "编号样文本未转换",
         "note": "形似引用但紧邻数量说法，原样保留", "rule": "citation.in_text"},
        {"id": "30|equation", "action": "公式（伪表格）原样保留",
         "note": "1 行表格排版的公式与编号不做改动"},
        # 驳回后 before==after（保留原文）：不构成改动，但带的 rule 仍须可解析
        {"id": "41|段首缩进空格删除", "action": "段首缩进空格删除",
         "before": "　　正文", "after": "　　正文", "note": "已驳回", "rule": "typography.body"},
        {"id": "41|段首缩进空格删除b", "action": "段首缩进空格删除",
         "before": "　　正文", "after": "　　正文", "note": "已驳回", "rule": "typography.nope"},
    ]
    bad = change_items_without_rule(SimpleNamespace(items=items), rules)
    assert bad == [("41|段首缩进空格删除b", "段首缩进空格删除", "typography.nope")], \
        f"说明型条目未按约定豁免，或驳回条目的坏引用没被抓到：{bad}"
