"""F53 回归钉：render_audit 的 KIND／REJECTABLE 与渲染器实际动作一一对应。

背景（审读 F53）：渲染器的动作表硬编码在 render_audit，两处脱节——
  1. KIND.get(action, "A") 兜底让渲染器实际产生的引用动作
     （正文引用改作者-年份／条目按目标刊体例重排／文献表排序／
     引用体系转换（已跳过）／编号样文本未转换）落成 A 档；
  2. 引用转换触及字符的两个动作按档位应是 B，说明型条目（只写 note、
     before/after 为空）混进「版式操作」还被计入「共 N 处调整」。

本文件守三条线：
  - 渲染器源码里每个 audit.add / restyle_audited 写死的动作名都在 KIND
    （不许兜底），KIND 里也没有渲染器已不产生的死键；
  - REJECTABLE 精确等于「有 audit.blocked 分支的 B 档动作」：引用转换的
    B 档动作渲染器只有 --citation 整批开关，没有逐条 blocked 分支，
    不能给勾选框（无效承诺）；
  - 说明型条目（N 档）不计入改动数，但必须单独立节显示。
"""

import ast
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from conftest import TOOLS  # noqa: E402

from render_audit import KIND, REJECTABLE, GRADE_DESC, render  # noqa: E402

RENDERER = TOOLS / "render_from_ir.py"


def _renderer_actions():
    """解析渲染器源码，收集写死的动作名。

    动作名出现在三处调用里：
      - audit.add(idx, btype, action, ...)          → 第 3 个位置参数
      - audit.blocked(idx, action)                  → 第 2 个位置参数
      - restyle_audited(p, audit, idx, btype, action, ...) → 第 5 个
    三元表达式两侧都收（「摘要套版」/「关键词套版」）。只认接收者名为
    audit 的 add/blocked，避开 fmt_snapshot 里集合的 .add 之类同名方法。
    """
    tree = ast.parse(RENDERER.read_text(encoding="utf-8"))
    added, blocked = set(), set()

    def strs(node):
        # IfExp 只收两侧取值（「摘要套版」/「关键词套版」），
        # 不收条件表达式里的普通字符串（t == "abstract"）。
        if isinstance(node, ast.IfExp):
            return strs(node.body) | strs(node.orelse)
        return {n.value for n in ast.walk(node)
                if isinstance(n, ast.Constant) and isinstance(n.value, str)}

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        f = node.func
        recv = getattr(f, "value", None)
        if isinstance(f, ast.Attribute) and isinstance(recv, ast.Name) \
                and recv.id == "audit":
            if f.attr == "add" and len(node.args) >= 3:
                added |= strs(node.args[2])
            elif f.attr == "blocked" and len(node.args) >= 2:
                blocked |= strs(node.args[1])
        elif isinstance(f, ast.Name) and f.id == "restyle_audited" \
                and len(node.args) >= 5:
            added |= strs(node.args[4])
    return added, blocked


# ---------- 一一对应：渲染器动作名 vs KIND ----------

def test_every_renderer_action_has_explicit_kind():
    """KIND.get 的「A」兜底不许吃掉任何真实动作——漏登记即档位造假。"""
    added, _ = _renderer_actions()
    unclassified = sorted(added - set(KIND))
    assert not unclassified, f"渲染器动作没进 KIND，会兜底成 A 档：{unclassified}"
    stale = sorted(set(KIND) - added)
    assert not stale, f"KIND 里的死键（渲染器已不再产生）：{stale}"


def test_citation_actions_graded_by_whether_they_touch_text():
    """引用动作按「是否触及字符」定档：改字的 B，纯重排的 A，只说明的 N。"""
    assert KIND["正文引用改作者-年份"] == "B"
    assert KIND["条目按目标刊体例重排"] == "B"
    assert KIND["文献表排序"] == "A"
    assert KIND["引用体系转换（已跳过）"] == "N"
    assert KIND["编号样文本未转换"] == "N"
    assert "N" in GRADE_DESC


# ---------- REJECTABLE 的边界 ----------

def test_rejectable_exactly_matches_blocked_b_grade_actions():
    """REJECTABLE == 有 audit.blocked 分支的 B 档动作。

    少了：声明「逐条可驳回」的动作在页面上没有勾选框；
    多了：渲染器没有 blocked 分支的动作给勾选框，勾了也不生效。
    """
    added, blocked = _renderer_actions()
    assert added, "渲染器源码里一个动作名都没解析到，收集器本身失效"
    assert blocked, "渲染器源码里没有任何 blocked 分支，收集器本身失效"
    b_grades = {a for a, g in KIND.items() if g == "B"}
    assert REJECTABLE == b_grades & blocked, (
        f"REJECTABLE 与「blocked 分支 ∩ B 档」不一致：多了 "
        f"{sorted(REJECTABLE - (b_grades & blocked))}，少了 "
        f"{sorted((b_grades & blocked) - REJECTABLE)}")


def test_citation_b_actions_stay_unrejectable():
    """引用转换只有 --citation 整批开关，逐条驳回未实现，不给勾选框。"""
    assert "正文引用改作者-年份" not in REJECTABLE
    assert "条目按目标刊体例重排" not in REJECTABLE


# ---------- F53 证据复现：审计页上的表现 ----------

def _item(action, idx=6, before="", after="", note="", rule="r"):
    return {"id": f"{idx}|{action}", "idx": idx, "sec": "", "block": "body",
            "action": action, "before": before, "after": after,
            "note": note, "rule": rule, "rejected": False}


def _kpi(html_out, label):
    m = re.search(rf"<b>(\d+)</b><span>{label}</span>", html_out)
    assert m, f"KPI 里没有「{label}」"
    return int(m.group(1))


def _section(html_out, start, stop):
    return html_out.split(start, 1)[1].split(stop, 1)[0]


def test_b_grade_actions_render_reject_checkbox():
    """带关键词分隔符／文献条目编号改动的审计：该条目有驳回勾选框。"""
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        _item("关键词分隔符归一", idx=7, before="GD P CPI", after="GDP；CPI"),
        _item("文献条目编号归一", idx=8, before="[1] 张某. 研究.", after="〔1〕张某. 研究."),
    ]}
    html_out = render(audit, job="t1")
    for it in audit["items"]:
        assert f'value="{it["id"]}"' in html_out, f"{it['action']} 没有驳回勾选框"
    assert "保留原文" in html_out


def test_citation_b_grade_items_count_in_kpi():
    """引用转换审计的 KPI 计入 B 档：改字条目进「触及字符」表，不是折叠组。"""
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        _item("正文引用改作者-年份", idx=12,
              before="文献[1]显示规模效应显著。",
              after="文献（张某，2020）显示规模效应显著。",
              rule="citation.in_text"),
        _item("条目按目标刊体例重排", idx=20,
              before="[1] 张某. 规模效应研究. 统计研究, 2020.",
              after="张某. 规模效应研究[J]. 统计研究, 2020(3): 12-20.",
              rule="citation.reference_list.templates"),
    ]}
    html_out = render(audit, job="t1")

    assert [it["grade"] for it in audit["items"]] == ["B", "B"], "引用改字条目应定 B 档"
    assert _kpi(html_out, "字符级改动（B 档）") == 2, "B 档 KPI 漏计引用转换"
    assert _kpi(html_out, "版式操作（A 档）") == 0
    text_part = _section(html_out, 'id="textTable"', 'id="layoutSection"')
    for action in ("正文引用改作者-年份", "条目按目标刊体例重排"):
        assert action in text_part, f"{action} 没出现在「触及字符的改动」表里"
    layout_part = _section(html_out, 'id="layoutSection"', "</body>")
    for action in ("正文引用改作者-年份", "条目按目标刊体例重排"):
        assert action not in layout_part, f"{action} 被藏进版式折叠组"
    # 无逐条驳回分支：不给勾选框
    assert 'value="12|正文引用改作者-年份"' not in html_out


def test_note_items_kept_visible_but_not_counted():
    """说明条目（N 档）不计入「共 N 处调整」与 A/B KPI，但单独立节显示。"""
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        _item("引用体系转换（已跳过）", idx=-3,
              note="第 3 条文献在正文里没有对应引用；双向校验不通过，整批不转",
              rule="citation.reference_list.must_require_bidirectional_match"),
        _item("编号样文本未转换", idx=12,
              note="形似引用但紧邻数量/区间说法，无法安全判定，原样保留：…调研对象为[90, 100]家企业…",
              rule="citation.in_text"),
        _item("页面设置", idx=-1, note="上下 2.54cm 左右 3.17cm"),
    ]}
    html_out = render(audit)

    m = re.search(r"共 (\d+) 处调整", html_out)
    assert m, "页眉没有「共 N 处调整」"
    assert m.group(1) == "1", "说明条目被计进了改动总数"
    assert _kpi(html_out, "版式操作（A 档）") == 1, "说明条目不该算版式操作"
    assert _kpi(html_out, "字符级改动（B 档）") == 0

    assert 'id="noteSection"' in html_out, "说明条目没有显示区，藏起来等于当没发生"
    note_part = _section(html_out, 'id="noteSection"', "</body>")
    layout_part = _section(html_out, 'id="layoutSection"', 'id="noteSection"')
    for action in ("引用体系转换（已跳过）", "编号样文本未转换"):
        assert action in note_part, f"{action} 没出现在说明节"
        assert action not in layout_part, f"{action} 被算成版式操作折叠组"


def test_render_is_deterministic():
    """同一份审计渲染两次，输出逐字节相同（承诺③：同输入同输出）。"""
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        _item("关键词分隔符归一", idx=7, before="GDP CPI", after="GDP；CPI"),
        _item("正文引用改作者-年份", idx=12,
              before="文献[1]显示。", after="文献（张某，2020）显示。"),
        _item("编号样文本未转换", idx=13,
              note="形似引用但紧邻数量说法，原样保留：…[90, 100]家…"),
        _item("页面设置", idx=-1, note="上下 2.54cm"),
    ]}
    first = render(audit, job="t1")
    second = render(audit, job="t1")
    assert first == second


def test_batch_reject_checkbox_per_action_group():
    """同类 B 档条目按动作分组，每组一个「全选本类」开关（DESIGN §3）。

    勾选框仍是逐条 name=rid（后端接口不变）；分组开关只是前端把本组
    可驳回的勾选框一起点掉，驳回粒度没有变粗。
    """
    audit = {"source": "x.docx", "target": "统计研究", "items": [
        _item("段尾空白删除", idx=3, before="甲  ", after="甲"),
        _item("段尾空白删除", idx=4, before="乙  ", after="乙"),
        _item("关键词分隔符归一", idx=7, before="GDP CPI", after="GDP；CPI"),
    ]}
    html_out = render(audit, job="t1")
    assert html_out.count('class="pickg"') == 2, "每个动作组应各有一个批量开关"
    assert 'data-g="0"' in html_out and 'data-g="1"' in html_out
    # 逐条驳回勾选框仍在（两条段尾 + 一条关键词）
    assert html_out.count('name="rid"') == 3
    # 分组标题行确实是全表宽，不会把列挤错
    assert 'class="aghead"' in html_out
