"""F18 回归钉：总字数上限必须按规则集声明的 scope 过滤后再判。

修前症状（审计 F18）：9 本声明 scope=正文 的刊按全文（含标题/摘要/参考文献）
计数去对正文上限——tjyj（strict=true）出假红档，其余 8 本出假黄档；报告
「统计范围」标注为正文而数字实为全文口径，计数与标注自相矛盾。成因：唯一
调用点 compliance_report.check 不传 scope，count_words 的 body 分支是死代码，
scope 字段全文件从不参与过滤。

修法口径（compliance_report.check）：把 limits.total.scope 接进统计——
「正文」只数 body 类块，「全文」「全文（含图表）」数全部块；写法未识别或
未声明保守按全文计并在报告注明回退。判决句、待办、KPI、「统计范围」格的
口径标注一律跟随实际统计的块。

用例即审计复现：内存 IR（body=17000 + abstract=1000 + reference_item=1000
个中文字符，B 口径全文合计 19000）对 tjyj（上限 18000、scope=正文、strict）
——修前判「正文超上限」红档加「正文缩减」待办，修后应判「正文字数在范围内」。
"""

import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from conftest import RULES  # noqa: E402

import compliance_report as crp  # noqa: E402


def scope_ir():
    """审计复现用的内存 IR：三块纯中文文本，B 口径恰好 17000+1000+1000。"""
    return {
        "source": "repro.docx",
        "blocks": [
            {"type": "abstract", "text": "摘要" * 500},        # 1000 个中文字符
            {"type": "body", "text": "数字金融" * 4250},       # 17000 个中文字符
            {"type": "reference_item", "text": "文献" * 500},  # 1000 个中文字符
        ],
    }


def load_rules(name):
    return json.loads((RULES / f"{name}.json").read_text(encoding="utf-8"))


def scope_row(html_out):
    """取出「统计范围」一格的可见内容。"""
    m = re.search(r"统计范围</td><td>(.*?)</td>", html_out, re.S)
    assert m, "报告里没有「统计范围」这一行"
    return m.group(1)


# ---------- scope=正文：只数正文块 ----------

def test_body_scope_counts_body_only():
    """tjyj（scope=正文、strict）：全文 19000 不再顶到正文上限上。"""
    rules = load_rules("tjyj")
    res = crp.check(scope_ir(), rules)
    tl = res["total_limit"]
    assert tl["actual"] == 17000, f"应只统计正文块（17000），实际 {tl['actual']}"
    assert tl["over"] == tl["actual"] - tl["value"] < 0

    html_out = crp.render(scope_ir(), rules, res)
    assert "正文字数在范围内" in html_out, "判决句应按正文口径判合格"
    assert "正文超" not in html_out, "全文数字冒充正文口径的假超限仍在"
    assert "正文缩减" not in html_out, "待办不该派发假缩减任务"
    assert 'class="pill bad"' not in html_out, "strict 超限的红档应消失"
    # 三种口径表同样只统计正文块；「统计范围」格与实际计算一致
    assert res["words"][1][1].startswith("17,000")
    assert "正文口径字数" in html_out
    assert scope_row(html_out) == "正文"


# ---------- scope=全文：摘要与参考文献必须计入 ----------

@pytest.mark.skipif(not (RULES / "jjgl.json").is_file(),
                    reason="jjgl 属完整规则库（不开源），公开检出无此刊")
def test_full_scope_counts_whole_manuscript():
    """jjgl（scope=全文）：19000 全文对上限出黄档，判决口径必须如实说「全文」。"""
    rules = load_rules("jjgl")
    res = crp.check(scope_ir(), rules)
    tl = res["total_limit"]
    assert tl["actual"] == 19000, f"scope=全文 应统计全部块，实际 {tl['actual']}"
    assert tl["over"] == tl["actual"] - tl["value"] > 0

    html_out = crp.render(scope_ir(), rules, res)
    assert f"全文超建议值 {tl['over']:,} 字" in html_out
    assert "正文超" not in html_out, "全文口径的判决不得谎称正文"
    assert f"全文缩减 {tl['over']:,} 字" in html_out
    assert res["words"][1][1].startswith("19,000")
    assert scope_row(html_out) == "全文"


# ---------- scope 未识别 / 未声明：保守按全文并在报告注明 ----------

def test_unknown_scope_falls_back_to_full_text_with_note():
    """scope 写法没认出时保守按全文计，报告注明回退并照抄原声明。"""
    rules = load_rules("tjyj")
    rules["limits"]["total"]["scope"] = "全文（不含附录）"  # 仓库 23 本里没有的写法
    res = crp.check(scope_ir(), rules)
    tl = res["total_limit"]
    assert tl["actual"] == 19000, "未识别的 scope 保守按全文计"
    assert tl["scope_note"], "回退必须带说明"

    html_out = crp.render(scope_ir(), rules, res)
    assert "全文（不含附录）" in html_out, "规则集原声明应原样呈现"
    assert "写法未识别，按全文口径统计" in html_out, "回退未注明"
    assert f"全文超上限 {tl['over']:,} 字" in html_out, "判决口径应与实际统计一致"


def test_missing_scope_falls_back_and_says_so():
    """未声明 scope 时不再默认标「正文」：按全文计并注明，标签不再撒谎。"""
    rules = load_rules("tjyj")
    del rules["limits"]["total"]["scope"]
    res = crp.check(scope_ir(), rules)
    tl = res["total_limit"]
    assert tl["actual"] == 19000
    assert tl["scope"] == "" and tl["scope_note"], "缺失 scope 应记录回退说明"

    html_out = crp.render(scope_ir(), rules, res)
    assert "规则集未声明统计范围，按全文口径统计" in scope_row(html_out)


# ---------- 承诺③：同一输入两次渲染结果相同 ----------

def test_render_is_deterministic():
    rules = load_rules("tjyj")
    res = crp.check(scope_ir(), rules)
    assert crp.render(scope_ir(), rules, res) == crp.render(scope_ir(), rules, res)
