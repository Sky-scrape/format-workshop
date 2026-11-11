"""总字数下限与标题上限：规则集有两种写法，报告和规格页都得认。

- 下限：cjyj 写在 limits.total.min，cmjj 写在 limits.total_min。只读上限时，
  「不接受短文」的刊在报告里等于没有下限，也不会算 under。
- 标题上限：多数刊写 title_chars_max，cjwt 写 title_max_chars。只认一个键，
  那本刊的标题字数限制就在报告/规格页上消失。
- 统计范围：cjwt 的「正文（含字符）」就是正文口径，不是「正文 + 字符数」；
  认不出来会退回全文，把参考文献算进正文上限。

用例从入库的 cjyj / tjyj 合成（mutated_rules），不依赖私有刊文件。
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

from conftest import mutated_rules  # noqa: E402

import compliance_report as crp  # noqa: E402
import render_spec  # noqa: E402


def _rules(tmp_path, base, name, **keys):
    path = mutated_rules(tmp_path, base, name, **keys)
    return json.loads(path.read_text(encoding="utf-8"))


def _ir(*blocks):
    return {"source": "字数稿.docx", "blocks": list(blocks)}


def _body(n):
    return {"type": "body", "text": "正" * n}


# ---------------------------------------------------------------- 下限（两种键名）

def test_total_min_inside_total_is_read(tmp_path):
    """cjyj 的 limits.total.min：报告读到下限 15000，算出 under。"""
    rules = _rules(tmp_path, "cjyj", "cjyj.json")
    res = crp.check(_ir(_body(8000)), rules)
    tl = res["total_limit"]

    assert tl["lo"] == 15000
    assert tl["actual"] == 8000
    assert tl["under"] == 7000, "低于下限的部分没算出来"

    page = crp.render(_ir(_body(8000)), rules, res)
    assert "下限 15,000" in page, "三种字数口径里没标下限"
    assert "低于下限 7,000" in page, "实测低于下限没提示"


def test_total_min_top_level_is_read(tmp_path):
    """cmjj 风格 limits.total_min：同样要读到并算 under。"""
    rules = _rules(tmp_path, "tjyj", "tjyj_total_min.json",
                   **{"limits.total_min": 15000})
    res = crp.check(_ir(_body(8000)), rules)
    tl = res["total_limit"]

    assert tl["lo"] == 15000
    assert tl["under"] == 7000

    page = crp.render(_ir(_body(8000)), rules, res)
    assert "下限 15,000" in page


def test_total_with_only_lower_bound_still_reported(tmp_path):
    """只登记下限、没登记上限的刊，下限照样进报告。"""
    rules = _rules(tmp_path, "tjyj", "tjyj_min_only.json",
                   **{"limits.total.value": None, "limits.total.min": 15000})
    res = crp.check(_ir(_body(8000)), rules)
    tl = res["total_limit"]

    assert tl["value"] is None
    assert tl["lo"] == 15000 and tl["under"] == 7000
    assert "≥15,000" in crp.render(_ir(_body(8000)), rules, res)


# ---------------------------------------------------------------- 标题上限（另一个键名）

def test_title_max_chars_key_is_read(tmp_path):
    """cjwt 风格 title_max_chars：报告要按它判标题超限。"""
    rules = _rules(tmp_path, "tjyj", "tjyj_title.json",
                   **{"limits.title_max_chars": 22})
    title = "数字金融发展对企业融资约束与创新产出的影响机制研究"
    assert len(title) > 22, "夹具前提：标题超过 22 字"
    ir = _ir({"type": "front_matter", "text": title})
    res = crp.check(ir, rules)
    ti = res["title_limit"]

    assert ti["max"] == 22
    assert ti["over"] == len(title) - 22
    assert "22 字" in crp.render(ir, rules, res)


# ---------------------------------------------------------------- 统计范围（cjwt 写法）

def test_body_with_chars_scope_counts_body_only(tmp_path):
    """「正文（含字符）」按正文口径统计，参考文献与摘要不算进去。"""
    rules = _rules(tmp_path, "tjyj", "tjyj_scope.json",
                   **{"limits.total.scope": "正文（含字符）"})
    ir = _ir({"type": "abstract", "text": "摘" * 300},
             _body(1000),
             {"type": "reference_item", "text": "文" * 300})
    res = crp.check(ir, rules)
    tl = res["total_limit"]

    assert tl["scope"] == "正文（含字符）"
    assert tl["scope_label"] == "正文"
    assert tl["actual"] == 1000, "退回全文口径了：摘要/文献被算进正文上限"

    page = crp.render(ir, rules, res)
    assert "正文（含字符）" in page, "规则集原声明应如实展示"


# ---------------------------------------------------------------- 规格页同口径

def test_spec_shows_total_min_inside_total(tmp_path):
    """render_spec 读 limits.total.min，显示区间而不是只剩上限。"""
    rules = _rules(tmp_path, "cjyj", "cjyj_spec.json")
    page = render_spec.build(rules)
    assert "15,000–25,000 字" in page


def test_spec_reads_top_level_total_min_and_title_max_chars(tmp_path):
    """render_spec 同样认 limits.total_min 与 title_max_chars。"""
    rules = _rules(tmp_path, "tjyj", "tjyj_spec.json",
                   **{"limits.total_min": 15000, "limits.title_max_chars": 22})
    page = render_spec.build(rules)
    assert "15,000–18,000 字" in page
    assert "不超过 22 字" in page
