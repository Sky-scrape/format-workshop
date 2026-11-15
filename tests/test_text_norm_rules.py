"""N-02/N-03/N-04/N-06 字符级归一（2026-11-15 落地）。

守四件事：
1. 四条规则按语境生效，界线不被放宽（中拉之间空格、千分位、西文标点不动）；
2. 渲染器真的改到成品里，且逐块进审计、可在报告里逐条驳回；
3. 驳回一条只回退那一条，其余归一照常；
4. 交付前的文字闸与渲染器共用 tools/text_norm.py 的同一实现。
"""

import json
import sys
from pathlib import Path

from docx import Document

from conftest import RULES, invariant_failures, minimal_doc, render_into

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))

import text_norm  # noqa: E402

ACTIONS = ["正文连续空格压缩", "数字与中文间空格删除",
           "中文标点半角转全角", "英文标点全角转半角"]


def _rules(base="tjyj"):
    return json.loads((RULES / f"{base}.json").read_text(encoding="utf-8"))


# ---------------------------------------------------------------- 纯文本口径

def test_rules_respect_context_boundaries():
    """语境判定：中文句子转全角，西文句子转半角，跨界的一律不动。"""
    assert text_norm.cn_punct("研究发现, 结论") == "研究发现， 结论"
    assert text_norm.cn_punct("Smith, 2020") == "Smith, 2020"
    assert text_norm.cn_punct("占比 1,000 家") == "占比 1,000 家"
    assert text_norm.cn_punct("12:30 开会") == "12:30 开会"
    assert text_norm.en_punct("(Smith，2020)") == "(Smith, 2020)"
    assert text_norm.en_punct("A；B") == "A; B"
    assert text_norm.en_punct("版本 3．14") == "版本 3.14"
    assert text_norm.en_punct("中文，句子") == "中文，句子"


def test_spaces_only_between_chinese_and_digits():
    """N-02/N-06 的边界：中拉之间与全角空格不动。"""
    assert text_norm.drop_cjk_spaces("样本 共 1200 个") == "样本共1200个"
    assert text_norm.drop_cjk_spaces("数字金融 A 股") == "数字金融 A 股"
    assert text_norm.drop_cjk_spaces("摘　要") == "摘　要"
    # 段首「全角空格＋半角空格」是手工缩进，归 N-01 管；N-06 不许吃掉
    # 后面那个半角空格，否则驳回段首缩进删除后原文仍然被改。
    assert text_norm.drop_cjk_spaces("\u3000 稳健性检验") == "\u3000 稳健性检验"
    assert text_norm.collapse_spaces("研究  发现") == "研究 发现"
    # 段首/段尾不归 N-02：驳回段尾删除后不能又被压成一个空格
    assert text_norm.collapse_spaces("正文  ") == "正文  "


def test_normalize_is_idempotent():
    samples = ["研究发现,  数字金融 显著缓解了融资约束 。样本 共 １２００ 个。",
               "(Smith，2020) shows X", "数字金融 A 股", "占比 1,000 家"]
    for s in samples:
        once = text_norm.normalize(s)
        assert text_norm.normalize(once) == once


# ---------------------------------------------------------------- 渲染与审计

def _doc(path):
    minimal_doc(
        path,
        abstract="内容提要：Abstract: This study, based on A-share data，finds results.",
        keywords="关键词：数字金融；融资约束；企业创新",
        extra_paragraphs=[
            "一、引言",
            "研究发现,  数字金融 显著缓解了融资约束 。样本 共 １２００ 个, 详见 Smith，2020。",
        ])
    return path


def test_renderer_applies_all_four_rules_and_audits(tmp_path):
    p = render_into(tmp_path / "norm", _doc(tmp_path / "norm.docx"), target="tjyj")
    body = next(t for t in p.output_texts.values()
                if "数字金融显著缓解" in t)
    assert body == "研究发现，数字金融显著缓解了融资约束。样本共1200个，详见 Smith, 2020。"
    assert "A-share data, finds results." in " ".join(p.output_texts.values())

    got = {it["action"] for it in p.items}
    assert set(ACTIONS) <= got, f"四条归一没有全部进审计：{got}"

    # 每条都要能解析回规则集字段，且记了前后片段
    for it in p.items:
        if it["action"] in ACTIONS:
            assert it.get("rule") == "typography.body"
            assert it.get("before") or it.get("after")

    problems = invariant_failures(p)
    assert not problems, "文字闸把已落地的白名单归一判成了改字：" + "；".join(problems)


def test_reject_one_rule_keeps_the_others(tmp_path):
    src = _doc(tmp_path / "reject.docx")
    first = render_into(tmp_path / "first", src, target="tjyj")
    body_idx = next(i for i, t in first.output_texts.items()
                    if "数字金融显著缓解" in t)
    items = [it for it in first.items
             if it["idx"] == body_idx and it["action"] == "中文标点半角转全角"]
    assert len(items) == 1, "夹具里应恰有一条中文标点归一"
    rejected = render_into(tmp_path / "rejected", src, target="tjyj",
                           rejects=[items[0]["id"]])
    body = next(t for t in rejected.output_texts.values()
                if "数字金融显著缓解" in t)
    # 驳回的那条保留半角逗号；空格、数字与英文标点照常归一
    assert body == "研究发现, 数字金融显著缓解了融资约束。样本共1200个, 详见 Smith, 2020。"
