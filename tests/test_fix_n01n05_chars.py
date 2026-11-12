"""N-01 段尾空白与 N-05 全角字母数字归一（2026-11-13 落地）。

两条都是 DESIGN §4.1 白名单内的 B 档字符归一：逐条进审计、可驳回，
交付前的运行时文字闸（tools/text_guard.py）按同一口径放行。本文件守：

  - 渲染器真的删了段尾空白、真的把全角字母数字转了半角，正文其余字符不动；
  - 两条动作各出一条审计，依据能解析回规则集；
  - 驳回后原文保留（可重复）；
  - 守卫的 normalize 认这两项白名单，但不放宽正文内部空格。
"""

import json
import sys
from pathlib import Path

from docx import Document

from conftest import RULES, render_into

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from rules_ref import resolvable  # noqa: E402
import text_guard  # noqa: E402

TRAIL_ACTION = "段尾空白删除"
ALNUM_ACTION = "全角字母数字归一"


def build_doc(path):
    """一份含段尾空白与全角数字的稿子：摘要、正文各一处。"""
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察２０２０年数字金融的影响。  ")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("一、引言")
    doc.add_paragraph("本文使用２０２０年的Ａ股样本，研究发现融资约束显著下降。  ")
    doc.add_paragraph("稳健性检验表明上述结论在更换度量方式后依然成立。")
    doc.save(str(path))
    return path


def _body_idx(result, needle):
    return next(i for i, t in result.source_texts.items() if needle in t)


def test_trailing_spaces_stripped_and_audited(tmp_path):
    src = build_doc(tmp_path / "chars.docx")
    result = render_into(tmp_path / "run", src, target="tjyj")
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))

    idx = _body_idx(result, "融资约束显著下降")
    before = result.source_texts[idx]
    after = result.output_texts[idx]
    assert before.endswith("  ")
    assert not after.endswith((" ", "\t", "\u3000")), "段尾空白没有删掉"
    assert after == before.rstrip(" \t\u3000").replace("２０２０", "2020") \
        .replace("Ａ股", "A股"), "除段尾空白与全角数字外还改了别的字符"

    items = [it for it in result.items if it["action"] == TRAIL_ACTION and it["idx"] == idx]
    assert len(items) == 1, "段尾空白删除没有逐块进审计"
    assert resolvable(rules, items[0]["rule"]), f"依据解析不到：{items[0]['rule']}"
    assert items[0]["before"] != items[0]["after"]


def test_fullwidth_alnum_converted_and_audited(tmp_path):
    src = build_doc(tmp_path / "chars.docx")
    result = render_into(tmp_path / "run", src, target="tjyj")
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))

    idx = _body_idx(result, "融资约束显著下降")
    items = [it for it in result.items if it["action"] == ALNUM_ACTION and it["idx"] == idx]
    assert len(items) == 1, "全角字母数字归一没有逐块进审计"
    assert "２０２０" not in result.output_texts[idx]
    assert "2020" in result.output_texts[idx]
    assert resolvable(rules, items[0]["rule"]), f"依据解析不到：{items[0]['rule']}"
    assert "２０２０" in items[0]["before"] or "２０２０" in items[0]["after"]

    # 摘要段同样适用（标签后的正文一起归一）
    abs_idx = _body_idx(result, "本文考察")
    assert "２０２０" not in result.output_texts[abs_idx]


def test_reject_keeps_original_trailing_and_fullwidth(tmp_path):
    src = build_doc(tmp_path / "chars.docx")
    first = render_into(tmp_path / "first", src, target="tjyj")
    idx = _body_idx(first, "融资约束显著下降")
    ids = [it["id"] for it in first.items
           if it["idx"] == idx and it["action"] in (TRAIL_ACTION, ALNUM_ACTION)]
    assert len(ids) == 2, "夹具里应恰有段尾与全角两条可驳回条目"

    rejected = render_into(tmp_path / "rejected", src, target="tjyj", rejects=ids)
    assert rejected.output_texts[idx] == first.source_texts[idx], \
        "驳回后没有完整保留原文"
    statuses = [it for it in rejected.items
                if it["idx"] == idx and it["action"] in (TRAIL_ACTION, ALNUM_ACTION)]
    assert all(it.get("rejected") for it in statuses), "驳回条目没有标成已驳回"


def test_guard_accepts_new_whitelist_without_loosening_inner_spaces():
    assert text_guard.normalize("本文２０２０年  ", "body") == "本文2020年"
    # 正文内部空格仍不在白名单：不会因为新加的两项被顺带吃掉
    assert text_guard.normalize("数字金融 A 股", "body") != \
        text_guard.normalize("数字金融A股", "body")
    # 非正文块不吃 N-05/N-01 段尾：文献条目只按 N-07 剥编号前缀，
    # 其后的全角字符与段尾空白原样保留
    assert text_guard.normalize("［１］张三  ", "reference_item") == "张三  "
    assert text_guard.normalize("标题１２", "heading1") == "标题１２"
    # 题注只抹编号前缀，不做全角字母数字归一
    assert text_guard.normalize("表１ ＡＢ", "table_caption") == "ＡＢ"
