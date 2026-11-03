"""F58/F59 回归钉：测试公共助手契约与夹具独立性。

- F58 收口：conftest.minimal_doc（最小稿件骨架）、mutated_rules（复制入库
  规则集按点路径改键）替代了散在各承诺测试里的重复构造。这里钉住两个入口
  的行为契约——契约漂移（悄悄改字、碰了点名之外的键）比重复更危险。
- F59：test_numbering.build_headings_docx 的中文数字改字面量清单后，夹具
  数据不再经过被测函数 cn_numeral。钉法是毒化：把 cn_numeral 换成抛错，
  夹具照常产出字面量值——若有人回退成经被测函数造数，当场红。

夹具独立性通过 importlib 加载 test_numbering.py、直接调用它的
build_headings_docx 来对账——钉的就是那个夹具本身。渲染器加载需要 tools
在 sys.path，与 test_numbering.py 本身的姿势一致。
"""

import importlib.util
import sys
from pathlib import Path

from docx import Document

from conftest import RULES, minimal_doc, mutated_rules

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))


def _load_test_numbering():
    """按 test_numbering.py 的姿势加载它，拿 build_headings_docx 与 render。"""
    spec = importlib.util.spec_from_file_location(
        "test_numbering_for_f59", ROOT / "tests" / "test_numbering.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_minimal_doc_keeps_caller_text_verbatim(tmp_path):
    """F58 契约：minimal_doc 不悄悄换任何字，骨架顺序固定，keywords 可省。

    文字逐字由调用方给定是承诺测试的底线——助手若做任何归一（去空白、换
    标点），下游守卫的逐字比对就会失真。
    """
    path = tmp_path / "m.docx"
    doc = minimal_doc(path, title="标题甲", abstract="内容提要：正文乙。",
                      keywords="关键词：丙；丁", extra_paragraphs=["第一段。", "第二段。"])
    texts = [p.text for p in Document(str(path)).paragraphs]
    assert texts == ["标题甲", "内容提要：正文乙。", "关键词：丙；丁", "第一段。", "第二段。"]

    path2 = tmp_path / "m2.docx"
    minimal_doc(path2, title="标题甲", abstract="摘要乙。")
    texts2 = [p.text for p in Document(str(path2)).paragraphs]
    assert texts2 == ["标题甲", "摘要乙。"], "keywords=None 时不得冒出关键词段"
    assert doc is not None


def test_mutated_rules_touches_only_named_keys(tmp_path):
    """F58 契约：mutated_rules 只动点名的键，其余逐字保留（值与出处不碰）。

    规则集是带出处的手工数据，助手若在改键时丢了别的字段（出处、measured
    原值），测试就在悄悄伪造规则集。
    """
    rules_path = mutated_rules(tmp_path, "tjyj", "probe.json",
                               **{"citation.system": "numeric",
                                  "headings.levels.1": "1.",
                                  "typography.footnote": None})
    import json

    original = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    mutated = json.loads(rules_path.read_text(encoding="utf-8"))

    assert mutated["citation"]["system"] == "numeric"
    assert mutated["headings"]["levels"][1] == "1."
    assert "footnote" not in mutated["typography"]
    # 点名之外的键逐字原样
    assert mutated["submission"] == original["submission"]
    assert mutated["blocks"] == original["blocks"]
    assert mutated["citation"]["reference_list"] == original["citation"]["reference_list"]


def test_numbering_fixture_does_not_depend_on_cn_numeral(tmp_path):
    """F59 回归：夹具数据必须不经过被测函数 cn_numeral。

    test_numbering.build_headings_docx 若回头用 render.cn_numeral 造
    「一、／（一）」，它错了夹具跟着错，端到端用例永远假绿。钉法：把
    cn_numeral 毒化成抛错，夹具照常产出、文字仍是字面量清单里的值——
    依赖被测函数就会当场红。
    """
    tn = _load_test_numbering()
    original = tn.render.cn_numeral
    tn.render.cn_numeral = lambda n: (_ for _ in ()).throw(
        AssertionError("夹具不应调用被测函数 cn_numeral（F59）"))
    try:
        src = tmp_path / "poisoned.docx"
        tn.build_headings_docx(src, sections=3, subs_per_section=2)
        texts = [p.text for p in Document(str(src)).paragraphs]
        # 字面量清单的值逐字在场——夹具数据与 cn_numeral 的行为完全无关
        assert texts[1] == "一、第1章", f"第 1 章标题不对：{texts[1]!r}"
        assert texts[3] == "（一）第1章第1节", f"第 1 节标题不对：{texts[3]!r}"
        assert texts[13] == "三、第3章", f"第 3 章标题不对：{texts[13]!r}"
        assert texts[17] == "（二）第3章第2节", f"第 3 章第 2 节标题不对：{texts[17]!r}"
    finally:
        tn.render.cn_numeral = original
