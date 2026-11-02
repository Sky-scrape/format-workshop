"""P4-1 回归钉：引用体例判别闸（citation.system）。

修前：勾选闸只看「规则集有没有条目模板 + 稿子里有没有方括号」，cjyj 这类
声明编号实引（citation.system=numbered）的刊也被开出「转作者-年份」，勾选后
「…文献一[1]和文献二[2]。」被改成「…文献一（张三，2020）和文献二（Smith，
2021）。」——映射与审计都对，但对要求编号实引的刊，交付稿的引用体例直接错。

体例判别字段规则集里本来就有（citation.system，入库刊全部已声明，枚举
numbered / author-year），本轮补的是：journal-schema §4 的显式契约、
DESIGN §1.4 的开关边界、apply_citation 入口硬闸。app 勾选闸加同款条件归
串行阶段 D，不在本文件钉（「cjyj 不再 offered」的 UI 断言随 D 落地）。

钉住四件事：
1. 四本入库刊的 system 声明与其体例一致（cjyj/gjjryj=numbered，mwm/tjyj=author-year）
2. 编号实引刊（cjyj、ciejournal、gjjryj）勾了 --citation 也不转：正文与文献表
   逐字原样，审计记「该刊声明编号实引，不转」，rule 指回 citation.system
3. 作者-年份刊（mwm）勾了照常转换——硬闸不误伤
4. in_text_style 的引用层闸（既有 PB1 行为不退）：编号刊的 bracket「[]」
   是编号标注的写法，不进作者-年份括注
"""

import json
import sys
from pathlib import Path

import pytest
from docx import Document

from conftest import RULES, TOOLS, run_tool

sys.path.insert(0, str(TOOLS))
import citation_ref as cr  # noqa: E402

SCHEME = {"cjyj": "numbered", "gjjryj": "numbered",
          "mwm": "author-year", "tjyj": "author-year"}
# 编号制 + 有条目模板：修前勾选闸会开的两本（gjjryj 无模板，闸本来就关）
NUMBERED_WITH_TEMPLATES = ("cjyj", "ciejournal")


def test_declared_scheme_matches_journal_style():
    """体例判别字段的值必须与刊物体例一致，这是两道闸共同的判别依据。"""
    for jid, want in SCHEME.items():
        rules = json.loads((RULES / f"{jid}.json").read_text(encoding="utf-8"))
        got = rules["citation"]["system"]
        assert got == want, f"{jid} 的 citation.system={got!r}，与其体例（{want}）不符"


def build_doc(path):
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.add_paragraph("[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.")
    doc.save(str(path))
    return path


def run_with_citation(tmp_path, jid):
    """合成稿（编号引用 + 可双向对上的文献表）→ 目标刊渲染，勾选引用转换。"""
    d = Path(tmp_path)
    docx = build_doc(d / "manuscript.docx")
    run_tool("parse_docx.py", docx, "-o", d / "manuscript.ir.json")
    out = d / "out.docx"
    run_tool("render_from_ir.py", d / "manuscript.ir.json",
             RULES / f"{jid}.json", "-o", out, "--citation")
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    return [p.text for p in Document(str(out)).paragraphs], audit


@pytest.mark.parametrize("jid", [*NUMBERED_WITH_TEMPLATES, "gjjryj"])
def test_numbered_journal_never_converts(tmp_path, jid):
    """编号实引刊勾了也不转：正文与文献表原样，审计说明跳过原因。"""
    texts, audit = run_with_citation(tmp_path, jid)
    body = "\n".join(t for t in texts if t.startswith("正文"))
    assert "文献一[1]和文献二[2]" in body, f"{jid} 是编号实引刊，正文编号引用被改写"
    entries = [t for t in texts if t.strip().startswith("[")]
    assert any("张三. 数字金融与企业融资" in t for t in entries), \
        f"{jid} 的文献条目被改写"
    assert not [i for i in audit["items"] if i["action"] == "正文引用改作者-年份"], \
        f"{jid} 出现了作者-年份改写条目"
    skipped = [i for i in audit["items"] if i["action"] == "引用体系转换（已跳过）"]
    assert skipped, f"{jid} 勾选了转换却没有跳过说明"
    assert "编号实引" in skipped[0]["note"], skipped[0]["note"]
    assert skipped[0]["rule"] == "citation.system", skipped[0]["rule"]


def test_author_year_journal_still_converts(tmp_path):
    """对照：author-year 刊（mwm）勾了照常转换，硬闸不许误伤。"""
    texts, audit = run_with_citation(tmp_path, "mwm")
    body = "\n".join(t for t in texts if t.startswith("正文"))
    assert "文献一[1]和文献二[2]" not in body, "author-year 刊的转换被硬闸误伤"
    assert "（张三，2020）" in body and "（Smith，2021）" in body
    actions = [i["action"] for i in audit["items"]]
    assert "正文引用改作者-年份" in actions
    assert not [i for i in audit["items"] if i["action"] == "引用体系转换（已跳过）"]


def test_in_text_style_refuses_numbered_journal_bracket():
    """引用层既有闸（PB1）：编号刊的 bracket「[]」描述编号标注，不进作者-年份。"""
    rules = json.loads((RULES / "cjyj.json").read_text(encoding="utf-8"))
    (open_b, close_b), _sep = cr.in_text_style(rules)
    assert (open_b, close_b) == ("（", "）")
