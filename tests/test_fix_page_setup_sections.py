"""F32 回归钉：render_sample 的页面设置要遍历全部节。

问题单拆分：render_from_ir.py 端由「渲染与引用」路负责；本条目负责
render_sample.py 的同款修复，两路实现同一遍历姿势（for sec in doc.sections）。

此前 `sec = doc.sections[0]` 只给第一节按规则集设纸张与页边距，多节文档
其余节保持 Word 默认版面（Letter 215.9×279.4、边距 25.4/25.4/31.75/31.75）。

render_sample 从空白文档生成、正常输出单节，夹具在子进程里把初始文档造成
两节（等价于文档里存在分节符的情形）——修复前第二节永远停在 Word 默认值。
用例：
- 两节夹具 → 两节的纸张与四边距都等于规则集值；规则集值特意改成与
  Word 默认、与 tjyj 原值（210×297、22/20/25/24）都不同的数，
  证明数值读自规则集而非代码里的第二真源；
- 单节基准（CLI 全流程）→ 唯一一节等于入库规则集原值，遍历姿势
  不改变单节路径的既有产出。

夹具姿势与 test_citation_convert.py 相同：子进程直传参数列表、不经 shell；
Document 的两节补丁必须打在 build 所在的子进程内，父进程只断言成品。
"""

import json
import subprocess
import sys

import pytest
from docx import Document

from conftest import RULES, ROOT, TOOLS, run_tool

# 与 render_from_ir.page_diffs 同一容差：mm↔twips 往返舍入实测 ≤0.009mm
TOL = 0.05

# 子进程内执行的构建脚本：argv = [tools目录, 规则集路径, 输出路径, 模式]
BUILD_SCRIPT = """\
import json
import sys
from pathlib import Path

sys.path.insert(0, sys.argv[1])
import render_sample
from docx import Document as _fresh

if sys.argv[4] == "two":
    def two_section_document():
        d = _fresh()
        d.add_section()   # 克隆节属性（此刻还是 Word 默认）成第 0 节，sentinel 成末节
        return d
    render_sample.Document = two_section_document

rules = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
print("built:", render_sample.build(rules, sys.argv[3]))
"""


def section_mm(docx_path):
    """每节的纸张与四边距（mm），顺序即文档节序。"""
    doc = Document(str(docx_path))
    return [(s.page_width.mm, s.page_height.mm,
             s.top_margin.mm, s.bottom_margin.mm,
             s.left_margin.mm, s.right_margin.mm)
            for s in doc.sections]


def build_sample(tmp_path, name, mode, mutate=None):
    """子进程跑 render_sample.build；mode="two" 时初始文档带两节。"""
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    if mutate is not None:
        mutate(rules)
    rules_path = tmp_path / f"{name}.json"
    rules_path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")

    script = tmp_path / "build_sample.py"
    script.write_text(BUILD_SCRIPT, encoding="utf-8")
    out = tmp_path / f"{name}.docx"
    r = subprocess.run(
        [sys.executable, str(script), str(TOOLS), str(rules_path), str(out), mode],
        capture_output=True, text=True, cwd=str(ROOT))
    assert r.returncode == 0, f"{name} 构建失败：{r.stderr[-500:]}"
    return out


def distinct_page(rules):
    """把页面值改成与 Word 默认、tjyj 原值都不同的数——
    断言值只可能来自规则集（无第二真源）。"""
    rules["page"]["size_mm"]["width"] = 200
    rules["page"]["size_mm"]["height"] = 280
    for side, v in (("top", 30), ("bottom", 35), ("left", 40), ("right", 45)):
        rules["page"]["margin_mm"][side] = v


def test_page_setup_reaches_every_section(tmp_path):
    """F32 证据复现：两节文档的每一节都要按规则集设置纸张与四边距。"""
    out = build_sample(tmp_path, "two_sections", "two", mutate=distinct_page)

    secs = section_mm(out)
    assert len(secs) == 2, f"夹具应产出两节，实际 {len(secs)} 节"
    want = (200, 280, 30, 35, 40, 45)
    for i, sec in enumerate(secs):
        assert sec == pytest.approx(want, abs=TOL), \
            f"第 {i} 节页面值 {sec} ≠ 规则集值 {want}（只修 sections[0] 会漏掉其余节）"


def test_single_section_cli_output_matches_ruleset(tmp_path):
    """单节基准：CLI 全流程产出的唯一一节等于入库规则集原值。"""
    out = tmp_path / "single.docx"
    run_tool("render_sample.py", RULES / "tjyj.json", "-o", out)

    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    p = rules["page"]
    want = (p["size_mm"]["width"], p["size_mm"]["height"],
            p["margin_mm"]["top"], p["margin_mm"]["bottom"],
            p["margin_mm"]["left"], p["margin_mm"]["right"])
    secs = section_mm(out)
    assert len(secs) == 1
    assert secs[0] == pytest.approx(want, abs=TOL), \
        f"单节页面值 {secs[0]} ≠ 规则集值 {want}"


def _two_section_manuscript(path):
    """带分节符的最小稿件：第一节标题，第二节一段正文。"""
    from docx import Document as _fresh
    doc = _fresh()
    doc.add_paragraph("数字金融与商业银行风险承担")
    doc.add_section()
    doc.add_paragraph("本文研究数字金融对商业银行风险承担的影响。")
    doc.save(str(path))


def test_render_from_ir_page_setup_reaches_every_section(tmp_path):
    """F32 复验盲区补钉：主渲染器的页面设置同样要遍历全部节。

    回炉复验只核了 render_sample.py；render_from_ir.py 的 sections[0]
    漏网（分节稿其余节保持 Word 默认版面，审计却只有一条「页面设置」）。
    """
    from conftest import render_into

    src = tmp_path / "twosec.docx"
    _two_section_manuscript(src)
    pipe = render_into(tmp_path / "work", src, target="tjyj")

    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    p = rules["page"]
    want = (p["size_mm"]["width"], p["size_mm"]["height"],
            p["margin_mm"]["top"], p["margin_mm"]["bottom"],
            p["margin_mm"]["left"], p["margin_mm"]["right"])
    secs = section_mm(pipe.docx)
    assert len(secs) == 2, f"夹具应产出两节，实际 {len(secs)} 节"
    for i, sec in enumerate(secs):
        assert sec == pytest.approx(want, abs=TOL), \
            f"第 {i} 节页面值 {sec} ≠ 规则集值 {want}（主渲染器只设 sections[0] 会漏掉其余节）"
    page_items = [it for it in pipe.audit["items"] if it["action"] == "页面设置"]
    assert page_items and "全部 2 节" in page_items[0]["note"], \
        "页面设置条目应注明作用范围（应用于全部 N 节）"
