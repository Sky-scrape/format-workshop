"""第四承诺：规则集是唯一真源。

排版器里只要藏一个写死的字号，样板就会与规则集悄悄分家——而用户看到的
规格页说 10.5pt，实际下载的是 12pt，这比不给规格页更糟。
"""

import ast
import json
import subprocess

import pytest

from conftest import (RULES, SOURCE_JOURNAL, TARGET_JOURNAL, TOOLS, render_into,
                      run_sizes)


def test_fixtures_use_shipped_journals():
    """夹具只能用仓库里真实存在的规则集。

    完整规则库不入库（见 `.gitignore`），如果哪天夹具换成了私有刊，本机照样
    跑得通、CI 必挂——这个坑踩过一次，用用例钉住。
    """
    tracked = subprocess.run(["git", "ls-files", "rules/journal"],
                             capture_output=True, text=True)
    if tracked.returncode != 0:
        pytest.skip("不在 git 工作区内，无法判断规则集是否已入库")

    for jid in (SOURCE_JOURNAL, TARGET_JOURNAL):
        assert f"rules/journal/{jid}.json" in tracked.stdout, (
            f"{jid}.json 没有入库：夹具用了私有规则集，CI 会失败")


def test_no_hardcoded_format_values_in_renderers():
    """渲染相关脚本里不得出现 Pt(10.5) / Mm(25) 这类字面量。

    格式值必须从规则集读。这条是静态检查，跑得比渲染快，能挡住「先图省事
    写死，回头再改」。
    """
    offenders = []
    for py in sorted(TOOLS.glob("*.py")):
        tree = ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id in {"Pt", "Mm"} and node.args):
                arg = node.args[0]
                if isinstance(arg, ast.Constant) and isinstance(arg.value, (int, float)):
                    offenders.append(f"{py.name}:{node.lineno} {node.func.id}({arg.value})")
    assert not offenders, "发现硬编码格式值：\n" + "\n".join(offenders)


def test_changing_the_ruleset_changes_the_output(baseline, sample_docx, tmp_path):
    """把正文号改大 1.5pt，成品正文就必须跟着变——证明这个值确实来自规则集。"""
    rules = json.loads((RULES / f"{TARGET_JOURNAL}.json").read_text(encoding="utf-8"))
    original = rules["typography"]["body"]["size_pt"]
    rules["typography"]["body"]["size_pt"] = original + 1.5

    alt_rules = tmp_path / "alt.json"
    alt_rules.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")

    alt = render_into(tmp_path / "alt", sample_docx, rules_path=alt_rules)

    body_idx = {i for i, t in alt.block_type.items() if t == "body"}
    sizes = run_sizes(alt.docx)
    measured = {i: sizes[i] for i in body_idx if i in sizes}
    assert measured, "成品里没有可测量的正文段落"

    expected = round(original + 1.5, 2)
    wrong = {i: s for i, s in measured.items() if s != {expected}}
    assert not wrong, f"这些正文段落没跟着规则集变：{list(wrong.items())[:3]}"

    base_sizes = run_sizes(baseline.docx)
    base_body = {i: base_sizes[i] for i in body_idx if i in base_sizes}
    assert all(s == {round(original, 2)} for s in base_body.values()), \
        "基准稿的正文号与规则集原值不符"
