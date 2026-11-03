"""本机全库矩阵（F43）：把「夹具只用 3 本入库刊」的教训制度化。

conftest 的运行时守卫只阻止夹具用私有刊（CI 上会 FileNotFoundError），它不利用
本机私有刊扩大覆盖——既有矩阵（test_manuscript_matrix）因此永远只测入库的 3 本。
本文件把整库拉进回归：

1. 样张矩阵——对 rules/journal/ 下每个存在的规则集跑 render_sample.build
   （纯 Python，无 LibreOffice），断言不崩，且样张里的显式字号全部落在 Word 的
   半磅刻度上（w:sz/w:szCs 以半磅计，值必须是整数；规则集侧的半磅约束由
   test_size_grid.py 在同一次定向运行里把守，这里守的是渲染落地端，含
   styles.xml 与任何绕过 python-docx 字号属性的原始 XML 写入路径）。
2. 流水线矩阵——对每本规则集，用合成稿（conftest.sample_docx，项目约定的
   源稿形态）走 parse→render→invariant_failures，四条不变式全过。

参数化按「文件存在才跑」：本机 23 本全跑，CI 只有入库的 3 本也能过。
全库规模的下限守卫（test_full_library_coverage_floor）防目录被清空或私有库
缩水时矩阵静默归零。
"""

import json
import xml.etree.ElementTree as ET
import zipfile

import pytest
from docx.oxml.ns import qn

import render_sample
from conftest import NS_W, RULES, invariant_failures, render_into, tracked_rulesets

# README 与 DESIGN 记载全库 23 本。下限取 20，容忍个别刊被有意移出库，
# 但绝不允许缩水到入库的 3 本而无人察觉。
FULL_LIBRARY_FLOOR = 20

SIZE_TAGS = ("w:sz", "w:szCs")


def ruleset_paths():
    """本机盘上实有的规则集（文件存在才参数化）。"""
    return sorted(RULES.glob("*.json"))


def test_full_library_coverage_floor():
    """覆盖数守卫：矩阵不允许静默缩水。

    - 入库规则集必须在盘上——目录被清空/误删时这里先炸（CI 与本机都守），
      否则两条矩阵参数化为空集，「全绿」是假象。
    - 盘上规则集多于入库集合，说明这台机器带着私有库（本机 23 > 入库 3），
      覆盖数必须达到全库规模下限。纯 CI 克隆（盘上==入库）不做此断言，
      它本来只有入库的 3 本。
    """
    tracked = tracked_rulesets()
    if tracked is None:
        pytest.skip("git 不可用，无法对照入库规则集")
    on_disk = {p.stem for p in ruleset_paths()}
    missing = sorted(tracked - on_disk)
    assert not missing, f"入库规则集不在盘上（目录被清空或误删）：{missing}"
    if len(on_disk) > len(tracked):
        assert len(on_disk) >= FULL_LIBRARY_FLOOR, (
            f"本机规则库应达全库规模（≥{FULL_LIBRARY_FLOOR} 本），"
            f"实际只有 {len(on_disk)} 本：{sorted(on_disk)}")


def explicit_sizes(docx_path):
    """document.xml 与 styles.xml 里全部显式 w:sz/w:szCs 的取值。

    直接读 XML 而不走 python-docx 的字号属性：后者从半磅整数反推，读回来
    永远在刻度上，测不出问题；前者连原始写入路径一起守。自家渲染器产出的
    XML，解析前仍拒绝 DOCTYPE/ENTITY（与 conftest.layout_signature 同一口径）。
    """
    vals = []
    with zipfile.ZipFile(str(docx_path)) as z:
        for name in ("word/document.xml", "word/styles.xml"):
            raw = z.read(name)
            if b"<!DOCTYPE" in raw or b"<!ENTITY" in raw:
                raise ValueError(f"{name} 带有 DOCTYPE/ENTITY 声明，拒绝解析")
            root = ET.fromstring(raw)
            for tag in SIZE_TAGS:
                for el in root.iter(NS_W + tag.split(":")[1]):
                    vals.append((name, tag, el.get(qn("w:val"))))
    return vals


def on_half_point_grid(val):
    """w:sz 的 val 以半磅计（24 即 12pt）：必须是正整数半磅。"""
    try:
        hp = float(val)
    except (TypeError, ValueError):
        return False
    return hp > 0 and hp.is_integer()


@pytest.mark.parametrize("rules_path", ruleset_paths(), ids=lambda p: p.stem)
def test_sample_build_full_library(rules_path, tmp_path):
    """①全库样张矩阵：每本规则集渲染样张不崩，字号全在半磅刻度上。"""
    rules = json.loads(rules_path.read_text(encoding="utf-8"))
    out = tmp_path / "sample.docx"
    render_sample.build(rules, str(out))
    assert out.exists() and out.stat().st_size > 0, f"{rules_path.stem} 样张未落盘"
    bad = [(name, tag, val) for name, tag, val in explicit_sizes(out)
           if not on_half_point_grid(val)]
    assert not bad, f"{rules_path.stem} 样张有非半磅字号：{bad}"


@pytest.mark.parametrize("rules_path", ruleset_paths(), ids=lambda p: p.stem)
def test_pipeline_full_library(sample_docx, tmp_path, rules_path):
    """②全库流水线矩阵：合成稿过 parse→render，四条不变式全过。

    rules_path 显式传给 render_into——conftest 的私有刊守卫因此按设计放行
    （守卫挡的是夹具隐式依赖私有刊导致 CI 挂；这里的目标刊来自盘上实有文件，
    CI 上这些参数化用例根本不会生成）。
    """
    p = render_into(tmp_path / f"mx-{rules_path.stem}", sample_docx,
                    target=rules_path.stem, rules_path=rules_path)
    problems = invariant_failures(p)
    assert not problems, f"{rules_path.stem}: " + "；".join(problems)
