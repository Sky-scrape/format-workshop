"""D-08 回归钉：分发包不带内部实测报告等私有材料。

dist\\格式工坊\\ 根曾实测存在 测试报告-2026-11-01.md，首段引用真实投稿稿
（训练论文/目标落空之后.docx）与本机完整规则库规模（23 本）——都是仓库
刻意不公开的私有素材。build_portable 的 EXCLUDE_FILES 没有它，逐份追名单
不如给根目录散落文档立白名单（README/DESIGN/CONTRIBUTING，LICENSE 无后缀
不受影响）；build_installer 对旧构建残留的 dist 再清一遍。
"""

from pathlib import Path

import pytest

from conftest import ROOT, TOOLS

from package_docs import ROOT_DOC_WHITELIST, is_private_doc, strip_private_docs
import build_portable


def test_the_internal_report_is_excluded():
    """实测报告按私有文档处理；三份项目文档照常进包。"""
    assert is_private_doc(("测试报告-2026-11-01.md",), "测试报告-2026-11-01.md")
    for name in ("README.md", "DESIGN.md", "CONTRIBUTING.md"):
        assert not is_private_doc((name,), name), f"{name} 是项目文档，要进包"


def test_only_root_level_docs_are_whitelisted():
    """子目录的 .md（rules/ 等）不在此列，别把规则集说明误伤。"""
    assert not is_private_doc(("rules", "readme.md"), "readme.md")
    assert not is_private_doc(("docs", "x.md"), "x.md")   # docs 整目录另有排除


def test_build_portable_excludes_it_via_is_excluded():
    """接到 is_excluded 上才是真的排除（单测 is_private_doc 不算数）。"""
    assert build_portable.is_excluded(Path("测试报告-2026-11-01.md")) is True
    assert build_portable.is_excluded(Path("README.md")) is False
    assert build_portable.is_excluded(Path("LICENSE")) is False, "AGPL 要随包携带"


def test_strip_private_docs_cleans_stale_dists(tmp_path):
    """旧构建残留的包：散落的私有 md 删掉，白名单与子目录不动。"""
    portable = tmp_path / "格式工坊"
    (portable / "rules").mkdir(parents=True)
    for name in ("README.md", "DESIGN.md", "CONTRIBUTING.md",
                 "测试报告-2026-11-01.md"):
        (portable / name).write_text("x", encoding="utf-8")
    (portable / "rules" / "notes.md").write_text("keep", encoding="utf-8")

    strip_private_docs(portable)

    remaining = {p.name for p in portable.glob("*.md")}
    assert remaining == ROOT_DOC_WHITELIST
    assert (portable / "rules" / "notes.md").exists()


def test_installer_cleans_stale_dists_before_packing():
    """build_installer 要对 dist 残留动手（它直接铺进 Program Files）。"""
    src = (TOOLS / "build_installer.py").read_text(encoding="utf-8")
    assert "strip_private_docs(portable)" in src


def test_repo_root_audit():
    """根目录审计：白名单之外的散落 .md 只允许是那份已知报告。

    再往根目录加内部文档时，这条会红——提醒把它加进排除（或白名单）。
    """
    offenders = {p.name for p in ROOT.glob("*.md")} - ROOT_DOC_WHITELIST
    assert offenders <= {"测试报告-2026-11-01.md"}, \
        f"根目录出现新的非白名单文档：{offenders}"
