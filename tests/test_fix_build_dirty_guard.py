"""便携包构建的脏树闸门：发布必须锚到一个已提交状态。

曾经的洞：打包发生在提交之前，包内容对应 da78114、BUILD-INFO 却记 6261d5d
——日常版没有精确的 git 锚点。现在工作区不干净直接拒绝构建（--allow-dirty
是显式逃生口），顺序只能是：改完 → 提交 → 再打包。
"""

import subprocess

from conftest import ROOT, TOOLS

import build_portable


def _patch_git(monkeypatch, result):
    monkeypatch.setattr(build_portable.subprocess, "run",
                        lambda *a, **k: result)


def test_dirty_tree_detected(monkeypatch):
    _patch_git(monkeypatch, subprocess.CompletedProcess([], 0, stdout=" M desktop.py\n"))
    assert build_portable.working_tree_dirty() is True


def test_clean_tree_passes(monkeypatch):
    _patch_git(monkeypatch, subprocess.CompletedProcess([], 0, stdout=""))
    assert build_portable.working_tree_dirty() is False


def test_git_failure_not_flagged_dirty(monkeypatch):
    """git 报错/非仓库（返回码非 0）不拦构建，维持旧行为。"""
    _patch_git(monkeypatch, subprocess.CompletedProcess([], 128, stdout=""))
    assert build_portable.working_tree_dirty() is False


def test_git_missing_not_flagged_dirty(monkeypatch):
    """git 不存在（异常路径）同样放行——闸门失灵时不能把构建彻底锁死。"""
    def boom(*a, **k):
        raise FileNotFoundError("git")
    monkeypatch.setattr(build_portable.subprocess, "run", boom)
    assert build_portable.working_tree_dirty() is False


def test_gate_wired_into_main_with_escape_hatch():
    """源码钉：闸门必须真的在 main 里且带 --allow-dirty 逃生口。"""
    src = (TOOLS / "build_portable.py").read_text(encoding="utf-8")
    assert "working_tree_dirty() and not a.allow_dirty" in src
    assert '"--allow-dirty"' in src
