"""发布链回归钉：notes 生成、tag 探测、上游同步与源码接线。

正式版只从 GitHub Releases 下发（发布纪律见 README「发布」节）：
--upload 是唯一外发开关，且要求 HEAD 与上游一致；构建锚定已提交状态
（脏树闸门双保险）。这里钉住纯逻辑与源码接线，不真连 GitHub。
"""

import subprocess

from conftest import TOOLS

import release


def test_release_notes_first_release_fallback():
    assert release.release_notes("") == "首次发布。"


def test_release_notes_lists_commits_since_tag(monkeypatch):
    r = subprocess.CompletedProcess([], 0, stdout="- fix: a\n- feat: b\n")
    monkeypatch.setattr(release.subprocess, "run", lambda *a, **k: r)
    assert release.release_notes("v0.9") == "- fix: a\n- feat: b"


def test_last_tag_empty_when_git_has_no_tags(monkeypatch):
    r = subprocess.CompletedProcess([], 128, stdout="")
    monkeypatch.setattr(release.subprocess, "run", lambda *a, **k: r)
    assert release.last_tag() == ""


def test_upstream_synced_detects_drift(monkeypatch):
    out = iter(["abc123\n", "def456\n"])
    monkeypatch.setattr(release.subprocess, "run",
                        lambda *a, **k: subprocess.CompletedProcess([], 0, stdout=next(out)))
    ok, head = release.upstream_synced()
    assert ok is False
    assert head == "abc123"


def test_upstream_synced_true_when_identical(monkeypatch):
    out = iter(["abc123\n", "abc123\n"])
    monkeypatch.setattr(release.subprocess, "run",
                        lambda *a, **k: subprocess.CompletedProcess([], 0, stdout=next(out)))
    ok, head = release.upstream_synced()
    assert ok is True and head == "abc123"


def test_source_pins():
    """源码钉：内联命令列表、显式 --upload 开关、上游核对与脏树闸门都在位。"""
    src = (TOOLS / "release.py").read_text(encoding="utf-8")
    assert 'subprocess.run([sys.executable, "tools/build_portable.py", "--zip"]' in src
    assert '"tools/build_installer.py", "--version", ver' in src
    assert "if not a.upload:" in src
    assert '"--upload"' in src
    assert "upstream_synced()" in src
    assert "working_tree_dirty()" in src
    assert src.count("shell=False") >= 6


def test_upload_stages_ascii_asset_names():
    """GitHub 资产名剥非 ASCII（v1.0.0 首发把「格式工坊-1.0.0.zip」剥成了
    「-1.0.0.zip」）：上传前必须先拷成 format-workshop-<ver> ASCII 名。"""
    src = (TOOLS / "release.py").read_text(encoding="utf-8")
    assert '"format-workshop-" + ver + ".zip"' in src
    assert '"format-workshop-setup-" + ver + ".exe"' in src
    assert "shutil.copy2(zip_path, up_zip)" in src
    assert "#格式工坊 " in src        # 中文留在显示标签
    assert "shutil.rmtree(stage" in src  # 暂存目录用完即清
