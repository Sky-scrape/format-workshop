"""只读安装下的数据目录回退。

这是真实踩过的坑：安装器把程序铺进 Program Files 后，程序仍按
``ROOT/outputs``、``ROOT/config`` 写任务、日志与配置；普通用户双击启动
就 PermissionError。这里守着“程序目录不可写时自动落用户目录”。
"""

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import app_paths  # noqa: E402


@pytest.fixture(autouse=True)
def _clear_cache():
    app_paths._CACHE.clear()
    yield
    app_paths._CACHE.clear()


def test_writable_project_keeps_data_in_place(tmp_path):
    """绿色包/开发版：程序目录可写就继续用它，数据跟着软件走。"""
    assert app_paths.data_root(tmp_path) == tmp_path
    assert not list(tmp_path.glob(".write-probe-*")), "探测文件必须清掉"


def test_read_only_project_falls_back_to_localappdata(tmp_path, monkeypatch):
    """Program Files 不可写 → 落到 %LOCALAPPDATA%\\格式工坊。"""
    root = tmp_path / "Program Files" / "格式工坊"
    root.mkdir(parents=True)
    local = tmp_path / "local"
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    monkeypatch.setattr(app_paths, "is_writable", lambda p: Path(p) != root)

    assert app_paths.data_root(root) == local / "格式工坊"
    assert (local / "格式工坊").is_dir()


def test_env_override_wins(tmp_path, monkeypatch):
    """PAPER_DATA_DIR 强制指定数据目录（测试/多份并存）。"""
    target = tmp_path / "data"
    monkeypatch.setenv("PAPER_DATA_DIR", str(target))

    assert app_paths.data_root(tmp_path) == target
    assert target.is_dir()


def test_no_writable_place_raises_with_actionable_message(tmp_path, monkeypatch):
    """两处都不可写时明确报错，而不是运行到一半丢稿件。"""
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    monkeypatch.setattr(app_paths, "is_writable", lambda p: False)

    with pytest.raises(OSError, match="没有可写的数据目录"):
        app_paths.data_root(tmp_path)


def test_app_work_follows_data_root():
    """app 的任务目录必须挂在数据根下，不能再写死 ROOT/outputs。"""
    import app

    assert app.WORK == app.DATA_ROOT / "outputs" / "web"
