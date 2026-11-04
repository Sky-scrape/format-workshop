"""D-12 回归钉：app_ctl 的入口行为与自述一致，开发版判据不再靠隐性端口。

曾经的漂移：
- 用法块写「没在跑就静默启动，已在跑就直接打开页面」，实现里根本没有
  「已在跑」的判定，与同文件「壳不占端口」的自述矛盾；
- main() 无参默认是浏览器服务模式，裸跑 python tools/app_ctl.py 会静默
  拉起 keepalive 常驻，与「桌面壳不使用端口/守护」相悖；
- shortcut 的开发版判据只剩 PAPER_PORT != 5000（开发版.bat 顺手 export 的），
  开发用户直接调 shortcut 会建「格式工坊.lnk」盖掉使用版的同名快捷方式。

现在无参默认 = 拉起桌面壳（launch_shell），浏览器模式要显式 start/quiet；
开发版判据与 desktop.py 的窗口标题同源（ROOT/.venv）。
"""

import subprocess
import sys
import types
from pathlib import Path

import pytest

from conftest import ROOT, TOOLS

import app_ctl


@pytest.fixture()
def popen_calls(monkeypatch):
    """替身 Popen：只记录调用，不真拉进程。"""
    calls = []
    monkeypatch.setattr(app_ctl.subprocess, "Popen",
                        lambda cmd, **kw: calls.append((cmd, kw)) or object())
    return calls


@pytest.fixture()
def no_start(monkeypatch):
    """浏览器服务模式的 start 换成探针，证明无参路径没走到它。"""
    seen = []
    monkeypatch.setattr(app_ctl, "start", lambda **kw: seen.append(kw) or 0)
    return seen


def _run_main(monkeypatch, *args):
    monkeypatch.setattr(app_ctl.sys, "argv", ["app_ctl.py", *args])
    return app_ctl.main()


def test_main_no_args_launches_the_desktop_shell(monkeypatch, popen_calls, no_start):
    """无参默认 = 桌面壳：pythonw desktop.py，不起 keepalive，不碰 start()。"""
    rc = _run_main(monkeypatch)

    assert rc == 0
    assert no_start == [], "无参不得再静默进入浏览器服务模式"
    assert len(popen_calls) == 1
    cmd, kw = popen_calls[0]
    assert cmd[-1] == str(ROOT / "desktop.py")
    assert "keepalive" not in " ".join(cmd), "壳不再经 keepalive"
    assert kw.get("cwd") == str(ROOT)
    assert kw.get("creationflags", 0) != 0, "Windows 下应 detached 拉起，控制台不跟着挂"


def test_browser_service_mode_is_explicit_now(monkeypatch, no_start):
    """start/quiet 显式子命令仍是浏览器服务模式（open_browser 各归各位）。"""
    assert _run_main(monkeypatch, "start") == 0
    assert no_start == [{"open_browser": True}]
    assert _run_main(monkeypatch, "quiet") == 0
    assert no_start[-1] == {"open_browser": False}


def test_stop_and_status_dispatch_unchanged(monkeypatch):
    """stop/status 的子命令别名不许因入口重排而丢失。"""
    for argv, target in ((("stop",), "stop"), (("停止",), "stop"),
                         (("status",), "status"), (("状态",), "status")):
        stub = types.SimpleNamespace()
        monkeypatch.setattr(app_ctl, target, lambda *a, **k: 7)
        assert _run_main(monkeypatch, *argv) == 7


def test_dev_checkout_judged_by_venv_dir(tmp_path):
    """开发版判据 = 源码目录带 .venv（与 desktop.py 的窗口标题同源）。"""
    assert app_ctl.is_dev_checkout(tmp_path) is False
    (tmp_path / ".venv").mkdir()
    assert app_ctl.is_dev_checkout(tmp_path) is True


def test_shortcut_no_longer_reads_the_port_for_its_dev_decision():
    """shortcut 判据不许再依赖 PAPER_PORT（隐性信号），钉住判据本身。"""
    src = (TOOLS / "app_ctl.py").read_text(encoding="utf-8")
    assert "dev = is_dev_checkout()" in src
    assert "PORT != 5000" not in src
    # 与 desktop.py 同源：同样的判据表达式
    assert '(root / ".venv").is_dir()' in src


def test_usage_doc_matches_the_entry_behavior():
    """用法块不许再描述不存在的「已在跑」判定与静默服务启动。"""
    src = (TOOLS / "app_ctl.py").read_text(encoding="utf-8")
    assert "已在跑就直接打开页面" not in src
    assert "没在跑就静默启动" not in src
    assert "launch_shell" in src, "无参默认拉起桌面壳要在头注里可查"


def test_desktop_shell_criterion_is_the_same_everywhere():
    """壳的形态判据只此一种：desktop.py 与 app_ctl 都是 ROOT/.venv。"""
    desktop_src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    assert '(ROOT / ".venv").is_dir()' in desktop_src
