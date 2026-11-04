r"""D-04 回归钉：控制面必须看得见桌面壳。

壳（pythonw desktop.py）既不监听端口、命令行也不叫 app.py——app_ctl 的
stop 只圈 keepalive/app\.py、status 的 proc_report 只过滤这两样，壳在跑时
status 输出「服务未运行」的完整故障相、stop 报「已停止」而壳原地不动；
更重的是覆盖安装/卸载链：installer.iss 两处钩子只调 app_ctl stop，壳进程
圈不住，{app}\python 被运行中的壳占用删不动。

修法：stop 的 kill_by_cmdline 加第三个 pattern（r"desktop\.py"，-Root 圈定
不变，与 F48 同款）；proc_report.ps1 过滤加 desktop.py 并输出 shell 角色，
status 如实显示壳 pid。iss 钩子本身不用改——stop 圈住壳它们就生效。
"""

import shutil
import subprocess
import types
from pathlib import Path

import pytest

from conftest import ROOT, TOOLS

import app_ctl


def _record_ps1(monkeypatch):
    calls = []
    monkeypatch.setattr(
        app_ctl, "run_ps1_full",
        lambda name, *args: calls.append((name, *args)) or ("", 2))
    return calls


def _quiet(monkeypatch):
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: False)
    monkeypatch.setattr(app_ctl, "time",
                        types.SimpleNamespace(sleep=lambda *_: None))


def test_stop_scopes_the_shell_with_the_same_root_rule(monkeypatch):
    """stop 要圈三类进程，壳的 pattern 与 -Root 圈定和守护/服务同款。"""
    calls = _record_ps1(monkeypatch)
    _quiet(monkeypatch)

    assert app_ctl.stop() == 0
    scoped = [c for c in calls if c[0] == "kill_by_cmdline.ps1"]
    assert ("kill_by_cmdline.ps1", "-Pattern", r"desktop\.py",
            "-Root", str(ROOT)) in scoped, "桌面壳必须被 -Root 圈定地结束"
    assert {c[2] for c in scoped} == {
        "keepalive", r"app\.py", r"desktop\.py"}, "三类 pattern 一个不能少"


def test_stop_reports_the_shell_label(monkeypatch, capsys):
    """壳被杀掉时报「桌面壳进程已停止」，不再只报守护/服务。"""
    monkeypatch.setattr(app_ctl, "run_ps1_full",
                        lambda name, *args: ("killed 4321\n", 0))
    _quiet(monkeypatch)

    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert "桌面壳进程已停止（pid 4321）" in out


def test_status_shows_the_shell_row(monkeypatch, capsys):
    """status 的进程段要显示壳：role=shell → 「壳」标签 + pid。"""
    monkeypatch.setattr(app_ctl, "run_ps1",
                        lambda *a: "111|shell|explorer.exe\n")
    monkeypatch.setattr(app_ctl, "listener_pids", lambda *a, **k: set())

    assert app_ctl.status() == 0
    out = capsys.readouterr().out
    assert "[2] 守护、服务与桌面壳进程：" in out
    assert "pid=111  壳" in out, "壳在跑时不能再输出整页「服务未运行」故障相"
    assert "点 × 或托盘退出" in out


def test_status_still_labels_guard_and_service(monkeypatch, capsys):
    """既有两类角色的中文标注不许被新增的 shell 分支顶掉。"""
    monkeypatch.setattr(
        app_ctl, "run_ps1",
        lambda *a: "1|keepalive|-\n2|app|pythonw.exe\n")
    monkeypatch.setattr(app_ctl, "listener_pids", lambda *a, **k: set())

    assert app_ctl.status() == 0
    out = capsys.readouterr().out
    assert "pid=1  守护" in out and "pid=2  服务" in out


def test_proc_report_knows_the_shell():
    r"""proc_report.ps1 过滤与角色：desktop\.py 进 filter，shell 角色输出。"""
    src = (TOOLS / "proc_report.ps1").read_text(encoding="utf-8")
    assert 'desktop\\.py' in src, "filter 必须含 desktop\\.py"
    assert '$role = "shell"' in src
    # ps1 带 UTF-8 BOM（PowerShell 需要），正文仍须纯 ASCII（自家规约）
    assert src.lstrip("\ufeff").isascii()


def test_installer_hooks_still_call_stop():
    """升级/卸载链靠 iss 两处钩子调 stop 生效——钩子不许被拆掉。"""
    src = (ROOT / "tools" / "installer" / "installer.iss").read_text(encoding="utf-8")
    assert 'Parameters: """tools\\app_ctl.py"" stop"' in src, "[UninstallRun] 调 stop"
    exec_call = ("Exec(ExpandConstant('{app}\\python\\pythonw.exe'), "
                 "'\"tools\\app_ctl.py\" stop'")
    assert exec_call in src, "PrepareToInstall 调 stop"
    assert "[UninstallDelete]" in src and "{app}\\python" in src


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_real_proc_report_still_parses():
    """真脚本冒烟：改动后的 ps1 在真 PowerShell 下可解析，输出行形合法。"""
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(TOOLS / "proc_report.ps1")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120)
    assert r.returncode == 0, f"proc_report.ps1 应 exit 0，实际 {r.returncode}"
    for ln in r.stdout.splitlines():
        pid, role, parent = ln.strip().split("|")
        assert pid.isdigit() and role in ("app", "keepalive", "shell")


# ------------------------------------------------- 评审③：stop 对壳先礼后兵
def test_stop_asks_the_shell_window_to_close_first(monkeypatch):
    """stop 对壳先按标题发 WM_CLOSE——壳自己的 × 语义（D-05 转换确认、
    D-01 几何保存都在那条路上）先走，窗口不在了才轮到 kill_by_cmdline
    强杀。直接 Stop-Process -Force 会把进行中的转换无确认斩断。"""
    calls = _record_ps1(monkeypatch)
    _quiet(monkeypatch)

    assert app_ctl.stop() == 0
    assert [c[0] for c in calls] == [
        "close_window_by_title.ps1",
        "kill_by_cmdline.ps1", "kill_by_cmdline.ps1", "kill_by_cmdline.ps1",
    ], "强杀之前必须先请壳自行关闭"
    assert calls[1][2] == "keepalive", "WM_CLOSE 只针对有窗口语义的壳"


def test_stop_closes_the_matching_editions_window(monkeypatch):
    """窗口标题按版本判据（is_dev_checkout 与 desktop.window_title 同源）
    精确匹配：两个版本的壳同时开在桌面上时互不误伤。"""
    calls = _record_ps1(monkeypatch)
    _quiet(monkeypatch)
    monkeypatch.setattr(app_ctl, "is_dev_checkout", lambda: True)
    app_ctl.stop()
    assert calls[0][2] == "格式工坊 · 开发版"

    calls = _record_ps1(monkeypatch)
    _quiet(monkeypatch)
    monkeypatch.setattr(app_ctl, "is_dev_checkout", lambda: False)
    app_ctl.stop()
    assert calls[0][2] == "格式工坊"


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_real_close_window_by_title_parses_and_is_ascii():
    """真脚本冒烟：查无此窗 exit 2 且零输出（同 kill_by_cmdline 的无匹配
    约定），脚本能被真 PowerShell 解析；正文必须纯 ASCII（自家规约）。"""
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(TOOLS / "close_window_by_title.ps1"),
         "-Title", "zzz_no_such_window_review", "-TimeoutSec", "1"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120)
    assert r.returncode == 2, f"查无此窗应 exit 2，实际 {r.returncode}"
    assert not r.stdout.strip(), "查无此窗不该有任何输出"
    src = (TOOLS / "close_window_by_title.ps1").read_text(encoding="utf-8")
    assert src.lstrip("\ufeff").isascii(), "ps1 正文必须纯 ASCII"


def test_status_tails_the_shell_log_too(tmp_path, monkeypatch, capsys):
    """评审⑥：[4] 最近日志要含壳的 outputs/desktop.log——壳是 status 的
    一等公民（D-04），它的故障留痕不能在状态自检里看不到。"""
    (tmp_path / "outputs").mkdir()
    (tmp_path / "outputs" / "desktop.log").write_text(
        "2026-11-05 10:00:00 install：装引擎拦截器失败\n", encoding="utf-8")
    monkeypatch.setattr(app_ctl, "ROOT", tmp_path)
    monkeypatch.setattr(app_ctl, "run_ps1", lambda *a: "")
    monkeypatch.setattr(app_ctl, "listener_pids", lambda *a, **k: set())
    monkeypatch.setattr(app_ctl, "http_ok", lambda *a, **k: None)

    assert app_ctl.status() == 0
    out = capsys.readouterr().out
    assert "desktop.log:" in out
    assert "装引擎拦截器失败" in out
