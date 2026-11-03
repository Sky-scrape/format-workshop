"""审读修复回归钉：stop() 报话三分——ps1 崩溃与「无匹配」可区分
（b7 审读③，APPCTLHON 的残留同形问题）。

审读实测：powershell -File 下 ps1 抛终止错误的退出码是 1、stdout 常为空，
与 kill_by_cmdline.ps1 旧「无匹配 exit 1 且无输出」的约定同形——stop() 的
rc==1-and-empty 分支把脚本崩溃照旧报成「未找到匹配进程」（缓解因素是随后
还有按端口 taskkill 兜底，但报话不实）。进程外无从分辨两种情形，区分要从
协议本身下手，于是无匹配改走 exit 2，stop() 的四档口径变成：
    0    正常（有匹配；个别杀失败时 stdout 带 failed 行）
    2    无匹配（且无输出）
    1    脚本自身崩溃（stdout 可能为空）——报「执行失败」
    -1   powershell 缺失/超时等异常（run_ps1_full 的异常出口）——报「执行失败」
"""

import shutil
import subprocess
import types

import pytest

from conftest import ROOT, TOOLS

import app_ctl

PS1 = TOOLS / "kill_by_cmdline.ps1"


def _quiet_stop(monkeypatch, ps1_result):
    monkeypatch.setattr(app_ctl, "run_ps1_full",
                        lambda name, *args: ps1_result)
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: False)
    monkeypatch.setattr(app_ctl, "time",
                        types.SimpleNamespace(sleep=lambda *_: None))


def test_stop_ps1_crash_exit1_reported_as_failure(monkeypatch, capsys):
    """复现审读③：ps1 崩溃（exit 1、stdout 空）必须报「执行失败」，
    不得再落进「未找到匹配进程」。"""
    _quiet_stop(monkeypatch, ("", 1))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("停止脚本执行失败") == 2, "守护与服务两条都要如实报告"
    assert "未找到匹配进程" not in out, "执行失败不许说成没有进程"


def test_stop_ps1_no_match_is_exit2(monkeypatch, capsys):
    """新约定：无匹配 exit 2 且无输出 → 「未找到」；exit 1 不再是它的记号。"""
    _quiet_stop(monkeypatch, ("", 2))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("未找到匹配进程") == 2
    assert "执行失败" not in out


def test_stop_ps1_crash_with_partial_output_carries_detail(monkeypatch, capsys):
    """崩溃前已输出的残段（failed/killed 行）作为 detail 带出来。"""
    _quiet_stop(monkeypatch, ("failed 4321: access denied\n", 1))
    assert app_ctl.stop() == 0
    out = capsys.readouterr().out
    assert out.count("停止脚本执行失败") == 2
    assert "failed 4321" in out, "崩溃前残段要跟着报出来"


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_real_ps1_no_match_exits_2_silently():
    """真脚本：无匹配 exit 2 且零输出（exit 1 让位给脚本崩溃的情形）。"""
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(PS1), "-Pattern", "zzz_no_such_process_review",
         "-Root", str(ROOT)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120)
    assert r.returncode == 2, f"无匹配应 exit 2，实际 {r.returncode}"
    assert not r.stdout.strip(), "无匹配不该有任何输出"


@pytest.mark.skipif(shutil.which("powershell") is None, reason="无 powershell")
def test_real_powershell_crash_exits_1(tmp_path):
    """钉住实测前提：-File 下 ps1 抛终止错误 → 退出码 1、stdout 空。

    「崩溃与无匹配可区分」整座桥都压在这条前提上；PowerShell 某天改掉它，
    这里先红，上面的协议要跟着重审。
    """
    p = tmp_path / "crash.ps1"
    p.write_text('throw "boom"\n', encoding="ascii")
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(p)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120)
    assert r.returncode == 1, f"实测前提变了：崩溃退出码 {r.returncode}"
    assert not r.stdout.strip()
