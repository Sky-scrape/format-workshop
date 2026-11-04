"""D-06 回归钉：按端口兜底杀进程前必须核对身份。

stop() 的兜底曾是 `f"127.0.0.1:{PORT}" in line` 的无边界子串匹配——监听
127.0.0.1:50001 的行同样命中（评审修正例：127.0.0.1:15000 不含
「127.0.0.1:5000」子串，不命中），随后 taskkill /F 不核对映像名，可以
强杀监听相邻端口的无关进程；F48 引入的 -Root 圈定被这条兜底绕掉。
start() 也只做一次 TCP connect 就报「服务已在运行」，别的程序占了 5000
时用户被领去看陌生页面。现在：按列精确比对本地地址；taskkill 前用
tasklist 核对映像名为 python/pythonw，查不到映像一律不动手；start 见到
非 python 占用者据实拒绝。

评审②补的另一半：listener_pids 只认 127.0.0.1 整列时，绑 0.0.0.0/[::]
的占用者 port_alive 探得到、监听集合里却没有——start 照样报「服务已在
运行」、status 报「无进程监听」。现在 start/status 按通配口径看监听者
（wildcard=True），start 再用 app.py 的 /healthz 身份标记做最终裁决
（映像名核对区分不了「5000 上的另一个 python 应用」）；stop 的按端口
兜底保持回环口径——通配绑定者绝不可能是本软件的服务（host 写死回环），
不代杀。
"""

import types

import pytest

from conftest import TOOLS

import app_ctl

NETSTAT = (
    "  TCP    127.0.0.1:5000     0.0.0.0:0      LISTENING     4321\n"
    "  TCP    127.0.0.1:50001    0.0.0.0:0      LISTENING     999\n"
    "  TCP    127.0.0.1:15000    0.0.0.0:0      LISTENING     888\n"
    "  TCP    192.168.1.5:5000   10.0.0.2:443   ESTABLISHED   777\n"
    "  TCP    0.0.0.0:5000       0.0.0.0:0      LISTENING     666\n"
    "  UDP    127.0.0.1:5000     *:*                          555\n"
    "  Proto  本地地址          外部地址        状态           PID\n"
)


def _fake_run(monkeypatch, expect_cmd, stdout):
    monkeypatch.setattr(
        app_ctl.subprocess, "run",
        lambda cmd, **kw: (cmd[0] == expect_cmd or pytest.fail(
            f"不该调 {cmd[0]}，只许 {expect_cmd}"))
        and types.SimpleNamespace(stdout=stdout))


def test_listener_pids_parses_by_columns_not_substring(monkeypatch):
    """默认口径（stop 兜底用）：只认精确监听 127.0.0.1:5000 的行；
    相邻端口/ established/表头/UDP 都滤掉。"""
    _fake_run(monkeypatch, "netstat", NETSTAT)
    assert app_ctl.listener_pids(port=5000) == {"4321"}


def test_listener_pids_wildcard_view_sees_wide_binds(monkeypatch):
    """通配口径（start/status 用）：绑 0.0.0.0/[::] 的监听者也算数——
    它们同样回应 127.0.0.1 的连接，漏掉就是「服务已在运行」的冒名相
    （评审②）。相邻端口、UDP、established 照旧滤掉。"""
    _fake_run(monkeypatch, "netstat", NETSTAT)
    assert app_ctl.listener_pids(port=5000, wildcard=True) == {"4321", "666"}


def test_pid_is_python_requires_a_python_image(monkeypatch):
    """动杀的门槛：映像名必须是 python/pythonw；查不到映像也不动手。"""
    _fake_run(monkeypatch, "tasklist", '"python.exe","4321","Console","1"\r\n')
    assert app_ctl._image_of("4321") == "python.exe"
    assert app_ctl.pid_is_python("4321") is True

    _fake_run(monkeypatch, "tasklist", '"pythonw.exe","99","Console","1"\r\n')
    assert app_ctl.pid_is_python("99") is True

    _fake_run(monkeypatch, "tasklist", '"msedge.exe","777","Console","1"\r\n')
    assert app_ctl.pid_is_python("777") is False

    _fake_run(monkeypatch, "tasklist",
              "INFO: No tasks are running which match the specified criteria.\r\n")
    assert app_ctl._image_of("0") is None
    assert app_ctl.pid_is_python("0") is False, "查不到映像就不许杀"


def _quiet_stop(monkeypatch, pids, ours):
    """stop() 全静音：ps1 无匹配、端口恒活、监听者与身份按参数给定。"""
    monkeypatch.setattr(app_ctl, "run_ps1_full", lambda name, *a: ("", 2))
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: True)
    monkeypatch.setattr(app_ctl, "listener_pids", lambda *a, **k: set(pids))
    monkeypatch.setattr(app_ctl, "pid_is_python", lambda pid: pid in ours)
    monkeypatch.setattr(app_ctl, "time",
                        types.SimpleNamespace(sleep=lambda *_: None))
    kills = []
    monkeypatch.setattr(
        app_ctl.subprocess, "run",
        lambda cmd, **kw: kills.append(list(cmd))
        or types.SimpleNamespace(stdout=""))
    return kills


def test_stop_fallback_refuses_foreign_occupant(monkeypatch, capsys):
    """监听相邻端口的无关进程：不再被无边界子串误中，即便抓到也不代杀。"""
    kills = _quiet_stop(monkeypatch, pids=["999"], ours=set())

    assert app_ctl.stop() == 0
    assert kills == [], "非 python 占用者不许被 taskkill /F 强杀"
    out = capsys.readouterr().out
    assert "999" in out and "未代杀" in out


def test_stop_fallback_kills_only_python_occupants(monkeypatch):
    """python 形态的监听者仍按端口结束（兜底的本职），且 PID 一一对应。"""
    kills = _quiet_stop(monkeypatch, pids=["4321", "999"], ours={"4321"})

    assert app_ctl.stop() == 0
    assert [c for c in kills if c and c[0] == "taskkill"] == [
        ["taskkill", "/PID", "4321", "/F"]]


def _start_with(monkeypatch, pids, ours, healthz=True):
    """start() 静音：端口恒活、通配监听者与身份按参数给定。

    healthz 是 /healthz 的裁决结果（True/False=探到明确答案，None=探不到）。
    """
    opened = []
    monkeypatch.setattr(app_ctl, "port_alive", lambda **kw: True)
    monkeypatch.setattr(app_ctl, "listener_pids", lambda *a, **k: set(pids))
    monkeypatch.setattr(app_ctl, "pid_is_python", lambda pid: pid in ours)
    monkeypatch.setattr(app_ctl, "healthz_says_ours",
                        lambda *a, **k: healthz)
    monkeypatch.setattr(app_ctl, "open_page", lambda *a, **k: opened.append(1))
    return opened


def test_start_refuses_a_foreign_port_occupant(monkeypatch, capsys):
    """端口被非 python 程序占着：据实拒绝，不再冒充「服务已在运行」。"""
    opened = _start_with(monkeypatch, pids=["999"], ours=set(), healthz=False)

    assert app_ctl.start() == 1
    assert opened == [], "领用户去看陌生页面等于冒名，不许开页面"
    out = capsys.readouterr().out
    assert "被别的程序占用" in out
    assert "999" in out, "拒绝报话要点名占用者"
    assert "服务已在运行" not in out


def test_start_still_recognizes_our_own_service(monkeypatch, capsys):
    """自家服务（healthz 报出自家标记）在跑时行为不变：报已运行并开页面。"""
    opened = _start_with(monkeypatch, pids=["4321"], ours={"4321"},
                         healthz=True)

    assert app_ctl.start() == 0
    assert opened == [1]
    assert "服务已在运行" in capsys.readouterr().out


def test_start_refuses_a_python_impersonator_thanks_to_healthz(
        monkeypatch, capsys):
    """评审②的死角：占用者是 python 形态（映像核对放行），但 /healthz
    报不出自家标记——5000 上的另一个 python 应用，照样据实拒绝。"""
    opened = _start_with(monkeypatch, pids=["4321"], ours={"4321"},
                         healthz=False)

    assert app_ctl.start() == 1
    assert opened == []
    assert "被别的程序占用" in capsys.readouterr().out


def test_start_falls_back_to_image_check_when_healthz_cannot_answer(
        monkeypatch, capsys):
    """healthz 探不到答案（占用者不说 HTTP 等）：退回映像名核对——
    非 python 形态认得出来，python 形态按旧口径放行。"""
    opened = _start_with(monkeypatch, pids=["999"], ours=set(), healthz=None)
    assert app_ctl.start() == 1, "非 python 占用者照样拒绝"
    assert "被别的程序占用" in capsys.readouterr().out

    opened = _start_with(monkeypatch, pids=["4321"], ours={"4321"},
                         healthz=None)
    assert app_ctl.start() == 0, "回环 python 监听者按旧口径认自家"
    assert opened == [1]


def test_unbounded_substring_match_is_gone():
    """旧的无边界子串匹配不许残留（改按列比对）。"""
    src = (TOOLS / "app_ctl.py").read_text(encoding="utf-8")
    assert '" in line' not in src


def test_stop_fallback_keeps_the_loopback_view():
    """stop 的按端口兜底保持回环口径：通配绑定者绝不可能是自家服务
    （app.run host 写死 127.0.0.1），不给它被代杀的机会（评审②）。
    通配口径只属于 start/status 的「看」，不属于 stop 的「杀」。"""
    src = (TOOLS / "app_ctl.py").read_text(encoding="utf-8")
    stop_body = src.split("def stop(", 1)[1].split("\ndef ", 1)[0]
    assert "listener_pids(wildcard=True)" not in stop_body
    assert "sorted(listener_pids())" in stop_body


def test_healthz_route_carries_the_identity_marker():
    """/healthz 是 start() 身份核验的依据：应答必须带 HEALTHZ_MARK，
    版本号与 APP_VERSION 同源（D-10）。"""
    import app

    assert app_ctl.HEALTHZ_MARK == "geshi-gongfang"
    app.app.config["TESTING"] = True
    r = app.app.test_client().get("/healthz")
    assert r.status_code == 200
    body = r.get_data(as_text=True)
    assert app_ctl.HEALTHZ_MARK in body, "标记对不上，start 认不出自家服务"
    assert app.APP_VERSION in body
