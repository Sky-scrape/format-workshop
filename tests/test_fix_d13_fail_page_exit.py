"""D-13 回归钉：启动失败页不许是死胡同。

引擎装载失败落纯静态 FAIL 页，文案「请查看 outputs\\desktop.log 与
outputs\\server.log 后重试」三个问题：①「后重试」没有任何对应动作；
②对便携包目标用户是开发者语言，没有按钮能打开 outputs 目录；③server.log
由 setup_logging 创建，而它只在 import app 成功后才被调用——恰在 import
阶段失败时 server.log 不存在，按文案去找会扑空。现在 FAIL 页有「重试」
（js_api retry 重跑 load_app_url）与「打开日志文件夹」（js_api
open_outputs 用资源管理器定位 outputs）两个动作，文案不再指向具体日志文件。
"""

import types

import desktop


def test_fail_page_has_actions_and_no_dead_copy():
    """失败页有重试与打开日志文件夹两个动作；不再指向可能不存在的 server.log。"""
    html = desktop.FAIL
    assert "重试" in html and "打开日志文件夹" in html
    assert "pywebview" in html and "api('retry')" in html
    assert "api('open_outputs')" in html
    assert "outputs" in html, "仍要告诉用户日志在 outputs 文件夹"
    assert "server.log" not in html, "import 阶段失败时 server.log 不存在，别指路"
    assert "desktop.log" not in html, "不再让无 Python 背景的用户手找日志文件"


def test_retry_reloads_the_app_url(monkeypatch):
    """重试 = 重跑 load_app_url（engine 失败不缓存，import 成功即换页）。"""
    seen = []
    monkeypatch.setattr(desktop, "load_app_url",
                        lambda win, url: seen.append((win, url)))
    bridge = desktop.ShellBridge()
    win = object()
    bridge._window = win

    assert bridge.retry() == {"ok": True}
    assert seen == [(win, "http://127.0.0.1/")]


def test_retry_without_a_window_reports_instead_of_raising(monkeypatch):
    monkeypatch.setattr(desktop.ShellBridge, "_owner", lambda self: None)
    assert desktop.ShellBridge().retry() == {"ok": False, "message": "窗口不存在"}


def test_open_outputs_creates_and_opens_the_folder(tmp_path, monkeypatch):
    """打开日志文件夹：outputs 不存在就建，然后交给资源管理器。"""
    opened = []
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop.os, "name", "nt")
    monkeypatch.setattr(desktop.os, "startfile",
                        lambda p: opened.append(p), raising=False)

    r = desktop.ShellBridge().open_outputs()

    assert r == {"ok": True}
    assert (tmp_path / "outputs").is_dir()
    assert opened == [tmp_path / "outputs"]


def test_open_outputs_reports_os_errors_without_raising(tmp_path, monkeypatch):
    """explorer 打不开也要给页面一句人话，不许把异常抛进 JS 里变 undefined。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop.os, "name", "nt")

    def boom(p):
        raise OSError("no explorer")
    monkeypatch.setattr(desktop.os, "startfile", boom, raising=False)

    r = desktop.ShellBridge().open_outputs()
    assert r["ok"] is False and "no explorer" in r["message"]


def test_retry_is_exposed_on_the_js_bridge():
    """FAIL 页由 js_api=ShellBridge 承载——两个新动作必须挂在桥上。"""
    for name in ("retry", "open_outputs"):
        assert callable(getattr(desktop.ShellBridge, name, None)), f"缺 {name}"
    # fail 页两处动作与桥方法同名，改名会让按钮静默失效
    assert "api('retry')" in desktop.FAIL and "api('open_outputs')" in desktop.FAIL
