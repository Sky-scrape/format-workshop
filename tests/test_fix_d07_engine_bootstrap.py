"""D-07 回归钉：引擎引导只有一份，壳也装崩溃取证，日志按形态分流。

desktop.engine() 曾手抄 app.py main() 的前三步初始化、漏掉
install_lifecycle_hooks——faulthandler/crash.log/atexit/心跳全没装，壳进程
（pythonnet/WebView2 COM + WSGI + Office 编排线程同住）原生崩了只剩
pythonw 静默消失；同时壳经 setup_logging 写 server.log，与可选的浏览器
服务模式共写一个 RotatingFileHandler，轮转 rename 在 Windows 上互踩。

现在 app.init_app() 是唯一入口（main 与壳共用），setup_logging 按文件名
分流：server.log 归服务模式、desktop.log 归壳；reap 的 os._exit 前补退出
留痕（os._exit 跳过 atexit，钩子的正常退出留痕跑不到）。
"""

import logging
import sys
import threading
import types

import app
import desktop


def test_init_app_wires_the_four_steps(tmp_path, monkeypatch):
    """init_app 依次做：建目录、sweep、装日志、按需装生命周期钩子。"""
    calls = []
    fake_log = types.SimpleNamespace(name="fakelog")
    monkeypatch.setattr(app, "WORK", tmp_path / "work")
    monkeypatch.setattr(app, "sweep", lambda *a, **k: calls.append("sweep"))
    monkeypatch.setattr(app, "setup_logging",
                        lambda name="server.log": calls.append(f"log:{name}")
                        or fake_log)
    monkeypatch.setattr(app, "install_lifecycle_hooks",
                        lambda log: calls.append(f"hooks:{log.name}"))

    out = app.init_app()
    assert calls == ["sweep", "log:server.log", "hooks:fakelog"]
    assert out is fake_log
    assert (tmp_path / "work").is_dir()

    calls.clear()
    app.init_app(install_hooks=False)
    assert calls == ["sweep", "log:server.log"], "install_hooks=False 不装钩子"


def test_main_and_engine_share_init_app():
    """两处引导都走 init_app，不许再手抄步骤。"""
    src = (app.ROOT / "app.py").read_text(encoding="utf-8")
    assert "def init_app(" in src
    assert "log = init_app()" in src
    assert "log = setup_logging()" not in src, "main 不再自己拼四步"
    desktop_src = (app.ROOT / "desktop.py").read_text(encoding="utf-8")
    assert "paper_app.init_app(install_hooks=True" in desktop_src
    for stolen in ("paper_app.WORK.mkdir", "paper_app.sweep()",
                   "paper_app.setup_logging()"):
        assert stolen not in desktop_src, f"壳不许再手抄：{stolen}"


def test_engine_boots_with_shell_log_name(monkeypatch):
    """壳的引擎引导要传 log_name="desktop.log"，与 server.log 分流。"""
    calls = []
    fake_app = types.SimpleNamespace(
        init_app=lambda **kw: calls.append(kw) or fake_app)
    monkeypatch.setitem(sys.modules, "app", fake_app)
    monkeypatch.setattr(desktop, "_ENGINE", None)

    first = desktop.engine()
    again = desktop.engine()
    assert calls == [{"install_hooks": True, "log_name": "desktop.log"}], \
        "壳只引导一次，且日志落 desktop.log"
    assert first is fake_app and again is first


def test_setup_logging_targets_the_named_file(tmp_path, monkeypatch):
    """日志文件名可分流：desktop.log 归壳，server.log 归服务，不共写。"""
    monkeypatch.setattr(app, "ROOT", tmp_path)
    flask_app = app.app                      # 模块名 app 与 Flask 实例同名
    before_app = list(flask_app.logger.handlers)
    logging.getLogger("werkzeug").handlers.clear()

    try:
        app.setup_logging("desktop.log")
        assert (tmp_path / "outputs" / "desktop.log").is_file()
        assert not (tmp_path / "outputs" / "server.log").exists()

        app.setup_logging()
        assert (tmp_path / "outputs" / "server.log").is_file()
    finally:
        # RotatingFileHandler 占着文件句柄，拆掉并关闭，临时目录才删得动
        for h in list(flask_app.logger.handlers):
            if h not in before_app:
                flask_app.logger.removeHandler(h)
                h.close()
        for h in list(logging.getLogger("werkzeug").handlers):
            logging.getLogger("werkzeug").removeHandler(h)
            h.close()


def test_reap_logs_its_exit_before_os_exit():
    """reap 的 os._exit 跳过 atexit——退出留痕必须写在其前面。"""
    src = (app.ROOT / "desktop.py").read_text(encoding="utf-8")
    reap = src.split("def reap_when_window_gone", 1)[1].split(
        "threading.Thread(target=reap_when_window_gone", 1)[0]
    assert "log_line(" in reap, "退出留痕经 log_line 落 outputs/desktop.log"
    assert "os._exit(0)" in reap
    assert reap.index("log_line(") < reap.index("os._exit(0)"), \
        "留痕必须在 os._exit 之前，写了才来得及落盘"


def test_service_and_shell_logs_are_distinct_files():
    """两个形态各写各的日志文件（轮转 rename 竞争的前提不成立）。"""
    app_src = (app.ROOT / "app.py").read_text(encoding="utf-8")
    assert 'def setup_logging(name="server.log")' in app_src
    desktop_src = (app.ROOT / "desktop.py").read_text(encoding="utf-8")
    assert 'log_name="desktop.log"' in desktop_src
