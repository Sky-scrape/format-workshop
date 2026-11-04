"""D-02 回归钉：install_engine 重试幂等，失败各路都有日志可查。

①幂等：Invoke 是投递到 UI 线程的异步动作，重试循环每 0.2s 重发 install()；
  UI 线程繁忙时第二个 install() 会在第一个生效前排入队列，不查「已装过」
  就无条件 AddWebResourceRequestedFilter + WebResourceRequested += 会装出
  双处理器，站内请求被处理两遍（:293-295 的防重入只挡跨调用）。
②可诊断：state["error"] 采集后无消费者、work() 的 except 只补
  deferral.Complete——用户看到 WebView2 原生错误页而 outputs/desktop.log
  零痕迹，FAIL 页却让用户去查这个文件。现在统一落 log_exc/log_line。

测试用假 WebView2（Pythonnet 的 System 在测试解释器里不存在）。
"""

import sys
import threading
import types
import time

import desktop


class _Event:
    """最小事件持有者：支持 +=，收到的处理器可取出来手动触发。"""

    def __init__(self):
        self.handlers = []

    def __iadd__(self, handler):
        self.handlers.append(handler)
        return self


class _Core:
    def __init__(self, fail_filter=False, core_present=True):
        self.filters = []
        self.WebResourceRequested = _Event()
        self.fail_filter = fail_filter
        self.core_present = core_present

    def AddWebResourceRequestedFilter(self, pattern, ctx):
        if self.fail_filter:
            raise RuntimeError("filter boom")
        self.filters.append(pattern)


def _fake_net(monkeypatch):
    """pythonnet 不在测试解释器里：System/Microsoft 全用替身。"""
    sysm = types.ModuleType("System")
    sysm.Action = lambda f: f                     # Action(fn) 恒等
    sysm.IO = types.SimpleNamespace(MemoryStream=lambda *a: None)
    monkeypatch.setitem(sys.modules, "System", sysm)
    ctx = types.ModuleType("Microsoft.Web.WebView2.Core")
    ctx.CoreWebView2WebResourceContext = types.SimpleNamespace(All="all")
    monkeypatch.setitem(sys.modules, "Microsoft.Web.WebView2.Core", ctx)


def _window(core, defer=False):
    native = types.SimpleNamespace(webview=types.SimpleNamespace(CoreWebView2=core))
    if defer:
        native.queue = []

        def invoke(action):
            native.queue.append(action)
        native.Invoke = invoke
    else:
        native.Invoke = lambda action: action()
    return types.SimpleNamespace(uid=f"d02-{id(core)}", native=native)


def _poll(fn, timeout=3.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if fn():
            return True
        time.sleep(0.05)
    return False


def test_install_is_idempotent_for_queued_retries(monkeypatch):
    """UI 线程繁忙时排队的重复 install() 不许再装一遍（双处理器根因）。"""
    _fake_net(monkeypatch)
    monkeypatch.setattr(desktop, "_ENGINE_INSTALLED", set())
    core = _Core()
    win = _window(core, defer=True)               # Invoke 只排队，不执行

    assert desktop.install_engine(win, timeout=0.5) is False  # 没人执行它

    for action in list(win.native.queue):         # UI 线程腾出手，逐个执行
        action()

    assert core.filters == [f"{desktop.LOCAL_URL}/*"], "filter 只装一次"
    assert len(core.WebResourceRequested.handlers) == 1, \
        "排队的重试不许叠加第二个请求处理器"


def test_successful_install_marks_the_window(monkeypatch):
    _fake_net(monkeypatch)
    monkeypatch.setattr(desktop, "_ENGINE_INSTALLED", set())
    core = _Core()
    win = _window(core)

    assert desktop.install_engine(win, timeout=2) is True
    assert win.uid in desktop._ENGINE_INSTALLED


def test_install_failure_is_logged(tmp_path, monkeypatch):
    """装拦截器抛错：desktop.log 留 traceback，放弃时带 error。"""
    _fake_net(monkeypatch)
    monkeypatch.setattr(desktop, "_ENGINE_INSTALLED", set())
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    core = _Core(fail_filter=True)

    assert desktop.install_engine(_window(core), timeout=0.3) is False

    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "装引擎拦截器失败" in log and "filter boom" in log
    assert "install_engine 放弃" in log and "RuntimeError" in log


def test_give_up_without_core_is_logged(tmp_path, monkeypatch):
    """WebView2 一直没就绪（CoreWebView2=None）：放弃也要落一行，不是无声超时。"""
    _fake_net(monkeypatch)
    monkeypatch.setattr(desktop, "_ENGINE_INSTALLED", set())
    monkeypatch.setattr(desktop, "ROOT", tmp_path)

    assert desktop.install_engine(_window(None), timeout=0.3) is False

    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "重试期内 WebView2 未就绪" in log


def test_work_failure_leaves_uri_and_traceback(tmp_path, monkeypatch):
    """请求处理异常：deferral 照常 Complete，desktop.log 带 URI 与栈。"""
    _fake_net(monkeypatch)
    monkeypatch.setattr(desktop, "_ENGINE_INSTALLED", set())
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    core = _Core()
    win = _window(core)
    assert desktop.install_engine(win, timeout=2) is True

    handler = core.WebResourceRequested.handlers[0]
    monkeypatch.setattr(desktop, "wsgi_call",
                        lambda *a, **k: (_ for _ in ()).throw(
                            RuntimeError("wsgi boom")))

    completed = []

    class _Deferral:
        def Complete(self):
            completed.append(1)

    class _Header:
        Key, Value = "Host", "127.0.0.1"

    class _Headers:
        def GetEnumerator(self):
            return iter([_Header()])

    args = types.SimpleNamespace(
        Request=types.SimpleNamespace(
            Uri="http://127.0.0.1/convert", Method="POST",
            Headers=_Headers(), Content=None),
        GetDeferral=lambda: _Deferral())
    handler(None, args)

    assert _poll(lambda: completed == [1]), "deferral 必须 Complete（防挂死）"
    assert _poll(lambda: "请求处理失败" in
                 (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8"))
    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "http://127.0.0.1/convert" in log, "定位是哪个请求失败"
    assert "wsgi boom" in log, "栈要跟着落盘"


def test_load_app_url_failure_lands_in_the_log(tmp_path, monkeypatch):
    """引擎装载失败：desktop.log 有 traceback，FAIL 页指的文件有内容。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop, "engine",
                        lambda: (_ for _ in ()).throw(RuntimeError("no engine")))
    shown = []
    win = types.SimpleNamespace(load_url=lambda u: shown.append(u),
                                load_html=lambda h: shown.append(h))

    desktop.load_app_url(win, f"{desktop.LOCAL_URL}/")
    assert _poll(lambda: (tmp_path / "outputs" / "desktop.log").exists())
    assert _poll(lambda: "no engine" in
                 (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8"))
    assert shown == [desktop.FAIL]


def test_wait_for_native_timeout_is_logged(tmp_path, monkeypatch):
    """native 一直不出现（WebView2 未创建）：放弃落一行。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    win = types.SimpleNamespace(uid="d02-nonative", native=None)

    assert desktop.install_engine(win, timeout=0.2) is False
    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "等不到 WebView2 native" in log


def test_log_helpers_shapes(tmp_path, monkeypatch):
    """log_line 带时间戳一行；log_exc 附 traceback；写不进不抛。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    desktop.log_line("hello")
    try:
        raise ValueError("bad thing")
    except ValueError:
        desktop.log_exc("上下文", "extra")

    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "[20" in log and "hello" in log
    assert "上下文：extra" in log and "ValueError" in log and "bad thing" in log

    monkeypatch.setattr(desktop, "ROOT", tmp_path / "blocked")
    monkeypatch.setattr(desktop.Path, "mkdir",
                        lambda *a, **k: (_ for _ in ()).throw(OSError("no")))
    try:
        raise RuntimeError("x")
    except RuntimeError:
        desktop.log_exc("写不进也不许抛")   # 不抛即通过
