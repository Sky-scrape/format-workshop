"""桌面壳的回归钉：进程内引擎（无端口、无本地服务）+ 站内新窗口。

两条用户可见的约定：
1. 壳里点「格式规格 ↗」开的是壳自己的窗口，不是带 127.0.0.1:端口 的浏览器；
2. 页面由进程内 Flask 通过 WebView2 请求拦截直接应答，不监听任何端口，
   也没有 keepalive / app.py 这样的后台服务进程。
"""

import json
import re
import sys
import types
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import desktop  # noqa: E402


@pytest.mark.parametrize("path", [
    "/journal/cjyj/spec",
    "/job/abc/triage.html",
    "/job/abc/audit.html",
])
def test_child_url_accepts_site_paths(path):
    assert desktop.LOCAL_URL == "http://127.0.0.1"
    assert desktop.child_url(path) == f"http://127.0.0.1{path}"


@pytest.mark.parametrize("bad", [
    "http://example.com/x",        # 外站
    "https://127.0.0.1/journal",   # 绝对地址由壳自己拼，不接受外部传入
    "//example.com/x",             # 协议相对，浏览器会当外站
    "javascript:alert(1)",
    "journal/cjyj/spec",           # 相对路径没有前导斜杠
    "",
    None,
])
def test_child_url_rejects_external_and_malformed(bad):
    assert desktop.child_url(bad) is None


def test_child_title_never_collides_with_main_window():
    """主窗收尾靠 FindWindowW 精确匹配标题；子窗同名会让进程退不掉。"""
    for path in ("/journal/cjyj/spec", "/job/x/triage.html",
                 "/job/x/audit.html", "/other"):
        title = desktop.child_title(path)
        assert title.startswith("格式工坊 · ")
        assert title != desktop.window_title()


def test_shell_bridge_rejects_external_without_opening_a_window():
    """外站地址在 import webview 之前就该被挡下，返回错误而不是抛异常。"""
    r = desktop.ShellBridge().open_in_app("http://example.com")
    assert r["ok"] is False


def test_statusbar_does_not_show_the_port():
    """状态栏是软件的身份证，不是服务的监控面板：版本 + 状态 + 数据边界。

    用户看到 127.0.0.1:端口 会以为软件在用浏览器/网页形态打开；端口与
    「服务」字样都不该出现在界面上。
    """
    html = (ROOT / "templates" / "base.html").read_text(encoding="utf-8")
    footer = html.split('<footer class="statusbar"', 1)[1].split("</footer>", 1)[0]
    assert "127.0.0.1" not in footer
    assert "服务运行中" not in footer
    assert "sbPort" not in html
    assert "qport" not in html


def test_statusbar_separators_use_sb_sep_not_middle_dot():
    """分隔符规约禁用「·」中点（各字体基线不一），页内元信息一律细竖线。

    状态栏右侧数据边界声明曾与左侧并用两种分隔符；窗口/标签页标题里的
    「·」是桌面惯例，不在规约之列，只约束 statusbar 页内部分。
    """
    html = (ROOT / "templates" / "base.html").read_text(encoding="utf-8")
    footer = html.split('<footer class="statusbar"', 1)[1].split("</footer>", 1)[0]
    assert "·" not in footer
    # 右侧数据边界与左侧同行同款：本地运行 | 数据不出本机
    assert '<span>本地运行</span><i class="sb-sep"></i><span>数据不出本机</span>' in footer


def test_shell_runs_flask_in_process_without_a_server():
    """桌面壳不再起 keepalive/app.py：请求由 WebView2 拦截后交给 WSGI。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    assert "WebResourceRequested" in src
    assert "wsgi_call" in src
    assert "HTTPConnection" not in src
    assert "spawn_keepalive" not in src
    assert "PAPER_PORT" not in src
    assert "/?from=" not in src


class _DummyEngine:
    """替身引擎：只验证 wsgi_call 组装的 WSGI environ 与响应解析。"""

    @staticmethod
    def app(environ, start_response):
        assert environ["REQUEST_METHOD"] == "POST"
        assert environ["PATH_INFO"] == "/echo"
        assert environ["QUERY_STRING"] == "q=1"
        assert environ["CONTENT_TYPE"] == "text/plain"
        assert environ["CONTENT_LENGTH"] == "3"
        assert environ["HTTP_HOST"] == "127.0.0.1"
        assert environ["REMOTE_ADDR"] == "127.0.0.1"
        body = environ["wsgi.input"].read()
        start_response("201 Created",
                       [("Content-Type", "text/plain"), ("X-Test", "1")])
        return [b"echo:", body]


def test_wsgi_call_builds_environ_and_parses_response(monkeypatch):
    monkeypatch.setattr(desktop, "engine", lambda: _DummyEngine)
    code, reason, headers, payload = desktop.wsgi_call(
        "POST", "http://127.0.0.1/echo?q=1",
        {"Host": "127.0.0.1", "Content-Type": "text/plain"}, b"abc")
    assert (code, reason) == (201, "Created")
    assert payload == b"echo:abc"
    assert ("X-Test", "1") in headers


def test_response_header_text_recomputes_length():
    text = desktop.response_header_text(
        [("Content-Type", "text/plain"), ("Content-Length", "999")], 3, "GET")
    assert "Content-Type: text/plain" in text
    assert "Content-Length: 3" in text
    assert "Content-Length: 999" not in text
    # HEAD 没 body，但应保留 Flask 给出的真实长度
    head = desktop.response_header_text([("Content-Length", "999")], 0, "HEAD")
    assert "Content-Length: 999" in head


def test_save_bridge_fetches_through_in_process_engine(monkeypatch):
    """另存为不再回环 HTTP：文件名解析沿用 attachment() 的 RFC 5987 写法。"""
    seen = {}

    def fake_wsgi(method, uri, headers=None, body=b""):
        seen.update(method=method, uri=uri)
        return (200, "OK",
                [("Content-Disposition",
                  "attachment; filename*=UTF-8''%E7%A8%BF.docx")],
                b"docx-bytes")

    monkeypatch.setattr(desktop, "wsgi_call", fake_wsgi)
    data, name = desktop.ShellBridge()._fetch("/job/x/download/docx")
    assert data == b"docx-bytes"
    assert name == "稿.docx"
    assert seen == {"method": "GET", "uri": "http://127.0.0.1/job/x/download/docx"}


def test_save_bridge_reports_non_200(monkeypatch):
    monkeypatch.setattr(desktop, "wsgi_call",
                        lambda *a, **k: (404, "Not Found", [], b""))
    with pytest.raises(RuntimeError):
        desktop.ShellBridge()._fetch("/missing")


# ---------------------------------------------------------------- D-01 几何记忆
def test_geometry_saved_on_closing_not_closed():
    """几何保存必须挂在 closing（窗口存活期）；closed/reap 只是兜底。

    pywebview 6.2.1 的 winforms on_close 先 del BrowserView.instances[uid]
    再 set closed——closed/reap 里读 window.width 是对已销毁实例实时解包
    get_size 的 None，必抛 TypeError 且被裸 except 吞掉：尺寸记忆因此
    从未落盘过（D-01）。
    """
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    assert "win.events.closed += lambda: save_window_state" not in src
    main_body = src.split("def main(", 1)[1]
    exit_branch = main_body.split('if close_action() != "tray":', 1)[1].split(
        "def boot(", 1)[0]
    assert "win.events.closing += confirm_exit_on_close" in exit_branch
    assert "win.events.closing += lambda: save_window_state(win)" in exit_branch
    reap = main_body.split("def reap_when_window_gone", 1)[1].split(
        "threading.Thread(target=reap_when_window_gone", 1)[0]
    assert "if not _GEOMETRY_SAVED:" in reap, "reap 只留兜底，不重复保存"


def test_save_uses_live_window_and_real_maximized(tmp_path, monkeypatch):
    """活窗：宽高取得到；maximized 用 IsZoomed 实测（构造参数只会记启动值）。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop, "CFG", tmp_path / "config" / "window.json")
    win = types.SimpleNamespace(width=1024, height=768)
    win.native = types.SimpleNamespace(
        Handle=types.SimpleNamespace(ToInt64=lambda: 4321))
    fake_user32 = types.SimpleNamespace(IsZoomed=lambda hwnd: 1)
    fake_ctypes = types.SimpleNamespace(
        windll=types.SimpleNamespace(user32=fake_user32))
    monkeypatch.setitem(sys.modules, "ctypes", fake_ctypes)

    desktop.save_window_state(win)

    st = json.loads((tmp_path / "config" / "window.json").read_text(encoding="utf-8"))
    assert st == {"w": 1024, "h": 768, "maximized": True}


def test_save_after_destroy_is_a_logged_no_op(tmp_path, monkeypatch):
    """销毁后的窗（closed/reap 形态）：不写文件，desktop.log 留痕不静默。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop, "CFG", tmp_path / "config" / "window.json")

    class _Dead:
        width = property(lambda self: (_ for _ in ()).throw(TypeError("gone")))
        height = 768

    desktop.save_window_state(_Dead())

    assert not (tmp_path / "config" / "window.json").exists()
    log = (tmp_path / "outputs" / "desktop.log").read_text(encoding="utf-8")
    assert "宽高取不到" in log


def test_tray_saves_geometry_while_window_alive():
    """托盘驻留路径：hide 前保存（hide 后窗口就是销毁后解包形态）。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    on_closing = src.split("def on_closing", 1)[1].split("def tray_ready", 1)[0]
    assert "save_window_state(self.window)" in on_closing
    assert on_closing.index("save_window_state(self.window)") < \
        on_closing.index("self.window.hide()")


def test_reap_and_bottom_handlers_go_through_the_shared_log():
    """壳侧落盘统一走 log_line/log_exc（D-02）：裸 append 只许在 log_line 一处。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    occurrences = [m.start() for m in re.finditer(
        re.escape('open("a", encoding="utf-8")'), src)]
    assert len(occurrences) == 1, "落盘只许经 log_line 一处，不许各写一遍"
    log_line_body = src.split("def log_line", 1)[1].split("def log_exc", 1)[0]
    assert 'open("a", encoding="utf-8")' in log_line_body
    assert "def log_line" in src and "def log_exc" in src
