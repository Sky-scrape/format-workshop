"""D-05 回归钉：关窗语义收敛成一套，默认路径有转换防护。

曾经的矛盾：docstring 承诺 exit 下点 × 弹三选（desktop.py close_action /
ask_on_close），但 ask_on_close 只被 on_closing 调、on_closing 只在 tray
模式注册——默认 exit 模式 closing 无订阅者，三选框是不可达死代码；reap
只盯窗口标题不看引擎忙闲，os._exit(0) 把整进程带走，/convert 的
stage_parse 同步跑在请求线程（壳里即 WSGI 桥 work 线程），关窗=转换线程
被硬杀，等几十秒的成果只剩半截 job 目录。

现在：exit 模式在 main() 直接注册 closing（confirm_exit_on_close）——
引擎忙（任务锁被持有）弹原生「退出将中断转换」确认，闲时即关即退；
reap 保留作兜底（closing 订阅有实测的不触发前科）；死代码删除，
docstring 改到与行为一致。转换子进程挂 Job Object 属后续加固，不在本条。
"""

import sys
import threading
import types
from pathlib import Path

import app
import desktop

ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------- 忙闲信号
def test_conversion_running_follows_the_job_lock():
    """锁被持有 = 转换在跑；锁残留但没人持不算忙。"""
    lock = app.job_lock("d05job")
    with lock:
        assert app.conversion_running() is True
    assert app.conversion_running() is False, "上一轮结束后残留的锁不算忙"


def test_conversion_running_counts_journal_locks_too():
    """journal:{jid} 的锁也计入——新建刊的 AI 抽取同样是在途工作。"""
    lock = app.job_lock("journal:d05j")
    with lock:
        assert app.conversion_running() is True
    assert app.conversion_running() is False


# ---------------------------------------------------------------- 壳侧防护
def test_exit_mode_subscribes_closing(monkeypatch):
    """exit 模式必须注册 closing 处理（曾经默认路径无订阅者）。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    main_body = src.split("def main(", 1)[1]
    guard = main_body.split("if close_action() != \"tray\":", 1)[1].split(
        "def boot(", 1)[0]
    assert "win.events.closing += confirm_exit_on_close" in guard
    assert "reap" in main_body, "reap 线程保留作兜底（closing 有不触发前科）"


def test_confirm_exit_asks_only_when_busy(monkeypatch):
    confirmed = []
    monkeypatch.setattr(desktop, "conversion_busy", lambda: True)
    monkeypatch.setattr(desktop, "_confirm_interrupt_box",
                        lambda: confirmed.append(1) or False)
    assert desktop.confirm_exit_on_close() is False, "用户取消 → 拦下本次关闭"
    assert confirmed == [1]

    monkeypatch.setattr(desktop, "conversion_busy", lambda: False)
    monkeypatch.setattr(desktop, "_confirm_interrupt_box",
                        lambda: (_ for _ in ()).throw(AssertionError("不该弹框")))
    assert desktop.confirm_exit_on_close() is True, "闲时即关即退，不弹框"


def test_conversion_busy_reflects_the_engine(monkeypatch):
    """引擎没起来（启动页阶段）不算忙；引擎报忙才算忙；引擎异常按闲兜底。"""
    monkeypatch.setattr(desktop, "_ENGINE", None)
    assert desktop.conversion_busy() is False

    monkeypatch.setattr(desktop, "_ENGINE",
                        types.SimpleNamespace(conversion_running=lambda: True))
    assert desktop.conversion_busy() is True

    def boom():
        raise RuntimeError("engine half-loaded")
    monkeypatch.setattr(desktop, "_ENGINE", types.SimpleNamespace(
        conversion_running=boom))
    assert desktop.conversion_busy() is False, "取不到忙闲时按闲处理，别拦关窗"


def test_confirm_box_uses_native_okcancel(monkeypatch):
    """确认框是 Win32 的（网页弹窗盖不住标题栏），OK 才放行。"""
    calls = {}

    class _User32:
        @staticmethod
        def MessageBoxW(hwnd, text, title, flags):
            calls.update(text=text, title=title, flags=flags)
            return 2                       # IDCANCEL
    fake_ctypes = types.SimpleNamespace(
        windll=types.SimpleNamespace(user32=_User32))
    monkeypatch.setitem(sys.modules, "ctypes", fake_ctypes)
    monkeypatch.setattr(desktop, "window_title", lambda: "格式工坊")

    assert desktop._confirm_interrupt_box() is False
    assert "中断转换" in calls["text"] and "半截" in calls["text"]
    assert calls["title"] == "格式工坊"
    # MB_OKCANCEL=0x1 | MB_ICONWARNING=0x30 | MB_TOPMOST | MB_SETFOREGROUND
    assert calls["flags"] & 0x1 and calls["flags"] & 0x30

    _User32.MessageBoxW = staticmethod(lambda *a: 1)   # IDOK
    assert desktop._confirm_interrupt_box() is True


# ---------------------------------------------------------------- 死代码清理
def test_three_choice_dead_code_is_gone():
    """不可达的三选框随旧语义一并删除（docstring 里的历史说明允许保留）。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    for gone in ("def ask_on_close", "IDC_MINIMIZE", "IDC_CLOSE",
                 "MB_YESNOCANCEL", "window.minimize()"):
        assert gone not in src, f"旧语义代码残留：{gone}"
    assert "confirm_exit_on_close" in src


def test_copy_matches_the_converged_semantics():
    """设置页与等待遮罩的文案与新语义一致（转换中会确认）。"""
    settings = (ROOT / "templates" / "settings.html").read_text(encoding="utf-8")
    assert "转换进行中会先确认" in settings
    workspace = (ROOT / "templates" / "workspace.html").read_text(encoding="utf-8")
    assert workspace.count("误关窗口会先弹确认") == 3, \
        "静态遮罩与两类等待文案（docx/.doc）都要说"
