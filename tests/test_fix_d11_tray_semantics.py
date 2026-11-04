"""D-11 回归钉：托盘驻留不许把软件藏进空气里。

tray 模式点 × 时 on_closing 直接 window.hide()，之前不校验托盘图标是否
可用；而托盘有两个不设防失败面：icon.run 在 daemon 线程里失败主流程无
感知（explorer 崩溃重启丢图标），notify_once 先把 tray-hint.flag 落盘再
notify、notify 抛错被吞——flag 已写，此后永不补发气泡。叠加净效果：窗口
隐藏、无图标、无气泡，进程还活着，用户唯一出路是任务管理器。

修法：hide 前确认 icon 存在且 icon.visible，不就绪就不 hide，弹原生
「托盘不可用，程序将退出」后放行关闭；flag 移到 notify 成功之后。

评审④补的另一半：放行关闭的两条路（图标不就绪、配置中途改回 exit）
此前 return True 前不存几何，而 tray 模式的 main() 又不注册 closing
几何回调——窗口尺寸就在这条路径上丢了。现在每条放行路径都趁窗口活着
先 save_window_state；reap 兜底也不再只挂 exit 模式（托盘驻留时绝不收）。
"""

import json
import types
from pathlib import Path

import pytest

import desktop


@pytest.fixture()
def tmp_root(tmp_path, monkeypatch):
    """托盘的 flag 与窗口几何都落在临时目录，绝不碰真实 config/。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop, "CFG", tmp_path / "config" / "window.json")
    return tmp_path


@pytest.fixture()
def tray_mode(monkeypatch):
    monkeypatch.setattr(desktop, "close_action", lambda: "tray")


def _tray(win, icon):
    t = desktop.Tray(win)
    t.icon = icon
    return t


def _icon(visible=True, notify_error=None):
    fired = []

    def notify(msg, title):
        if notify_error:
            raise notify_error
        fired.append((msg, title))
    return types.SimpleNamespace(visible=visible, notify=notify), fired


def _win():
    # width/height 必须给：放行关闭路径要趁窗口活着存几何（评审④），
    # 缺了就是「销毁后解包」形态，测不出真保存
    w = types.SimpleNamespace(width=1024, height=768, hidden=False)

    def hide():
        w.hidden = True
    w.hide = hide
    return w


def _saved_geometry(tmp_root):
    return json.loads(
        (tmp_root / "config" / "window.json").read_text(encoding="utf-8"))


def test_on_closing_hides_only_when_tray_is_ready(tmp_root, tray_mode, monkeypatch):
    icon, fired = _icon(visible=True)
    win = _win()
    t = _tray(win, icon)

    assert t.on_closing() is False, "托盘就绪：藏进去，不关窗"
    assert win.hidden is True
    assert len(fired) == 1, "首次驻留要有气泡告知"


def test_on_closing_refuses_to_hide_into_a_dead_tray(tmp_root, tray_mode,
                                                     monkeypatch):
    warned = []
    monkeypatch.setattr(desktop, "warn_tray_dead", lambda: warned.append(1))
    icon, fired = _icon(visible=False)     # icon.run 失败：visible 永远 False
    win = _win()
    t = _tray(win, icon)

    assert t.on_closing() is True, "托盘不可用：放行关闭，退出好过无窗僵尸"
    assert win.hidden is False, "不许 hide 进一个不存在的托盘"
    assert warned == [1]
    assert fired == []


def test_on_closing_without_an_icon_at_all_also_refuses(tmp_root, tray_mode,
                                                        monkeypatch):
    warned = []
    monkeypatch.setattr(desktop, "warn_tray_dead", lambda: warned.append(1))
    win = _win()
    t = desktop.Tray(win)                  # icon=None：托盘从未构造成功

    assert t.tray_ready() is False
    assert t.on_closing() is True
    assert win.hidden is False and warned == [1]


def test_on_closing_saves_geometry_before_allowing_the_close(tmp_root,
                                                             tray_mode,
                                                             monkeypatch):
    """评审④：图标不就绪放行关闭时，几何必须趁窗口活着先落盘——tray
    模式的 main() 没有注册 closing 几何回调，这条路径曾是尺寸记忆的洞。"""
    warned = []
    monkeypatch.setattr(desktop, "warn_tray_dead", lambda: warned.append(1))
    win = _win()
    t = desktop.Tray(win)                  # icon=None：托盘不可用

    assert t.on_closing() is True
    assert _saved_geometry(tmp_root) == {"w": 1024, "h": 768,
                                         "maximized": False}


def test_on_closing_still_saves_when_config_flipped_to_exit(tmp_root,
                                                            monkeypatch):
    """壳挂着 tray 处理器、配置中途改回 exit：放行关闭，几何同场保存，
    不许再像旧版那样悄悄 return None 什么都不带。"""
    monkeypatch.setattr(desktop, "close_action", lambda: "exit")
    win = _win()
    t = desktop.Tray(win)

    assert t.on_closing() is True
    assert _saved_geometry(tmp_root) == {"w": 1024, "h": 768,
                                         "maximized": False}


def test_reap_backstop_covers_tray_mode_too():
    """评审④：reap 兜底不能只挂 exit 模式——托盘模式放行关闭同样可能
    留下无窗进程。但托盘还活着且不是在退出时绝不能收：窗口是合法隐藏，
    收进程等于把驻留中的转换硬杀掉。"""
    src = (Path(__file__).resolve().parent.parent / "desktop.py").read_text(
        encoding="utf-8")
    assert 'if close_action() != "tray":\n        def reap_when_window_gone' \
        not in src, "reap 线程不许再被 exit 分支门控"
    reap = src.split("def reap_when_window_gone", 1)[1].split(
        "threading.Thread(target=reap_when_window_gone", 1)[0]
    assert "tray_ready()" in reap and "quitting" in reap, \
        "托盘驻留（icon 活着且非退出中）时 reap 不许收进程"


def test_tray_ready_is_false_when_icon_never_showed(tmp_root):
    t = desktop.Tray(_win())
    assert t.tray_ready() is False         # icon=None
    t.icon = types.SimpleNamespace()       # 连 visible 属性都没有也不炸
    assert t.tray_ready() is False


def test_notify_flag_is_written_only_after_successful_notify(tmp_root):
    """flag 写入时机是本条的另一半：notify 失败不落 flag，下次再补发。"""
    win = _win()

    icon, fired = _icon(notify_error=RuntimeError("no tray"))
    t = _tray(win, icon)
    t.notify_once()
    assert not (tmp_root / "config" / "tray-hint.flag").exists(), \
        "notify 失败就写 flag = 此后永不补发气泡"
    t.notify_once()                        # 下次再试，而不是被 flag 挡住
    assert (tmp_root / "config" / "tray-hint.flag").exists() is False

    icon, fired = _icon()
    t = _tray(win, icon)
    t.notify_once()
    assert len(fired) == 1
    assert (tmp_root / "config" / "tray-hint.flag").exists()

    icon, fired = _icon()
    t = _tray(win, icon)
    t.notify_once()
    assert fired == [], "flag 在盘上就不再骚扰"


def test_no_icon_never_writes_the_flag(tmp_root):
    """没有 icon 时连 flag 都不许写——写了就永远错过第一次告知。"""
    t = desktop.Tray(_win())
    t.notify_once()
    assert not (tmp_root / "config" / "tray-hint.flag").exists()


def test_flag_write_is_atomic(tmp_root):
    """flag 也走 atomic_io（与 D-14 同一写侧规约）。"""
    src = (Path(__file__).resolve().parent.parent / "desktop.py").read_text(
        encoding="utf-8")
    flag_fn = src.split("def notify_once", 1)[1].split("def run", 1)[0]
    assert "atomic_io.write_text" in flag_fn
    assert "flag.write_text" not in flag_fn


def test_warn_tray_dead_is_native_and_guarded():
    """告知用原生 MessageBox（失败页教训：网页弹窗盖不住原生窗口）。"""
    src = (Path(__file__).resolve().parent.parent / "desktop.py").read_text(
        encoding="utf-8")
    assert "def warn_tray_dead" in src
    assert "MessageBoxW" in src
