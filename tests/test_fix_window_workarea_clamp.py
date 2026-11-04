"""默认/记忆窗口尺寸的工作区钳制：窗口不许比屏幕大。

复测实测：1280x800 屏上默认 1280x860 首次启动底边被任务栏裁掉；来自更大
屏幕的记忆尺寸（拔副屏、换分辨率）同样越界，壳内子窗（规格/详情页）也有
同样问题。约定：create_window 的宽高必须先过 fit_to_work_area，按主屏
工作区（不含任务栏）钳制，留级联偏移余量。
"""

import sys
import types
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import desktop  # noqa: E402


class _FakeRect:
    def __init__(self):
        self.left = self.top = self.right = self.bottom = 0


def _fake_ctypes(monkeypatch, work=None, ok=True):
    """装一个最小 ctypes 替身：SystemParametersInfoW 把工作区写进 byref 的 RECT。"""
    calls = []

    def spi(msg, param, rect, update):
        calls.append(msg)
        if not ok:
            return 0
        rect.right, rect.bottom = work
        return 1

    fake = types.ModuleType("ctypes")
    fake.byref = lambda x: x
    fake.windll = types.SimpleNamespace(
        user32=types.SimpleNamespace(SystemParametersInfoW=spi))
    fake_wt = types.ModuleType("wintypes")
    fake_wt.RECT = _FakeRect
    fake.wintypes = fake_wt
    monkeypatch.setitem(sys.modules, "ctypes", fake)
    monkeypatch.setitem(sys.modules, "wintypes", fake_wt)
    return calls


def test_default_size_clamped_on_short_screen(monkeypatch):
    """默认 1280x860 在 1280x800 屏（工作区 1280x760）上必须被钳进工作区。"""
    _fake_ctypes(monkeypatch, work=(1280, 760))
    assert desktop.fit_to_work_area(1280, 860) == (1280 - 32, 760 - 32)


def test_big_screen_memory_clamped_on_small_screen(monkeypatch):
    """拔掉副屏后，来自 1920x1080 的记忆尺寸也不许把窗口撑出小屏。"""
    _fake_ctypes(monkeypatch, work=(1280, 760))
    assert desktop.fit_to_work_area(1920, 1000) == (1280 - 32, 760 - 32)


def test_normal_size_untouched_on_big_screen(monkeypatch):
    """放得下的尺寸一字不动——钳制只收越界的，不改别人的正常记忆。"""
    _fake_ctypes(monkeypatch, work=(1920, 1040))
    assert desktop.fit_to_work_area(1280, 860) == (1280, 860)


def test_spi_failure_passes_through(monkeypatch):
    """SPI 失败按原值放行：宁可维持旧行为，也不拿假尺寸瞎猜。"""
    _fake_ctypes(monkeypatch, work=(1280, 760), ok=False)
    assert desktop.fit_to_work_area(1280, 860) == (1280, 860)


def test_degenerate_work_area_stays_positive(monkeypatch):
    """极端小的工作区也不许算出非正数尺寸（min_size 托底是 WinForms 的事）。"""
    _fake_ctypes(monkeypatch, work=(100, 80))
    w, h = desktop.fit_to_work_area(1280, 860)
    assert w > 0 and h > 0


def test_non_windows_passes_through(monkeypatch):
    _fake_ctypes(monkeypatch, work=(1280, 760))
    monkeypatch.setattr(desktop.os, "name", "posix")
    assert desktop.fit_to_work_area(1280, 860) == (1280, 860)


def test_real_win32_path_smoke():
    """真 Windows 冒烟（不走替身）：超大尺寸必须被真实工作区钳住。"""
    if sys.platform != "win32":
        pytest.skip("仅 Windows")
    w, h = desktop.fit_to_work_area(12800, 9000)
    assert 0 < w < 12800 and 0 < h < 9000


def test_main_pins_the_clamp():
    """源码钉：主窗与子窗的宽高都必须来自 fit_to_work_area，不许被绕开。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    assert "fit_to_work_area(st.get(\"w\", 1280)" in src
    assert "width=w0, height=h0,\n        min_size=(980, 640)" in src
    assert "fit_to_work_area(1000, 780)" in src
