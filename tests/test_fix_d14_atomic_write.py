"""D-14 回归钉：小配置文件的写侧必须原子，不许生产半截 JSON。

读侧早有容错（app.py read_json_safe 原话「进程被硬杀是预期场景，半截 JSON
会留下」，desktop.py load_window_state 同款），但写侧三处都是裸 write_text：
app.py settings_shell（close.json）、desktop.py save_window_state（window.json）、
ai_assist.save_config（ai.json）。硬杀落在写入中段就是半截 JSON，用户看到的
是「设置静默回默认」。现在统一走 tools/atomic_io.write_text（tmp + os.replace），
save_window_state 失败还记 desktop.log。
"""

import json
import types
from pathlib import Path

import pytest

import atomic_io
import desktop


def test_write_text_creates_and_overwrites_atomically(tmp_path):
    """正常路径：建父目录、内容正确、目录里不留 *.tmp。"""
    p = tmp_path / "deep" / "cfg.json"

    atomic_io.write_text(p, '{"a": 1}')
    assert json.loads(p.read_text(encoding="utf-8")) == {"a": 1}

    atomic_io.write_text(p, '{"a": 2}')
    assert json.loads(p.read_text(encoding="utf-8")) == {"a": 2}
    assert not list(tmp_path.rglob("*.tmp")), "临时文件必须收走"


def test_failed_replace_keeps_the_old_file(tmp_path, monkeypatch):
    """替换失败时旧文件完好、临时文件收走——绝不留半截。"""
    p = tmp_path / "cfg.json"
    p.write_text("old-good", encoding="utf-8")

    def boom(src, dst):
        raise OSError("replace failed")
    monkeypatch.setattr(atomic_io.os, "replace", boom)

    with pytest.raises(OSError):
        atomic_io.write_text(p, "half-written")
    assert p.read_text(encoding="utf-8") == "old-good", "旧配置不许被写坏"
    assert not list(tmp_path.glob("*.tmp"))


def test_close_json_written_atomically(app_client, tmp_path, monkeypatch):
    """settings_shell 落 close.json 走原子写，不再裸 write_text。"""
    import app

    monkeypatch.setattr(app, "ROOT", tmp_path)
    r = app_client.post("/settings/shell", data={"close_action": "tray"})
    assert r.status_code == 302, "保存后照旧重定向回设置页"

    cfg = json.loads((tmp_path / "config" / "close.json").read_text(encoding="utf-8"))
    assert cfg == {"action": "tray"}
    assert not list(tmp_path.rglob("*.tmp"))

    src = (Path(app.__file__).parent / "app.py").read_text(encoding="utf-8")
    assert 'cfg_path.write_text' not in src, "settings_shell 不许再裸写"


@pytest.fixture()
def app_client():
    import app

    app.app.config["TESTING"] = True
    return app.app.test_client()


def test_window_state_written_atomically_and_failures_logged(tmp_path, monkeypatch):
    """window.json 原子写；宽高取不到/写失败都要留 desktop.log（曾静默丢）。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setattr(desktop, "CFG", tmp_path / "config" / "window.json")

    win = types.SimpleNamespace(width=1024, height=700, maximized=False)
    desktop.save_window_state(win)
    st = json.loads((tmp_path / "config" / "window.json").read_text(encoding="utf-8"))
    assert st == {"w": 1024, "h": 700, "maximized": False}
    assert not list(tmp_path.rglob("*.tmp"))

    # 窗口已销毁（closed/reap 里读 width 即此形态）：如实留痕，不写文件
    class _Dead:
        width = property(lambda self: (_ for _ in ()).throw(TypeError("gone")))
        height = 700
    (tmp_path / "outputs").mkdir(exist_ok=True)
    log_path = tmp_path / "outputs" / "desktop.log"
    before = log_path.read_text(encoding="utf-8") if log_path.exists() else ""
    desktop.save_window_state(_Dead())
    log = log_path.read_text(encoding="utf-8")
    assert "宽高取不到" in log[len(before):]

    # 落盘本身失败（磁盘/权限）：也要留痕
    import atomic_io

    def boom(path, text):
        raise OSError("disk full")
    monkeypatch.setattr(atomic_io, "write_text", boom)
    desktop.save_window_state(win)
    assert "save_window_state failed" in log_path.read_text(encoding="utf-8")


def test_ai_config_written_atomically(tmp_path, monkeypatch):
    """ai.json 原子写；DEFAULTS 过滤行为保持不变。"""
    import ai_assist

    cfg_path = tmp_path / "config" / "ai.json"
    monkeypatch.setattr(ai_assist, "CONFIG_PATH", cfg_path)

    raw = ai_assist.save_config({
        "base_url": "https://api.example.com/v1", "model": "demo",
        "evil_key": "x",               # 不在 DEFAULTS 里，必须被滤掉
        "api_key": None,               # None 不覆盖已有值
    })
    assert "evil_key" not in raw
    on_disk = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert on_disk == raw
    assert not list(tmp_path.rglob("*.tmp"))


def test_concurrent_writers_never_share_a_tmp_name(tmp_path, monkeypatch):
    """评审⑤：同进程两个写者并发写同一目标，临时名必须各不相同。

    临时名只带 pid 时两个写者共用同一个 tmp：先落位的 os.replace 把它挪走，
    另一个写者对着已不存在的路径替换，当场 FileNotFoundError——app.py 的
    write_json_atomic 为同类并发把临时名改成随机后缀在先，本 helper 不能
    与仓库既有结论相悖。
    """
    import threading

    p = tmp_path / "cfg.json"
    seen = []
    real_replace = atomic_io.os.replace

    def spy_replace(src, dst):
        seen.append(Path(src).name)
        return real_replace(src, dst)
    monkeypatch.setattr(atomic_io.os, "replace", spy_replace)
    atomic_io.write_text(p, "first")
    atomic_io.write_text(p, "second")
    assert len(set(seen)) == 2, "两个写者的临时名不许相同（评审⑤）"

    # 再来真并发的：任何写者都不许冒错误，最终内容必是某个完整 payload
    monkeypatch.setattr(atomic_io.os, "replace", real_replace)
    errors = []

    def work(i):
        try:
            atomic_io.write_text(p, f"payload-{i}-" + "字" * 2000)
        except Exception as e:            # noqa: BLE001
            errors.append(e)

    threads = [threading.Thread(target=work, args=(i,)) for i in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert errors == [], f"并发写不许互相踩：{errors[:2]}"
    content = p.read_text(encoding="utf-8")
    assert content.startswith("payload-") and content.endswith("字" * 2000), \
        "落盘内容必须是某个写者的完整 payload，不许是半截"
    assert not list(tmp_path.glob("*.tmp")), "临时文件必须收走"


def test_all_three_writers_route_through_the_helper():
    """三处写侧不许再出现裸 write_text 覆盖配置文件。"""
    root = Path(__file__).resolve().parent.parent
    app_src = (root / "app.py").read_text(encoding="utf-8")
    desktop_src = (root / "desktop.py").read_text(encoding="utf-8")
    ai_src = (root / "tools" / "ai_assist.py").read_text(encoding="utf-8")
    assert 'atomic_io.write_text' in app_src
    assert 'atomic_io.write_text' in desktop_src
    assert 'atomic_io.write_text' in ai_src
    assert "CFG.write_text" not in desktop_src
    assert "CONFIG_PATH.write_text" not in ai_src
