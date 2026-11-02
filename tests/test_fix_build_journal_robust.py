"""FDOC2 回归钉子：build_journal 清理旧 PDF 时文件被占用的容错。

批四验证的副作用记录：Windows 上旧 PDF 被阅读器占用时，pdf.unlink() 抛出的
PermissionError 没有任何处理，用户看到的是裸 traceback——不静默，但报错
不人话化。修复后两处清理点（convert_pdf 转换前、main 无转换器时）都走同一个
人话出口：点名文件、给出「关闭阅读器后重试」的出路、退出码 1。

夹具姿势沿用 test_fix_toolchain_isolation.py 的 build_journal 用例：进程内
import 工具模块 + monkeypatch。真被占用没法在 CI 里稳定复现，
PermissionError 用只对目标文件生效的假 unlink 模拟。
"""

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
sys.path.insert(0, str(TOOLS))


@pytest.fixture
def bj():
    import build_journal as bjmod
    return bjmod


def locked_unlink(monkeypatch, target):
    """让 Path.unlink 只对 target 抛 PermissionError，其余路径照常删。"""
    import pathlib
    orig = pathlib.Path.unlink

    def fake_unlink(self, missing_ok=False):
        if self == target:
            raise PermissionError(13, "模拟：另一个程序正在使用此文件")
        return orig(self, missing_ok=missing_ok)

    monkeypatch.setattr(pathlib.Path, "unlink", fake_unlink)


# ---------------------------------------------------------------- 转换前清理

def test_convert_pdf_locked_old_pdf_friendly_exit(tmp_path, bj, monkeypatch, capsys):
    """复现 FDOC2：转换前删旧 PDF 被占用 → 人话报错、退出码 1，且不再发起转换。"""
    sample = tmp_path / "a.sample.docx"
    sample.write_bytes(b"docx")
    pdf = tmp_path / "a.sample.pdf"
    pdf.write_bytes(b"OLD PDF")
    locked_unlink(monkeypatch, pdf)

    calls = []
    monkeypatch.setattr(bj, "run", lambda *a, **k: calls.append(a))

    with pytest.raises(SystemExit) as ei:
        bj.convert_pdf("soffice-fake", sample, pdf)
    assert ei.value.code == 1, "旧 PDF 删不掉必须以非零退出码停机，不是裸 traceback"
    out = capsys.readouterr().out
    assert "被占用" in out and "关闭阅读器" in out, f"报错不人话：{out[-300:]}"
    assert pdf.name in out, "报错没有点名被占用的文件"
    assert not calls, "旧 PDF 删不掉就不该带着它发起转换"
    assert pdf.exists(), "被占用的文件本就删不掉，不许被部分破坏"


def test_convert_pdf_unlocked_old_pdf_still_deleted(tmp_path, bj, monkeypatch, capsys):
    """正对照：没被占用时旧 PDF 照常删掉、转换照常进行——容错不误伤正常路径。"""
    sample = tmp_path / "b.sample.docx"
    sample.write_bytes(b"docx")
    pdf = tmp_path / "b.sample.pdf"
    pdf.write_bytes(b"OLD PDF")

    def fake_run(cmd, timeout=None, **kw):
        pdf.write_bytes(b"NEW PDF")

    monkeypatch.setattr(bj, "run", fake_run)
    bj.convert_pdf("soffice-fake", sample, pdf)  # 不应抛
    assert pdf.read_bytes() == b"NEW PDF"
    assert "被占用" not in capsys.readouterr().out


# ---------------------------------------------------------------- 无转换器清理

def test_main_no_soffice_locked_old_pdf_friendly_exit(tmp_path, bj, monkeypatch, capsys):
    """复现 FDOC2 第二处：无 LibreOffice 且旧 PDF 被占用 → 同一人话出口、退出码 1。"""
    ruleset = tmp_path / "lockedone.json"
    ruleset.write_text(
        json.dumps({"journal": {"id": "lockedone"}}, ensure_ascii=False),
        encoding="utf-8")
    pdf = tmp_path / "lockedone.sample.pdf"
    pdf.write_bytes(b"OLD PDF")
    locked_unlink(monkeypatch, pdf)

    monkeypatch.setattr(bj, "find_soffice", lambda: None)
    monkeypatch.setattr(bj, "run", lambda *a, **k: None)  # 样板/规格子进程全部桩掉
    monkeypatch.setattr(sys, "argv", ["build_journal.py", str(ruleset)])

    with pytest.raises(SystemExit) as ei:
        bj.main()
    assert ei.value.code == 1
    out = capsys.readouterr().out
    assert "被占用" in out and "关闭阅读器" in out, f"报错不人话：{out[-300:]}"
    assert pdf.name in out, "报错没有点名被占用的文件"
