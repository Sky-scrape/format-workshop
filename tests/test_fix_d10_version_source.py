"""D-10 回归钉：分发产物的版本号只有一个真源。

曾经的漂移：build_portable 的 zip 名用 --python 参数（3.12.10，解释器版本），
build_installer 用自己另写的 --version 默认值（1.0.0），dist 里
「格式工坊-3.12.10.zip」与「格式工坊-setup-1.0.0.exe」并存——使用者无法
判断哪个新，文件名也对不上软件版本；包内没有构建来源标记。

现在 zip 名与安装包默认版本都取 app.py 的 APP_VERSION（tools/app_version.py
正则解析源码，不为取版本 import 整个 app），构建时把版本与 git rev 写进
包内 BUILD-INFO.txt；状态栏/帮助弹窗本就经 context_processor 用这同一常量。
"""

import re
import sys
from pathlib import Path

import pytest

from conftest import ROOT, TOOLS

import app_version as appver


def test_single_source_matches_app_constant():
    """读取口解析的就是 app.py 里那个常量（状态栏同源）。"""
    import app

    assert appver.app_version() == app.APP_VERSION


def test_parses_from_source_without_importing_app(tmp_path):
    """假 app.py 也能解析（构建机没装依赖也不会炸）；缺常量当场失败。"""
    (tmp_path / "app.py").write_text(
        'APP_VERSION = "2.3.4"       # 注释\n', encoding="utf-8")
    assert appver.app_version(tmp_path) == "2.3.4"

    (tmp_path / "app.py").write_text("APP_VERSION = 'x'\n", encoding="utf-8")
    with pytest.raises(SystemExit):
        appver.app_version(tmp_path)


def test_build_info_records_version_source_and_time(tmp_path):
    """包内 BUILD-INFO.txt：软件版本、解释器版本、git rev、构建时间。"""
    import build_portable

    out = tmp_path / "格式工坊"
    out.mkdir()
    build_portable.write_build_info(out, "1.2.3", "3.12.10")

    info = (out / "BUILD-INFO.txt").read_text(encoding="utf-8")
    assert "软件版本 1.2.3" in info
    assert "Python 3.12.10" in info
    assert "构建时间" in info
    assert info.splitlines()[2].startswith("git ")


def test_zip_name_uses_app_version_not_interpreter_version():
    """zip 名里的 {ver} 必须是软件版本；--python 只进解释器目录的构建。"""
    src = (TOOLS / "build_portable.py").read_text(encoding="utf-8")
    assert "ver = appver.app_version()" in src
    assert 'f"{OUT_NAME}-{ver}.zip"' in src, "zip 名用软件版本"
    # 「py_ver = a.python」是合法的（解释器版本变量），裸「ver = a.python」不许有
    assert not re.search(r"(?<!py_)ver = a\.python", src), \
        "--python 不再冒充软件版本"
    # 解释器版本改走 py_ver，embeddable 三处调用不得被牵连
    for call in ("fetch_python(ROOT / \"build\" / \"cache\", py_ver)",
                 "configure_pth(py_dir, py_ver)",
                 "install_deps(py_dir, py_ver)"):
        assert call in src, f"解释器版本参数丢了：{call}"
    assert "write_build_info(out, ver, py_ver)" in src


def test_installer_default_version_comes_from_the_same_source():
    src = (TOOLS / "build_installer.py").read_text(encoding="utf-8")
    assert 'default=appver.app_version()' in src
    assert 'default="1.0.0"' not in src, "安装包不再自写一份默认版本"


def test_status_bar_reads_the_same_constant():
    """状态栏的版本号也是这处真源（context_processor 注入）。"""
    import app

    html_src = (ROOT / "templates" / "base.html").read_text(encoding="utf-8")
    assert "v{{ app_version }}" in html_src
    body = app.inject_version()
    assert body["app_version"] == app.APP_VERSION


def test_build_scripts_import_version_module_lazily_safe():
    """两个构建脚本都不 import app（构建机没有 flask 也能跑）。"""
    for name in ("build_portable.py", "build_installer.py"):
        src = (TOOLS / name).read_text(encoding="utf-8")
        assert "import app\n" not in src, f"{name} 不许为取版本 import app"
        assert "app_version" in src
