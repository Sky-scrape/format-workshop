"""D-03 回归钉：WebView2 用户数据目录落固定位置，不再一次性 %TEMP%。

pywebview 默认 private_mode=True，WinForms 后端此时把用户数据目录设成
tempfile.TemporaryDirectory().name——每次启动冷建一套完整浏览器 profile
（%TEMP% 实测含 EBWebView 的 tmp 目录数十个，单个 ~70MB）；exit 收尾的
os._exit 又抢在 clear_user_data 之前，临时目录基本必留，%TEMP% 残留不属
于任何卸载器管辖，GPU/代码缓存也无法跨次复用。

现在 webview.start(private_mode=False, storage_path=固定目录)：跨次复用、
二次启动更快；目录按窗口标题分使用版/开发版；安装包 [UninstallDelete]
带走缓存目录。隐私前提：app.py 对非 /static 响应一律 Cache-Control:
no-store——稿件内容不进 UDF，缓存只落静态资产（一并钉住）。
"""

import os
from pathlib import Path

import app
import desktop

ROOT = Path(__file__).resolve().parent.parent


def test_udf_lands_under_localappdata(tmp_path, monkeypatch):
    """固定目录在 %LOCALAPPDATA% 下，不再每次都是临时目录。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))

    udf = desktop.webview_udf()
    assert udf == tmp_path / "appdata" / "格式工坊" / "webview"
    assert "TEMP" not in str(udf).upper() or "appdata" in str(udf)


def test_udf_is_per_edition(tmp_path, monkeypatch):
    """使用版与开发版分目录（判据与窗口标题同源）：并存互不锁。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))

    assert desktop.webview_udf().parent.name == "格式工坊"

    (tmp_path / ".venv").mkdir()
    assert desktop.webview_udf().parent.name == "格式工坊 · 开发版"


def test_udf_has_a_fallback_without_localappdata(tmp_path, monkeypatch):
    """LOCALAPPDATA 缺失（非 Windows 的开发跑）落仓库内，不炸。"""
    monkeypatch.setattr(desktop, "ROOT", tmp_path)
    monkeypatch.delenv("LOCALAPPDATA", raising=False)

    udf = desktop.webview_udf()
    assert tmp_path in udf.parents or udf.is_relative_to(tmp_path)


def test_main_starts_webview_with_fixed_storage():
    """start() 显式 private_mode=False + storage_path=固定目录。"""
    src = (ROOT / "desktop.py").read_text(encoding="utf-8")
    main_body = src.split("def main(", 1)[1]
    assert "private_mode=False" in main_body
    assert "storage_path=str(udf)" in main_body
    assert "udf = webview_udf()" in main_body
    # 钉的是「不造临时 profile 目录」：webview_udf 的文档串还留着 D-03 的
    # 历史叙述（提 tempfile.TemporaryDirectory），只约束真调用——代码里
    # tempfile 只许出现在 log_line 的 %TEMP% 诊断兜底一处（D-02/F7-4），
    # 且必须不是 gettempdir() 建目录那种「当 home 用」的形态。
    code = "\n".join(
        ln for ln in src.splitlines()
        if not ln.lstrip().startswith("#") and "——每次启动冷建一套完整浏览器" not in ln)
    assert "TemporaryDirectory(" not in code, "壳自身不再造临时 profile 目录"
    assert code.count("tempfile.gettempdir") == 1, \
        "%TEMP% 只许诊断兜底一处，不给 profile 用"


def test_installer_removes_the_cached_udf():
    """卸载时带走固定 UDF（缓存而已，无用户数据）。"""
    src = (ROOT / "tools" / "installer" / "installer.iss").read_text(
        encoding="utf-8")
    assert 'Type: filesandordirs; Name: "{localappdata}\\格式工坊\\webview"' in src


def test_privacy_precondition_no_store_for_user_pages():
    """UDF 固定的隐私前提：非 /static 响应 no-store，稿件内容不进缓存。"""
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'Cache-Control", "no-store"' in src
    assert 'startswith("/static/")' in src
