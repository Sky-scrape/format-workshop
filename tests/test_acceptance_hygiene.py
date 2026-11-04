"""验收层 F：测试自身的纪律。

这些用例不测产品，测测试：临时目录是否真隔离、有没有偷偷联网、
验收文件里有没有把真实数据路径写死。它们失败时，说明测试本身不可信。
"""

from pathlib import Path

import pytest

import acceptance_support

pytestmark = pytest.mark.acceptance


def test_acceptance_work_is_under_tmp(acceptance_env):
    """F-01/F-02：任务目录在 pytest 临时目录下，不在项目 outputs/。"""
    work = acceptance_env["work"]
    tmp = acceptance_env["tmp"]
    assert work.is_relative_to(tmp)
    assert work != acceptance_support.ROOT / "outputs" / "web"
    assert acceptance_support.ROOT / "outputs" not in work.parents


def test_ai_config_is_not_the_real_one(acceptance_env):
    """F-01：AI 配置指向临时文件，不读也不写真实 config/ai.json。"""
    import ai_assist

    assert ai_assist.CONFIG_PATH.is_relative_to(acceptance_env["tmp"])
    assert ai_assist.CONFIG_PATH != acceptance_support.ROOT / "config" / "ai.json"


def test_network_is_blocked(acceptance_env):
    """F-05：验收环境封住 socket，任何真实外连都该立刻失败。"""
    import socket

    with pytest.raises(AssertionError, match="禁止联网"):
        socket.create_connection(("example.com", 80))


def test_acceptance_files_do_not_name_real_data_paths():
    """F-01：验收测试源码里不得出现真实稿件目录或真实输出目录。"""
    files = list((acceptance_support.ROOT / "tests").glob("test_acceptance_*.py"))
    files.append(acceptance_support.ROOT / "tests" / "acceptance_support.py")
    for path in files:
        if path.name == Path(__file__).name:
            continue  # 本文件自身就是检查器，字符串必然出现
        text = Path(path).read_text(encoding="utf-8")
        assert "训练论文" not in text, f"{path.name} 写死了真实稿件目录"
        assert "outputs/web" not in text, f"{path.name} 写死了真实输出目录"


def test_acceptance_suite_is_marked():
    """F-04：验收用例带 acceptance 标记，`pytest -m acceptance` 能单独跑。"""
    files = list((acceptance_support.ROOT / "tests").glob("test_acceptance_*.py"))
    assert files
    for path in files:
        text = Path(path).read_text(encoding="utf-8")
        assert "pytestmark = pytest.mark.acceptance" in text, \
            f"{path.name} 没有 acceptance 标记"
