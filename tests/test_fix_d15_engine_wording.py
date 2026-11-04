"""D-15 回归钉：旧「后台服务/端口」话语不得在用户可见文案里回潮。

引擎已收进壳进程（无端口、无后台服务），但四处还在说已不存在的「后台服务」：
设置页关窗两选项、500 错误页标题「服务内部错误」、两个 bat 头注
（over the 5001 service / over the service）、README 首段的 127.0.0.1:5000。
附带可发现性问题：关窗偏好藏在标题为「AI 设置」的页面——页名改「设置」，
AI 配置区块用自己的标题区分。
"""

from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture()
def client():
    import app

    app.app.config["TESTING"] = True
    return app.app.test_client()


def test_settings_page_is_named_settings_and_ai_block_self_titles(client):
    """想改关窗行为的人找得到「设置」；AI 配置区块不靠页名撑辨识度。"""
    html = client.get("/settings").get_data(as_text=True)
    assert "<h1>设置</h1>" in html
    assert "AI 设置" not in html
    assert "AI 服务连接" in html, "AI 区块要有自己的标题"


def test_close_option_copy_speaks_engine_not_service(client):
    """关窗两选项按进程内引擎口径说话，「后台服务」一词不许出现。"""
    html = client.get("/settings").get_data(as_text=True)
    assert "点 × 时程序连同内置引擎一并结束" in html
    assert "点 × 后程序收进屏幕右下角托盘" in html
    assert "软件继续运行" in html and "从托盘图标退出" in html
    assert "后台服务" not in html


def test_500_error_page_titled_software_error():
    """500 页标题不再是「服务内部错误」——本软件没有「服务」这个概念可内部。"""
    import app

    # error.html 经 base.html 引用 request（导航高亮），要有请求上下文才能渲染
    with app.app.test_request_context("/"):
        html, code = app.server_error(None)
    assert code == 500
    assert "软件内部错误" in html
    assert "服务内部错误" not in html


def test_readme_opens_with_inprocess_engine_wording():
    """README 首段的软件形态描述与默认双击形态（桌面壳）一致。"""
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    lead = next(line for line in readme.splitlines() if "稿子不出本机" in line)
    assert "127.0.0.1:5000" not in lead, "首段不再用端口形态介绍软件"
    assert "进程内引擎" in lead and "不监听端口" in lead
    # 「后台服务」只允许以否定口径出现（不起后台服务），与 README 后文同规约
    assert "不起后台服务" in lead
    lead_stripped = lead.replace("不起后台服务", "")
    assert "后台服务" not in lead_stripped


@pytest.mark.parametrize("bat", ["开发版.bat", "格式工坊.bat"])
def test_bat_headers_describe_the_shell_not_a_service(bat):
    """两个 bat 头注按进程内引擎口径改写，且仍保持 ASCII-only（自家规约）。"""
    src = (ROOT / bat).read_text(encoding="utf-8")
    assert src.isascii(), f"{bat} 必须保持纯 ASCII"
    assert "over the" not in src and "auto-starting" not in src
    assert "keepalive" not in src, "壳不再经 keepalive，头注别再提它"
    assert "in-process engine" in src
