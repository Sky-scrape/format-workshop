"""工作台页面：守两条曾经把人卡死的老路。

1. **报错页必须还带期刊列表**。`/convert` 的三个报错分支原先各写一遍
   `render_template`，都漏了 `groups` / `default_id`：页面能出来，但「选择期刊」
   里一本都没有，也没有一个单选被勾上，「开始格式转换」永远是灰的，只能刷新重来。
2. **期刊单选必须在表单里**。弹层为了避免被面板的 `overflow` 裁掉，被 JS 挪到了
   `<body>` 上，连带里面的 `<input name="journal">` 一起离开 `<form>`，于是无论
   选哪本，POST 都不带 journal 字段。现在靠 `form="convertForm"` 把关联显式钉住。
"""

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
RADIO = re.compile(r'<input type="radio"[^>]*>')
CHECKED = re.compile(r'<input type="radio"[^>]*\bchecked\b[^>]*>')


@pytest.fixture()
def app_module():
    import app

    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


def test_error_page_keeps_journal_list(app_module, client):
    """不带文件直接提交，报错页里依然能挑期刊。"""
    html = client.post("/convert", data={}).data.decode("utf-8")

    assert len(RADIO.findall(html)) == len(app_module.load_journals())
    assert CHECKED.search(html), "报错页没有预选任何期刊，按钮永远是灰的"


def test_error_page_keeps_chosen_journal(app_module, client):
    """报错后保留刚选的那本，重传文件就能直接再提交。"""
    items = app_module.load_journals()
    jid = items[-1]["id"]

    html = client.post("/convert", data={"journal": jid}).data.decode("utf-8")

    checked = CHECKED.search(html)
    assert checked, "没留住用户刚选的期刊"
    assert f'value="{jid}"' in checked.group(0)


def test_radios_are_tied_to_the_form(client):
    """单选靠 form 属性挂在表单上，弹层被挪到 body 也不影响提交。"""
    html = client.get("/workspace").data.decode("utf-8")

    radios = RADIO.findall(html)
    assert radios, "工作台没有渲染出期刊单选"
    assert all('form="convertForm"' in r for r in radios)


def test_legacy_doc_gets_its_own_waiting_copy(client):
    """.doc 要多走一步本机 Office 转换，等待遮罩必须说清此刻在做什么。

    这不是文案洁癖：.doc 的 COM 转换要 5–15 秒，比 docx 多一整步。遮罩上
    如果只有「正在处理」，用户会以为卡死，然后刷新页面把整轮白等掉。

    文案的边界仍是 F57 那条：只写**此刻真的在发生什么**，不轮播「正在解析／
    正在体检」这种按秒编排的假进度——我们确实知道 .doc 要先转格式、也
    确实知道它慢，这属于事实，不属于编造。
    """
    html = client.get("/workspace").data.decode("utf-8")

    # 三个槽位必须能被脚本取到，否则 setProcCopy 是空转
    for el in ('id="procTitle"', 'id="procHint"', 'id="procNote"'):
        assert el in html, f"等待遮罩缺少 {el}"
    assert "PROC_TEXT" in html, "没有按文件类型切换等待文案的分支"
    assert "正在转换 .doc" in html
    # 默认分支不能被顶掉：传 docx 时不该出现 .doc 字样
    assert "正在处理，可在历史页查看进度" in html
