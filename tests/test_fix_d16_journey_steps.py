"""D-16 回归钉：同一趟转换旅程只允许一套 4 步模型，且定义只有一处。

曾经的分叉：工作台流程条写死 3 步「1 上传 / 2 选刊 / 3 生成成品」且 1、2 两步
静态同亮；/convert 成功 redirect 到补录页（必经阶段）顶部却变成 4 步
「✓上传 / ✓选刊 / 3 补录 / 4 出稿」，两处组件样式也不同。新用户在
「3 步、正处第 2 步」的心智里点了开始转换，落地看到「第 3 步·补录」。

现在步骤定义收敛在模板 templates/_journey.html（宏 journey + 4 步常量），
工作台与补录页都从这里取；工作台是旅程起点（done=0, cur=1），第 2 步由
页面脚本随选刊点亮，3/4 步置灰。
"""

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "templates"

LABELS = ["上传稿件", "选择期刊", "补录信息", "生成成品"]
SPAN = re.compile(
    r'<span class="([^"]*)" data-step="(\d+)"><i>([^<]*)</i>([^<]*)</span>')


@pytest.fixture()
def client():
    import app

    app.app.config["TESTING"] = True
    return app.app.test_client()


def _journey_spans(html):
    m = re.search(r'<div class="journey"[^>]*>(.*?)</div>', html, re.S)
    assert m, "页面里没有渲染出旅程步骤条"
    return SPAN.findall(m.group(1))


def test_macro_is_the_single_four_step_model():
    """宏渲染的就是那套 4 步：done 步打 ✓、cur 步点亮、其余置灰。"""
    import app

    tpl = app.app.jinja_env.get_template("_journey.html")
    html = tpl.module.journey(done=2, cur=3)

    spans = _journey_spans(html)
    assert [s[3] for s in spans] == LABELS, "4 步的名字/顺序变了"
    assert [(s[1], s[0], s[2]) for s in spans] == [
        ("1", "done", "✓"), ("2", "done", "✓"),
        ("3", "now", "3"), ("4", "", "4")]


def test_workspace_is_the_journey_start_and_last_steps_are_gray(client):
    """工作台 = 旅程起点：第 1 步点亮、2–4 置灰；不再是 3 步、1/2 静态同亮。"""
    spans = _journey_spans(client.get("/workspace").get_data(as_text=True))

    assert [s[3] for s in spans] == LABELS
    assert [(s[0], s[2]) for s in spans] == [
        ("now", "1"), ("", "2"), ("", "3"), ("", "4")]


def test_workspace_keeps_lighting_step_two_on_journal_pick(client):
    """页面脚本里留着选刊点亮第 2 步的钩子（含预选刊的 change 分发）。"""
    src = (TEMPLATES / "workspace.html").read_text(encoding="utf-8")
    assert 'flowPick=document.querySelector(\'.journey [data-step="2"]\')' in src
    assert "flowPick.classList.add('now')" in src


def test_details_composes_the_same_macro():
    """补录页取同一宏：前两步已完成、当前在第 3 步补录。"""
    src = (TEMPLATES / "details.html").read_text(encoding="utf-8")
    assert "{{ journey(done=2, cur=3) }}" in src
    assert "journey(done=0, cur=1)" not in src, "补录页不是旅程起点，别抄工作台的参数"


def test_no_template_hardcodes_the_steps_anymore():
    """步骤定义不许再分叉：步骤标签与步骤标记只允许出现在 _journey.html。

    只挑旅程独有的两个标签做指纹（「上传稿件」「选择期刊」会撞上
    「请选择期刊」这类普通文案），外加 done/now 步骤标记的正则。
    """
    step_mark = re.compile(r'<span class="(?:done|now)"><i>[✓0-9]</i>')
    offenders = []
    for p in TEMPLATES.glob("*.html"):
        if p.name == "_journey.html":
            continue
        src = p.read_text(encoding="utf-8")
        if "生成成品" in src or "补录信息" in src or step_mark.search(src):
            offenders.append(p.name)
    assert not offenders, f"这些模板各自写了一份步骤定义：{offenders}"
    for page in ("workspace.html", "details.html"):
        assert '{% from "_journey.html" import journey %}' in (
            TEMPLATES / page).read_text(encoding="utf-8")


def test_old_component_names_are_gone(client):
    """旧组件名（flow-steps / flow-step）与本地 .steps 样式不允许残留。"""
    for name in ("base.html", "workspace.html", "details.html"):
        src = (TEMPLATES / name).read_text(encoding="utf-8")
        assert "flow-step" not in src and "flow-steps" not in src
        assert '.steps{' not in src, f"{name} 还留着本地步骤条样式"
    html = client.get("/workspace").get_data(as_text=True)
    assert 'class="journey"' in html
