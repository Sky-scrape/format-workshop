"""Web 健壮性五个已确认问题（F11 / F12 / F33 / F34 / F35）的回归钉子。

- F11  同任务并发渲染无锁：双击提交后两个子进程互删互写同一成品，产出与审计
  可能来自不同轮、后写者胜。→ stage_render 每任务一把 threading.Lock，
  抢不到立刻以 RenderBusy 挡回；busy 响应绝不写 error.json / meta；
  details / result 提交按钮禁用 + 遮罩。
- F12  解析失败后 retry 是死循环：IR 不存在仍导去补录再渲染。→ retry 发现
  IR 缺失改跑 stage_parse；failed.html 按 can_retry 分文案——解析失败给
  「重新上传」出路，渲染失败才说「换刊重试」。
- F33  run() 把 TimeoutExpired 异常对象传给 explain()，explain 第一行
  re.search 对非字符串必抛 TypeError，超时友好页永远不出现。→ explain
  先认 TimeoutExpired 实例再做正则。
- F34  meta.json 非原子写且读取零容错：硬杀留下截断 JSON 后 history 与任务页
  整页 500 无自愈。→ write_meta 改 tmp+os.replace；read_meta 视坏 meta 为
  缺失并从现场文件推断 stage；history 逐条 try/except 跳过坏行。
- F35  retry 路由缺通用异常兜底：非 ConvertError 直接未捕获 500 且不写
  error.json。→ 补 except Exception 走 fail_page，与 details / reject 同款。
"""

import json
import shutil
import subprocess
import threading
from pathlib import Path

import pytest

from conftest import RULES, run_tool

import app


# ---------------------------------------------------------------- 夹具

@pytest.fixture()
def app_module():
    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


@pytest.fixture()
def work(app_module, tmp_path, monkeypatch):
    """把任务目录指到临时位置，测试不碰仓库里的 outputs/web。"""
    w = tmp_path / "work"
    w.mkdir(parents=True)
    monkeypatch.setattr(app_module, "WORK", w)
    return w


@pytest.fixture()
def light_job(work):
    """只登记 meta 的任务：走路由逻辑够用，不必真跑一遍流水线。"""
    d = work / "a1b2c3d4e5f6"
    d.mkdir(parents=True)
    return d


@pytest.fixture()
def delivered_job(work, sample_docx):
    """一个已交付过的任务：真实跑完 parse → render，成品与审计都在
    （参照 test_fix_web_guards.delivered_job）。"""
    job = "f36a1b2c3d4e"
    d = work / job
    d.mkdir(parents=True)
    shutil.copy(sample_docx, d / "manuscript.docx")
    run_tool("parse_docx.py", d / "manuscript.docx", "-o", d / "manuscript.ir.json")
    run_tool("render_from_ir.py", d / "manuscript.ir.json", RULES / "tjyj.json",
             "-o", d / "manuscript.tjyj.docx")
    meta = {"job": job, "journal": "tjyj", "journal_name": "统计研究",
            "original": "source.docx", "stem": "source", "stage": "rendered",
            "docx": "manuscript.tjyj.docx"}
    (d / "meta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    return d


# ---------------------------------------------------------------- F11 任务锁

def test_concurrent_render_second_is_busy_and_first_round_intact(
        app_module, delivered_job, monkeypatch):
    """复现 F11：同一 job 两次快速调用 stage_render，第二次必须被锁挡住，
    且不能删、不能写第一轮的产物。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    audit_before = (d / "manuscript.tjyj.audit.json").read_bytes()

    started, release = threading.Event(), threading.Event()

    def slow_run(args):
        started.set()                      # 此时锁已被第一轮持有
        assert release.wait(30), "第一次渲染没有等到放行信号"
        raise app_module.ConvertError("第一次渲染被测试中止", "")

    monkeypatch.setattr(app_module, "run", slow_run)
    outcome = {}

    def first():
        try:
            outcome["result"] = app_module.stage_render(d, "tjyj", "source.docx")
        except Exception as e:             # noqa: BLE001 —— 线程里只留下异常对象
            outcome["error"] = e

    t = threading.Thread(target=first)
    t.start()
    assert started.wait(10), "第一次渲染没有启动"

    # 双击的第二个请求：必须立刻被 RenderBusy 挡回，而不是排队等第一轮
    with pytest.raises(app_module.RenderBusy):
        app_module.stage_render(d, "tjyj", "source.docx")

    release.set()
    t.join(30)
    assert isinstance(outcome.get("error"), app_module.ConvertError), \
        "第一轮按注入的失败结束，锁应随 finally 释放"

    # 第二个请求没有删第一轮产物，也没有留下 .staging 半成品
    assert (d / "manuscript.tjyj.docx").read_bytes() == docx_before
    assert (d / "manuscript.tjyj.audit.json").read_bytes() == audit_before
    assert not list(d.glob(".staging.*")), "被挡回的请求不许碰 .staging 临时名"

    # 锁已释放：下一轮渲染可以照常进入（失败也要是业务失败，不是忙碌）
    def fast_fail(args):
        raise app_module.ConvertError("业务失败", "")

    monkeypatch.setattr(app_module, "run", fast_fail)
    with pytest.raises(app_module.ConvertError) as ei:
        app_module.stage_render(d, "tjyj", "source.docx")
    assert not isinstance(ei.value, app_module.RenderBusy), "锁没有随 finally 释放"


def test_busy_route_returns_page_without_touching_disk(app_module, light_job,
                                                       client, monkeypatch):
    """忙碌态的响应绝不写 error.json / meta：此刻另一轮正在写这些文件，
    任何落笔都会把两轮的现场搅在一起。"""
    d = light_job
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj", "original": "s.docx",
                              "stage": "rendered", "docx": "manuscript.tjyj.docx"})

    def busy(*a, **k):
        raise app_module.RenderBusy(d.name)

    monkeypatch.setattr(app_module, "stage_render", busy)
    r = client.post(f"/job/{d.name}/details", data={})
    assert r.status_code == 200, "忙碌态要有友好页，不能 500"
    assert "正在处理中" in r.get_data(as_text=True)
    assert not (d / "error.json").exists(), "忙碌态不许写 error.json"
    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta.get("stage") == "rendered", "忙碌态不许把 meta 改成失败态"


def test_details_page_guards_double_submit(app_module, light_job, client):
    """复现 F11 前端：补录页提交按钮要禁用 + 遮罩（沿用工作台的遮罩模式）。"""
    app_module.write_meta(light_job, {"job": light_job.name, "journal": "tjyj",
                                      "original": "s.docx"})
    html = client.get(f"/job/{light_job.name}/details").data.decode("utf-8")
    assert 'id="fillForm"' in html, "补录表单没有 id，脚本无从挂提交守卫"
    assert 'id="renderOverlay"' in html, "补录页没有排版遮罩"
    assert "processing-overlay" in html, "遮罩没用上 base 里现成的样式"
    assert "b.disabled=true" in html, "提交按钮没有被禁用"


def test_result_page_guards_double_submit(app_module, light_job, client):
    """复现 F11 前端：结果页的换刊重排表单同样要禁用 + 遮罩。"""
    app_module.write_meta(light_job, {"job": light_job.name, "journal": "tjyj",
                                      "original": "s.docx",
                                      "docx": "manuscript.tjyj.docx"})
    html = client.get(f"/result/{light_job.name}").data.decode("utf-8")
    assert 'id="retryForm"' in html, "换刊表单没有 id，脚本无从挂提交守卫"
    assert 'id="renderOverlay"' in html, "结果页没有重排遮罩"
    assert "b.disabled=true" in html, "重排按钮没有被禁用"


# ---------------------------------------------------------------- F12 retry 死循环

def test_retry_without_ir_runs_parse_not_render(app_module, light_job, sample_docx,
                                                client, monkeypatch):
    """复现 F12：解析失败（IR 不存在）后调 retry，必须触发 parse 而不是
    再导去补录页渲染——坏文件换任何刊都不可能成功。"""
    d = light_job
    shutil.copy(sample_docx, d / "manuscript.docx")
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "source.docx", "stage": "uploaded"})

    calls = {"parse": None, "render": False}

    def fake_parse(dd, jid, original):
        calls["parse"] = (dd.name, jid, original)
        (dd / "manuscript.ir.json").write_text("{}", encoding="utf-8")
        return {}

    def fail_render(*a, **k):
        calls["render"] = True
        raise AssertionError("IR 缺失时 retry 不该去渲染")

    monkeypatch.setattr(app_module, "stage_parse", fake_parse)
    monkeypatch.setattr(app_module, "stage_render", fail_render)

    r = client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"})
    assert r.status_code == 302 and "/details" in r.headers["Location"], \
        "重跑解析后应回补录页，而不是继续往渲染走"

    assert calls["parse"] == (d.name, "cjyj", "source.docx"), \
        "retry 没有按新刊重跑解析（或参数不对）"
    assert not calls["render"], "IR 缺失时 retry 触发了渲染（死循环还在）"

    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta.get("stage") == "details" and meta.get("journal") == "cjyj"


def test_retry_with_unparseable_file_lands_on_reupload_page(app_module, light_job,
                                                            client):
    """复现 F12 的死循环出口：坏文件 retry 重跑解析仍失败时，失败页必须给
    「重新上传」出路，绝不能再出现「不需要重新上传」加 retry 表单。"""
    d = light_job
    (d / "manuscript.docx").write_bytes(b"not a real docx, just a broken file")
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "broken.docx", "stage": "uploaded"})

    r = client.post(f"/job/{d.name}/retry", data={})
    assert r.status_code == 200, "解析再失败要走错误页，而不是 500"
    html = r.get_data(as_text=True)
    assert "重新上传" in html, "解析失败没有给重新上传的出路"
    assert "不需要重新上传" not in html, "解析失败还在承诺「不需要重新上传」"
    assert f'action="/job/{d.name}/retry"' not in html, \
        "解析失败还给 retry 表单，死循环又接上了"


def test_render_failure_page_keeps_retry_promise(app_module, light_job, client):
    """对照：渲染阶段失败（IR 在）才保留「不需要重新上传」与换刊表单。"""
    d = light_job
    (d / "manuscript.ir.json").write_text("{}", encoding="utf-8")
    (d / "error.json").write_text(json.dumps(
        {"title": "转换没有完成", "hint": "", "detail": ""}, ensure_ascii=False),
        encoding="utf-8")
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "s.docx", "docx": "manuscript.tjyj.docx"})

    html = client.get(f"/job/{d.name}/failed").get_data(as_text=True)
    assert "不需要重新上传" in html, "渲染失败的换刊承诺被误改了"
    assert f'action="/job/{d.name}/retry"' in html, "渲染失败页丢了换刊重试表单"


# ---------------------------------------------------------------- F33 超时翻译

def test_explain_handles_timeout_instance():
    """复现 F33：explain(TimeoutExpired 实例) 必须返回「转换超时」，
    而不是在第一行 re.search 上抛 TypeError。"""
    e = subprocess.TimeoutExpired(cmd=["tools/render_from_ir.py"], timeout=300)
    title, hint = app.explain(e)
    assert title == "转换超时" and hint, f"超时异常对象被翻译成了 {title!r}"
    # 字符串路径照常工作（stderr 里带 timed out 的老口径）
    assert app.explain("Command '[…]' timed out after 300 s")[0] == "转换超时"


def test_run_timeout_raises_friendly_error(app_module, monkeypatch):
    """run() 的超时分支要给出「转换超时」，而不是 TypeError。"""
    monkeypatch.setattr(app_module, "STEP_TIMEOUT", 1)
    with pytest.raises(app_module.ConvertError) as ei:
        app_module.run(["-c", "import time; time.sleep(30)"])
    assert ei.value.title == "转换超时"


# ---------------------------------------------------------------- F34 meta 健壮性

def test_write_meta_is_atomic_and_read_meta_infers_from_scene(app_module, work):
    """复现 F34：硬杀留下的截断 meta 不许抛异常；write_meta 不留临时文件；
    读不出来时从现场文件推断 stage 与可用字段。"""
    good = work / "111111111111"
    good.mkdir()
    app_module.write_meta(good, {"job": good.name, "original": "t.docx"})
    assert json.loads((good / "meta.json").read_text(encoding="utf-8"))["original"] == "t.docx"
    assert not (good / "meta.json.tmp").exists(), "原子写不许残留临时文件"

    dead = work / "222222222222"
    dead.mkdir()
    (dead / "meta.json").write_text('{"job": "2222', encoding="utf-8")   # 硬杀截断
    (dead / "manuscript.docx").write_bytes(b"x")
    (dead / "manuscript.ir.json").write_text('{"source": "我的论文.docx"}',
                                             encoding="utf-8")
    (dead / "manuscript.cjyj.docx").write_bytes(b"x")
    (dead / "manuscript.cjyj.anon.docx").write_bytes(b"x")
    m = app_module.read_meta(dead)          # 不抛即过第一关
    assert m["original"] == "我的论文.docx" and m["stem"] == "我的论文"
    assert m["docx"] == "manuscript.cjyj.docx" and m["journal"] == "cjyj", \
        "匿名稿不能被当成署名成品，刊 id 也不该带上 .anon"
    assert "stage" not in m, "没有 error.json 就不能断言失败态"

    (dead / "error.json").write_text("{}", encoding="utf-8")
    assert app_module.read_meta(dead).get("stage") == "failed", \
        "最后一次落笔是失败页，就该推断为失败态"


def test_truncated_meta_keeps_history_and_failed_page_alive(app_module, work, client):
    """复现 F34：一条坏 meta 不得把 history 与任务页整页拖成 500。"""
    good = work / "333333333333"
    good.mkdir()
    app_module.write_meta(good, {"job": good.name, "original": "好稿.docx",
                                 "journal": "tjyj", "stage": "rendered",
                                 "docx": "manuscript.tjyj.docx"})
    bad = work / "444444444444"
    bad.mkdir()
    (bad / "meta.json").write_text('{"job": "4444', encoding="utf-8")   # 截断
    (bad / "error.json").write_text(json.dumps(
        {"title": "解析失败", "hint": "", "detail": ""}, ensure_ascii=False),
        encoding="utf-8")

    page = client.get("/history")
    assert page.status_code == 200, "一条坏 meta 把 history 拖成 500 了"
    assert "好稿.docx" in page.get_data(as_text=True), "好任务的记录被坏行连累丢失"

    dead = client.get(f"/job/{bad.name}/failed")
    assert dead.status_code == 200, "任务失败页在坏 meta 下 500，无法显示失败态"
    assert "解析失败" in dead.get_data(as_text=True)


# ---------------------------------------------------------------- F35 retry 兜底

def test_retry_unexpected_exception_goes_to_fail_page(app_module, light_job,
                                                      client, monkeypatch):
    """复现 F35：retry 链路里非 ConvertError（如 PermissionError）必须走
    fail_page 并留下 error.json，不能裸 500。"""
    d = light_job
    (d / "manuscript.docx").write_bytes(b"x")
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "s.docx", "stage": "rendered",
                              "docx": "manuscript.tjyj.docx"})

    def boom(*a, **k):
        raise PermissionError(13, "成品文件被另一个程序占用")

    monkeypatch.setattr(app_module, "stage_render", boom)
    r = client.post(f"/job/{d.name}/retry", data={})
    assert r.status_code == 200, "非 ConvertError 被 500 了，兜底没生效"
    assert "转换没有完成" in r.get_data(as_text=True)

    err = json.loads((d / "error.json").read_text(encoding="utf-8"))
    assert "PermissionError" in err.get("detail", ""), \
        "error.json 没有留下技术细节，failed 页无法还原现场"
