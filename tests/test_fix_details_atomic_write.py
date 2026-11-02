"""details.json 锁外写的原子性回归钉（2026-11-02 审读三问题之三）。

F11 的锁覆盖了 stage_render 内部，但没覆盖路由入口的共享写：details POST
在抢锁前落笔 details.json，而持锁的上一轮渲染在 _stage_render 首步就要
read_json 它（对半截 JSON 抛异常）。原先的 write_text 非原子写留下毫秒级
窗口——命中时第一轮以 JSONDecodeError 走 except Exception，误报失败页。
「同值幂等写无害」的论证有缺口：同值幂等只保证值相同，不保证写入动作与
持锁者读不交错。

修法与 write_meta（F34）同款：先落临时名再 os.replace。读方要么见完整旧
值、要么见完整新值，永远见不到半个 JSON。

两个测试面：
- **确定性交错**：用 os.replace 钩子在「临时名已写完、尚未替换」的临界点
  模拟并发读者——此刻读主文件必须是完整的旧值。这是非原子写做不到的性质
  （写到一半的文件就是半截 JSON）。
- **并发烟测**：读者线程紧循环解析 details.json，同时连续 POST 表单，
  任何一次半截 JSON 都会以异常形式暴露。
"""

import json
import threading
import time

import pytest


@pytest.fixture()
def app_module():
    import app

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


def _post_details(client, job, abstract, monkeypatch):
    """提交补录表单；渲染被桩掉，焦点只在 details.json 的落盘。"""
    import app as app_module

    monkeypatch.setattr(app_module, "stage_render", lambda *a, **k: None)
    return client.post(f"/job/{job}/details",
                       data={"english_abstract": abstract, "classification": "F06"})


def test_details_write_is_atomic_at_the_commit_point(app_module, light_job,
                                                     client, monkeypatch):
    """确定性交错：替换发生前一刻并发读，必须读到完整的旧值。

    os.replace 的前一瞬间相当于「写动作进行中」的最坏时点——非原子写的
   世界里此刻文件可能只写了一半；原子写下主文件仍是完整的旧 JSON。
    """
    d = light_job
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "s.docx"})
    # 预置旧 details.json：临界点必须有「完整的旧值」可读。没有旧值，替换
    # 前一刻读到 None，「读到完整旧值」的断言就是空的。
    preset_old = "第一轮提交的英文摘要。" * 40
    (d / "details.json").write_text(
        json.dumps({"english_abstract": preset_old}, ensure_ascii=False),
        encoding="utf-8")
    submitted = "第二轮表单提交的英文摘要。" * 40

    real_replace = app_module.os.replace
    seen = {"replaced_details": False, "during": None}

    def spy_replace(src, dst):
        # 只认 details.json 的落位（meta.json 的 os.replace 与本测试无关，
        # 不分目标会把 write_meta 的触发误当 details 的提交点，探针读到
        # 替换完成后的新值——审读遗留：假绿测试修正）。
        if str(dst).endswith("details.json"):
            seen["replaced_details"] = True
            # 临界点：新值已写进临时名、主文件还是旧值。并发读者此刻读主文件。
            seen["during"] = app_module.read_json(d / "details.json")
        return real_replace(src, dst)

    monkeypatch.setattr(app_module.os, "replace", spy_replace)
    r = _post_details(client, d.name, submitted, monkeypatch)
    assert r.status_code in (200, 302), "路由本身不许失败"

    # 实现若退回非原子 write_text，details.json 的 os.replace 不会发生，
    # 这条断言即红——测试对回归有区分度，不再假绿。
    assert seen["replaced_details"], "details.json 必须经 os.replace 原子落位"

    during = seen["during"]
    assert isinstance(during, dict), "临界点读到了非 JSON 内容（半截文件）"
    assert during["english_abstract"] == preset_old, \
        "临界点必须读到完整的旧值，而不是半截新值"

    final = app_module.read_json(d / "details.json")
    assert final["english_abstract"] == submitted
    assert not (d / "details.json.tmp").exists(), "临时名必须清掉，不许留垃圾"


def test_details_write_survives_reader_racing_writer(app_module, light_job,
                                                     client, monkeypatch):
    """并发烟测：读者紧循环解析，写入方连续 POST，半截 JSON 必现形。"""
    d = light_job
    app_module.write_meta(d, {"job": d.name, "journal": "tjyj",
                              "original": "s.docx"})
    payloads = [f"第 {i} 轮提交的英文摘要，内容足够长以放大写入窗口。" * 30
                for i in range(4)]

    stop = threading.Event()
    errors, observed = [], set()

    def reader():
        p = d / "details.json"
        while not stop.is_set():
            try:
                data = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
            except ValueError as e:
                # 半截/拼接 JSON 会从这里现形：原子写下这必须永不发生
                errors.append(e)
                return
            except OSError:
                # Windows 共享冲突是毫秒不到的瞬时窗口，生产读方（_stage_render）
                # 同样重试；只有重试耗尽才与半截 JSON 同罪
                try:
                    data = json.loads(p.read_text(encoding="utf-8"))
                except (OSError, ValueError) as e:
                    errors.append(e)
                    return
            v = data.get("english_abstract")
            if v:
                observed.add(v)
            time.sleep(0.001)      # 贴近真实读频率，也把句柄还给写入方

    t = threading.Thread(target=reader)
    t.start()
    try:
        for abstract in payloads:
            _post_details(client, d.name, abstract, monkeypatch)
    finally:
        stop.set()
        t.join(30)

    assert not errors, f"并发读见到 {len(errors)} 次坏 JSON：{errors[:2]}"
    # 读者见到的每一个值都必须是某一轮的完整值（绝无半截/拼接值）
    assert observed <= set(payloads), f"读到了不完整或拼接的值：{observed - set(payloads)}"
    # 收尾后盘上必须是最后一轮的完整值（调度让读者是否赶上末轮不确定，
    # 但文件本身必须完整且与提交一致）
    final = json.loads((d / "details.json").read_text(encoding="utf-8"))
    assert final["english_abstract"] == payloads[-1]
    assert not (d / "details.json.tmp").exists()
