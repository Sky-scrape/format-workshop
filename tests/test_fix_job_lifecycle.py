"""任务提交、删除与下载边界的并发/故障回归测试。

覆盖单次成功测试看不到的交错：
- JSON 多写者不能使用同一个临时文件；
- 忙碌请求不得在锁外改写 details/reject/citation；
- 删除与 TTL 不得绕过渲染锁，过期链接不得继续下载；
- meta 中的成品名不能逃逸任务目录；
- 截断 IR 要触发重解析；
- 成品提交中途失败时，旧成品、审计与 meta 必须恢复。
"""

import json
import os
import shutil
import time
from pathlib import Path

import pytest

import app
from conftest import RULES, run_tool


@pytest.fixture()
def app_module():
    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


@pytest.fixture()
def work(app_module, tmp_path, monkeypatch):
    w = tmp_path / "work"
    w.mkdir(parents=True)
    monkeypatch.setattr(app_module, "WORK", w)
    return w


@pytest.fixture()
def light_job(work):
    d = work / "a1b2c3d4e5f6"
    d.mkdir()
    return d


def _meta(app_module, d, **updates):
    data = {"job": d.name, "journal": "tjyj", "journal_name": "统计研究",
            "original": "source.docx", "stem": "source", "stage": "rendered",
            "docx": "manuscript.tjyj.docx"}
    data.update(updates)
    app_module.write_meta(d, data)
    return data


@pytest.fixture()
def delivered_job(app_module, work, sample_docx):
    """真实跑一轮，提交故障测试才能核对旧 DOCX/审计是否逐字节恢复。"""
    d = work / "f36a1b2c3d4e"
    d.mkdir()
    shutil.copy(sample_docx, d / "manuscript.docx")
    run_tool("parse_docx.py", d / "manuscript.docx", "-o", d / "manuscript.ir.json")
    run_tool("render_from_ir.py", d / "manuscript.ir.json", RULES / "tjyj.json",
             "-o", d / "manuscript.tjyj.docx")
    _meta(app_module, d)
    # 换刊/重排要过结构性阻断：tjyj 的合成稿没有英文标题，先在补录记录里备好。
    (d / "details.json").write_text(json.dumps(
        {"english_title": "The Spillover Effects of US Monetary Policy"},
        ensure_ascii=False), encoding="utf-8")
    return d


# ---------------------------------------------------------------- 原子写与锁外副作用


def test_atomic_json_writers_use_distinct_temp_files(app_module, light_job, monkeypatch):
    """连续两次写入直接捕获临时名：必须不同、同目录且结束后清理。

    不用屏障强行调度线程；Windows 下快线程先完成会让屏障假失败。
    唯一临时名这一结构断言直接钉住多写者互相覆盖/删除临时名的根因。
    """
    d = light_job
    target = d / "details.json"
    real_replace = app_module.replace_with_retry
    temporary_names = []

    def spy_replace(src, dst):
        temporary_names.append(Path(src).name)
        assert Path(src).parent == d and Path(dst) == target
        return real_replace(src, dst)

    monkeypatch.setattr(app_module, "replace_with_retry", spy_replace)
    for i in range(2):
        app_module.write_json_atomic(target, {"value": f"完整值-{i}"})

    assert len(set(temporary_names)) == 2, f"两个写者共用了临时名：{temporary_names}"
    assert json.loads(target.read_text(encoding="utf-8"))["value"] == "完整值-1"
    assert not list(d.glob(".details.json.*.tmp")), "临时文件必须清干净"


def test_busy_details_preserves_previous_inputs(app_module, light_job, client,
                                                monkeypatch):
    d = light_job
    _meta(app_module, d, citation=True)
    # 这个任务只登记了 meta、没有 IR；本用例钉锁语义，先绕开缺件体检。
    monkeypatch.setattr(app_module, "_rows_raw", lambda *a, **k: [])
    old = {"english_abstract": "上一轮摘要", "funding": "上一轮基金"}
    app_module.write_json_atomic(d / "details.json", old)
    before = (d / "meta.json").read_bytes()
    lock = app_module.job_lock(d.name)
    assert lock.acquire(blocking=False)
    try:
        # 使用真实 stage_render：锁被持有时自然触发 RenderBusy，prepare 不得执行。
        r = client.post(f"/job/{d.name}/details", data={
            "english_abstract": "被挡回的摘要不应落盘", "funding": "被挡回的基金"})
    finally:
        lock.release()

    assert r.status_code == 200 and "正在处理中" in r.get_data(as_text=True)
    assert json.loads((d / "details.json").read_text(encoding="utf-8")) == old
    assert (d / "meta.json").read_bytes() == before, "忙碌请求改写了 citation 或 meta"


def test_busy_reject_preserves_previous_selection(app_module, light_job, client):
    d = light_job
    _meta(app_module, d)
    app_module.write_json_atomic(d / "reject.json", {"ids": ["old-id"]})
    before = (d / "reject.json").read_bytes()
    lock = app_module.job_lock(d.name)
    assert lock.acquire(blocking=False)
    try:
        r = client.post(f"/job/{d.name}/reject", data={"rid": "new-id"})
    finally:
        lock.release()

    assert r.status_code == 200 and "正在处理中" in r.get_data(as_text=True)
    assert (d / "reject.json").read_bytes() == before


# ---------------------------------------------------------------- 删除与 TTL


def test_delete_busy_job_is_honest_and_preserves_lock(app_module, light_job, client):
    d = light_job
    _meta(app_module, d)
    marker = d / "manuscript.tjyj.docx"
    marker.write_bytes(b"old")
    lock = app_module.job_lock(d.name)
    assert lock.acquire(blocking=False)
    try:
        r = client.post(f"/job/{d.name}/delete")
        assert r.status_code == 409
        assert "没有删除" in r.get_data(as_text=True)
        assert marker.read_bytes() == b"old"
        assert app_module.job_lock(d.name) is lock, "忙碌时回收旧锁会制造第二把锁"
    finally:
        lock.release()


def test_sweep_skips_busy_and_unrelated_dirs(app_module, work):
    busy = work / "111111111111"
    idle = work / "222222222222"
    other = work / "keep-me"
    for d in (busy, idle, other):
        d.mkdir()
        (d / "marker").write_text("x", encoding="utf-8")
    lock = app_module.job_lock(busy.name)
    assert lock.acquire(blocking=False)
    old = time.time() - app_module.JOB_TTL - 10
    try:
        for d in (busy, idle, other):
            os.utime(d, (old, old))
        app_module.sweep()
        assert busy.is_dir(), "TTL 删除了正在渲染的任务"
        assert not idle.exists(), "过期的空闲任务没有清理"
        assert other.is_dir(), "非任务目录不能被 TTL 按年龄误删"
    finally:
        lock.release()
        with app_module._JOB_LOCKS_GUARD:
            app_module._JOB_LOCKS.pop(busy.name, None)


def test_expired_direct_download_is_not_served(app_module, light_job, client):
    d = light_job
    _meta(app_module, d)
    (d / "manuscript.tjyj.docx").write_bytes(b"private")
    old = time.time() - app_module.JOB_TTL - 10
    os.utime(d, (old, old))

    assert client.get(f"/job/{d.name}/download/docx").status_code == 404
    assert not d.exists(), "空闲过期任务应在访问时清理"


def test_failed_delete_does_not_report_success(app_module, light_job, client, monkeypatch):
    d = light_job
    _meta(app_module, d)

    def occupied(*args, **kwargs):
        raise PermissionError(13, "文件仍被占用")

    monkeypatch.setattr(app_module.shutil, "rmtree", occupied)
    r = client.post(f"/job/{d.name}/delete")
    assert r.status_code == 409
    assert "暂未删除" in r.get_data(as_text=True)
    assert d.is_dir()


# ---------------------------------------------------------------- 路径与损坏现场恢复


@pytest.mark.parametrize("bad_name", ["../outside.docx", "C:/outside.docx", "..\\outside.docx"])
def test_download_rejects_meta_path_escape(app_module, light_job, client, bad_name):
    d = light_job
    _meta(app_module, d, docx=bad_name)
    assert client.get(f"/job/{d.name}/download/docx").status_code == 404


def test_retry_reparses_truncated_ir(app_module, light_job, sample_docx, client, monkeypatch):
    d = light_job
    shutil.copy(sample_docx, d / "manuscript.docx")
    (d / "manuscript.ir.json").write_text('{"source":', encoding="utf-8")
    _meta(app_module, d, stage="uploaded", docx="")
    calls = []

    def fake_parse(dd, jid, original):
        calls.append((jid, original))
        app_module.write_json_atomic(dd / "manuscript.ir.json", {"source": original})

    monkeypatch.setattr(app_module, "stage_parse", fake_parse)
    r = client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"})

    assert r.status_code == 302 and "/details" in r.headers["Location"]
    assert calls == [("cjyj", "source.docx")]
    assert json.loads((d / "manuscript.ir.json").read_text(encoding="utf-8")) == {
        "source": "source.docx"}


# ---------------------------------------------------------------- 提交窗口回滚


def test_commit_failure_restores_same_journal_deliverables(
        app_module, delivered_job, client, monkeypatch):
    """新 DOCX 已替换、审计替换再失败：必须恢复旧 DOCX，而不是留下跨轮组合。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    audit_before = (d / "manuscript.tjyj.audit.json").read_bytes()
    meta_before = (d / "meta.json").read_bytes()
    real_replace = app_module.os.replace

    def fail_on_audit(src, dst):
        if Path(dst).name == "manuscript.tjyj.audit.json":
            raise PermissionError(13, "审计文件被占用")
        return real_replace(src, dst)

    monkeypatch.setattr(app_module.os, "replace", fail_on_audit)
    monkeypatch.setattr(app_module.time, "sleep", lambda _seconds: None)
    r = client.post(f"/job/{d.name}/retry", data={"journal": "tjyj"})

    assert r.status_code == 200, "提交失败应走可读错误页"
    assert (d / "manuscript.tjyj.docx").read_bytes() == docx_before
    assert (d / "manuscript.tjyj.audit.json").read_bytes() == audit_before
    assert (d / "meta.json").read_bytes() == meta_before
    assert not list(d.glob(".rollback.*")), "回滚成功后不能遗留旧稿备份"
    assert client.get(f"/job/{d.name}/download/docx").data == docx_before


def test_switch_journal_commit_failure_keeps_old_meta_and_product(
        app_module, delivered_job, client, monkeypatch):
    """换刊成品落位失败时，meta 不得提前指向新刊。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    meta_before = (d / "meta.json").read_bytes()
    real_replace = app_module.os.replace

    def fail_on_new_docx(src, dst):
        if Path(dst).name == "manuscript.cjyj.docx":
            raise PermissionError(13, "新成品被占用")
        return real_replace(src, dst)

    monkeypatch.setattr(app_module.os, "replace", fail_on_new_docx)
    monkeypatch.setattr(app_module.time, "sleep", lambda _seconds: None)
    r = client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"})

    assert r.status_code == 200
    assert (d / "meta.json").read_bytes() == meta_before
    assert (d / "manuscript.tjyj.docx").read_bytes() == docx_before
    assert not (d / "manuscript.cjyj.docx").exists()
    assert not list(d.glob(".rollback.*"))


# ---------------------------------------------------------------- 规则集发布


def _publish_fixture(app_module, tmp_path, monkeypatch):
    rules_dir = tmp_path / "journal"
    drafts = tmp_path / "drafts"
    draft = drafts / "testjournal"
    rules_dir.mkdir(parents=True)
    draft.mkdir(parents=True)
    rules = json.loads((RULES / "tjyj.json").read_text(encoding="utf-8"))
    rules["journal"]["id"] = "testjournal"
    app_module.write_json_atomic(draft / "draft.json", rules)
    app_module.write_json_atomic(draft / "report.json", {
        "ok_to_publish": True, "errors": [],
        "stats": {"extracted": 0, "unverified": 0, "inherited": 0, "dropped": 0}})
    monkeypatch.setattr(app_module, "RULES_DIR", rules_dir)
    monkeypatch.setattr(app_module, "DRAFT_DIR", drafts)
    return rules_dir, draft


def test_publish_build_failure_keeps_formal_rules_and_draft(
        app_module, client, tmp_path, monkeypatch):
    rules_dir, draft = _publish_fixture(app_module, tmp_path, monkeypatch)
    target = rules_dir / "testjournal.json"
    target.write_bytes(b'{"old": true}')

    def fail_build(_args):
        raise app_module.ConvertError("规格页构建失败", "测试注入")

    monkeypatch.setattr(app_module, "run", fail_build)
    r = client.post("/journals/draft/testjournal/publish",
                    data={"confirm": "on", "overwrite": "on"})

    assert r.status_code == 500
    assert "正式规则库未修改" in r.get_data(as_text=True)
    assert target.read_bytes() == b'{"old": true}'
    assert (draft / "draft.json").is_file(), "失败后不能删掉可恢复的草稿"


def test_successful_publish_updates_artifacts_and_removes_stale_pdf(
        app_module, client, tmp_path, monkeypatch):
    rules_dir, draft = _publish_fixture(app_module, tmp_path, monkeypatch)
    (rules_dir / "testjournal.sample.pdf").write_bytes(b"old pdf")

    def fake_build(args):
        # build_journal 以草稿 JSON 的父目录为输出目录；这里只造两个必需产物，
        # 不生成 PDF，模拟未安装 LibreOffice 的合法路径。
        assert str(draft / "draft.json") in [str(x) for x in args]
        (draft / "testjournal.sample.docx").write_bytes(b"new sample")
        (draft / "testjournal.spec.html").write_text("new spec", encoding="utf-8")
        return ""

    monkeypatch.setattr(app_module, "run", fake_build)
    r = client.post("/journals/draft/testjournal/publish", data={"confirm": "on"})

    assert r.status_code == 302 and "/journal/testjournal/spec" in r.headers["Location"]
    assert (rules_dir / "testjournal.sample.docx").read_bytes() == b"new sample"
    assert (rules_dir / "testjournal.spec.html").read_text(encoding="utf-8") == "new spec"
    assert not (rules_dir / "testjournal.sample.pdf").exists(), "旧 PDF 不得冒充本轮预览"
    assert not draft.exists(), "全部成功后才清理草稿"
    assert json.loads((rules_dir / "testjournal.json").read_text(
        encoding="utf-8"))["journal"]["id"] == "testjournal"


def test_stage_parse_respects_job_lock(app_module, light_job):
    """解析与删除/渲染使用同一把锁；忙碌时不得启动另一个解析子进程。"""
    d = light_job
    lock = app_module.job_lock(d.name)
    assert lock.acquire(blocking=False)
    try:
        with pytest.raises(app_module.RenderBusy):
            app_module.stage_parse(d, "tjyj", "source.docx")
    finally:
        lock.release()


def test_download_and_report_preview_are_locked_during_commit(
        app_module, light_job, client):
    d = light_job
    _meta(app_module, d)
    (d / "manuscript.tjyj.docx").write_bytes(b"docx")
    (d / "audit.html").write_text("audit", encoding="utf-8")
    lock = app_module.job_lock(d.name)
    assert lock.acquire(blocking=False)
    try:
        download = client.get(f"/job/{d.name}/download/docx")
        preview = client.get(f"/job/{d.name}/audit.html")
    finally:
        lock.release()
    assert download.status_code == 409 and preview.status_code == 409
    assert "任务正在更新" in download.get_data(as_text=True)


def test_rollback_failure_preserves_only_old_docx_backup(
        app_module, delivered_job, client, monkeypatch):
    """审计落位失败且旧 DOCX 恢复也失败时，唯一旧稿备份必须仍在。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    real_replace = app_module.os.replace

    def fail_commit_and_rollback(src, dst):
        src_name, dst_name = Path(src).name, Path(dst).name
        if dst_name == "manuscript.tjyj.audit.json":
            raise PermissionError(13, "审计落位失败")
        if dst_name == "manuscript.tjyj.docx" and src_name.startswith(".rollback."):
            raise PermissionError(13, "旧稿恢复也失败")
        return real_replace(src, dst)

    monkeypatch.setattr(app_module.os, "replace", fail_commit_and_rollback)
    monkeypatch.setattr(app_module.time, "sleep", lambda _seconds: None)
    r = client.post(f"/job/{d.name}/retry", data={"journal": "tjyj"})

    assert r.status_code == 200
    backups = list(d.glob(".rollback.*"))
    assert backups, "回滚失败后不能删除唯一旧稿备份"
    assert any(p.read_bytes() == docx_before for p in backups)
    assert "备份保留" in r.get_data(as_text=True)


def test_concurrent_publish_is_rejected_before_touching_draft(
        app_module, client, tmp_path, monkeypatch):
    rules_dir, draft = _publish_fixture(app_module, tmp_path, monkeypatch)
    before = (draft / "draft.json").read_bytes()
    lock = app_module.job_lock("journal:testjournal")
    assert lock.acquire(blocking=False)
    try:
        r = client.post("/journals/draft/testjournal/publish", data={"confirm": "on"})
    finally:
        lock.release()

    assert r.status_code == 409
    assert (draft / "draft.json").read_bytes() == before
    assert not (rules_dir / "testjournal.json").exists()
