"""F37 回归：app.py 路由主链路的端到端用例。

修复前路由层几乎没有测试：上传→details→render→result 的主链路、
reject / retry / download、413 与假 docx / 空文件，全靠人肉点页面验。
本文件把这些入口钉住。三条环境纪律：

1. WORK 指到临时目录，测试不碰仓库里的 outputs/web
   （同 test_fix_web_robust.work 的做法）。
2. 渲染走真实流水线（render_sample 的合成稿），保证「result 可达」
   不是 mock 出来的假绿；只有 details POST 的路由顺序契约用 mock 钉
   （见 test_details_post 先落盘再渲染）。
3. 不测真实 LibreOffice .doc 转换——环境不可依赖。
"""

import io
import json
import shutil
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


def the_only_job(work):
    """刚上传的任务：目录只有一个，名字是 12 位十六进制。"""
    dirs = [d for d in work.iterdir() if d.is_dir()]
    assert len(dirs) == 1, f"应当恰好产生一个任务目录，实际 {len(dirs)} 个"
    return dirs[0]


def upload(client, name, payload, journal="tjyj"):
    return client.post("/convert", data={
        "journal": journal,
        "file": (io.BytesIO(payload), name),
    })


@pytest.fixture()
def delivered_job(app_module, work, sample_docx):
    """一个已交付的任务：真实跑完 parse → render，成品与审计都在
    （参照 test_fix_web_robust.delivered_job）。"""
    job = "f37a1b2c3d4e"
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


# ---------------------------------------------------------------- 上传守卫

def test_fake_docx_lands_on_friendly_error_page(client, work):
    """复现 F37：改了后缀的假 docx 要给「另存为」的友好错误页，不能 500。"""
    r = upload(client, "假稿.docx", b"this is not a zip archive, just text")
    assert r.status_code == 200, "报错也应是可读的错误页，而不是裸 500"
    html = r.get_data(as_text=True)
    assert "这份文件打不开" in html
    assert "另存为" in html, "假 docx 的错误页没给用户可操作的出路"

    d = the_only_job(work)
    err = json.loads((d / "error.json").read_text(encoding="utf-8"))
    assert err["title"] == "这份文件打不开"

    # 失败页可回访，且解析没成功 → 只给「重新上传」的出路（F12 文案契约）
    again = client.get(f"/job/{d.name}/failed")
    assert again.status_code == 200 and "重新上传" in again.get_data(as_text=True)


def test_empty_file_lands_on_friendly_error_page(client, work):
    """复现 F37：0 字节文件同样要友好错误页。"""
    r = upload(client, "空稿.docx", b"")
    assert r.status_code == 200
    assert "这份文件打不开" in r.get_data(as_text=True)
    assert (the_only_job(work) / "error.json").is_file()


def test_upload_guards_pdf_suffix_and_missing_journal(client):
    """复现 F37：非 docx/doc 后缀与未选期刊这两个 400 守卫。"""
    r = upload(client, "论文.pdf", b"%PDF-1.4 fake")
    assert r.status_code == 400
    assert "请向作者要源文件" in r.get_data(as_text=True)

    r = client.post("/convert", data={
        "file": (io.BytesIO(b"PK anything"), "稿.docx")})
    assert r.status_code == 400
    assert "请选择一本目标期刊" in r.get_data(as_text=True)


# ---------------------------------------------------------------- 主链路

def test_real_docx_upload_reaches_details_page(client, work, sample_docx):
    """复现 F37 主链路上半程：真实 docx 上传 → 302 → 补录页可达。"""
    r = upload(client, "source.docx", sample_docx.read_bytes())
    assert r.status_code == 302
    assert "/details" in r.headers["Location"]

    d = the_only_job(work)
    assert (d / "manuscript.ir.json").is_file(), "解析产物 IR 没落盘"
    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta["stage"] == "details" and meta["original"] == "source.docx"

    html = client.get(f"/job/{d.name}/details").get_data(as_text=True)
    assert html.count("fillForm") >= 1, "补录页表单没渲染出来"
    assert 'name="english_abstract"' in html, "英文要件输入框缺失"
    assert "tjyj" in html, "补录页没有标明目标期刊（此时 journal_name 尚未写入 meta）"


def test_details_post_persists_details_then_reaches_result(app_module, client,
                                                           work, sample_docx,
                                                           monkeypatch):
    """复现 F37 主链路下半程：details POST → 渲染（mock）→ result 可达。

    mock 顺带钉一条 F11 的顺序契约：details.json 必须在 stage_render
    之前落盘——持锁的渲染开头就要读到用户刚提交的补录内容。
    """
    r = upload(client, "source.docx", sample_docx.read_bytes())
    assert r.status_code == 302
    job = the_only_job(work).name

    seen = {}

    def fake_render(dd, jid, original):
        saved = json.loads((dd / "details.json").read_text(encoding="utf-8"))
        assert saved.get("english_abstract") == "This paper studies digital finance.", \
            "stage_render 启动时补录内容还没落盘（F11 顺序契约）"
        seen["args"] = (dd.name, jid, original)
        (dd / f"manuscript.{jid}.docx").write_bytes(b"PK fake deliverable")
        meta = app_module.read_meta(dd)
        meta.update({"job": dd.name, "journal": jid, "original": original,
                     "stem": "source", "stage": "rendered",
                     "docx": f"manuscript.{jid}.docx"})
        app_module.write_meta(dd, meta)
        return meta

    monkeypatch.setattr(app_module, "stage_render", fake_render)
    r = client.post(f"/job/{job}/details", data={
        "english_abstract": "This paper studies digital finance.",
        "funding": "国家社科基金（22BJY000）"})
    assert r.status_code == 302 and f"/result/{job}" in r.headers["Location"]
    assert seen["args"] == (job, "tjyj", "source.docx"), "stage_render 参数不对"

    saved = json.loads((work / job / "details.json").read_text(encoding="utf-8"))
    assert saved["english_abstract"] == "This paper studies digital finance."
    assert saved["funding"] == "国家社科基金（22BJY000）", "首页脚注字段没进 details.json"

    page = client.get(f"/result/{job}")
    assert page.status_code == 200
    assert "统计研究" in page.get_data(as_text=True)


# ---------------------------------------------------------------- 驳回与换刊

def test_reject_rerenders_and_marks_item_rejected(client, delivered_job):
    """复现 F37：勾「保留原文」提交 reject → 写 reject.json → 重渲染 →
    新审计把该条记为已驳回，meta 计数与结果页都如实反映。"""
    d = delivered_job
    audit = json.loads((d / "manuscript.tjyj.audit.json").read_text(encoding="utf-8"))
    rejectable = {"标题重新编号", "题注归一", "标签归一", "文献表标题归一",
                  "关键词分隔符归一", "文献条目编号归一",
                  "表注／图注层序调整", "段首缩进空格删除"}
    target = next(it for it in audit["items"]
                  if (it.get("before") or it.get("after")) and not it.get("rejected"))
    assert target["action"] in rejectable, f"选中的条目不可驳回：{target['action']}"

    r = client.post(f"/job/{d.name}/reject", data={"rid": target["id"]})
    assert r.status_code == 302 and f"/result/{d.name}" in r.headers["Location"]

    rej = json.loads((d / "reject.json").read_text(encoding="utf-8"))
    assert rej["ids"] == [target["id"]], "驳回清单没有按提交内容落盘"

    audit2 = json.loads((d / "manuscript.tjyj.audit.json").read_text(encoding="utf-8"))
    marked = next(it for it in audit2["items"] if it["id"] == target["id"])
    assert marked.get("rejected"), "重渲染后该条没有记为已驳回（还是改了原文）"

    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta["rejected"] == 1 and meta["stage"] == "rendered"
    assert "已驳回 1 条改动" in client.get(f"/result/{d.name}").get_data(as_text=True)


def test_retry_switches_journal_and_redelivers(client, delivered_job):
    """复现 F37：已交付任务换刊重排 → 新刊成品落正式文件名，结果页可达。"""
    d = delivered_job
    r = client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"})
    assert r.status_code == 302 and f"/result/{d.name}" in r.headers["Location"]

    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta["journal"] == "cjyj" and meta["docx"] == "manuscript.cjyj.docx"
    assert (d / "manuscript.cjyj.docx").is_file(), "换刊成品没有落位"
    assert not list(d.glob(".staging.*")), "落位后不许残留临时文件"
    assert client.get(f"/result/{d.name}").status_code == 200


# ---------------------------------------------------------------- 下载与限流

def test_download_serves_all_three_products(client, delivered_job):
    """复现 F37：成品 / 审计 / 体检三种下载都要可达，且是带可读名的附件；
    未知类别与匿名稿（未生成）404。"""
    d = delivered_job
    docx_bytes = (d / "manuscript.tjyj.docx").read_bytes()
    (d / "audit.html").write_text("<html>审计报告</html>", encoding="utf-8")
    (d / "triage.html").write_text("<html>投稿体检</html>", encoding="utf-8")

    r = client.get(f"/job/{d.name}/download/docx")
    assert r.status_code == 200 and r.data == docx_bytes
    cd = r.headers["Content-Disposition"]
    assert cd.startswith("attachment;") and "manuscript-tjyj-docx.docx" in cd, cd
    assert "filename*=UTF-8''" in cd, "中文名没有走 RFC 5987"

    for kind, fallback in (("audit", "paper-audit.html"), ("triage", "paper-triage.html")):
        r = client.get(f"/job/{d.name}/download/{kind}")
        assert r.status_code == 200, f"{kind} 下载 404"
        cd = r.headers["Content-Disposition"]
        assert cd.startswith("attachment;") and fallback in cd

    assert client.get(f"/job/{d.name}/download/nope").status_code == 404
    assert client.get(f"/job/{d.name}/download/docx-anon").status_code == 404, \
        "匿名稿没生成时不能把 404 之外的动静给出去"


def test_oversized_upload_returns_413_page(app_module, client, work, monkeypatch):
    """复现 F37：MAX_CONTENT_LENGTH 的 413 handler 要给出友好页，
    且请求在进入路由前就被拒——不产生任何任务目录。"""
    monkeypatch.setitem(app_module.app.config, "MAX_CONTENT_LENGTH", 64 * 1024)
    r = upload(client, "大稿.docx", b"\x00" * (128 * 1024))
    assert r.status_code == 413
    html = r.get_data(as_text=True)
    assert "413" in html and "文件超过" in html, "413 没有走友好错误页"
    assert not any(work.iterdir()), "超限请求不该创建任务目录"
