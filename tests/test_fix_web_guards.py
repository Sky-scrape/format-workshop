"""F07 / F08 / F36 三个已确认问题的回归钉子，用例即审计清单里的复现方法。

- F07  全部 POST 路由无 Host / Sec-Fetch-Site 校验：浏览器里任意网页可跨站
  表单 POST /settings 改 AI 端点。→ before_request 双闸（Host 本机 + 跨站 403）。
- F08  IR 的 idx 只数 w:p/w:tbl，消费端按 body 全部子元素定位：body 里出现
  位于块元素之前的书签等异类子元素后整体错位。→ 生产端改 dense 计数。
- F36  失败的重渲染先删光上一轮成品与审计：一次失败的换刊/驳回会毁掉
  已交付的结果页。→ 临时名渲染，全部成功后 os.replace 落位。
"""

import json
import shutil
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from conftest import RULES, body_texts, run_tool

import app


# ---------------------------------------------------------------- F07 守卫

@pytest.fixture()
def app_module():
    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


@pytest.fixture()
def ai_cfg_path(app_module, tmp_path, monkeypatch):
    """把 AI 配置指到临时文件，测试既不读也不写仓库里的 config/ai.json。"""
    p = tmp_path / "ai.json"
    monkeypatch.setattr(app_module.ai_assist, "CONFIG_PATH", p)
    return p


def _evil_settings():
    return {"base_url": "https://evil.example/v1", "model": "steal",
            "enabled": "on", "api_key": "k", "send_mode": "full"}


def test_cross_site_post_to_settings_is_403(client, ai_cfg_path):
    """复现 F07：第三方页面的跨站表单 POST /settings 必须被 403 挡下。"""
    r = client.post("/settings", data=_evil_settings(),
                    headers={"Sec-Fetch-Site": "cross-site"})
    assert r.status_code == 403
    assert not ai_cfg_path.exists(), "跨站请求竟然改写了 AI 配置"


def test_cross_site_post_to_convert_is_403(client, tmp_path):
    """跨站表单同样打不到 /convert 这类其余 POST 路由。"""
    r = client.post("/convert", data={"journal": "tjyj"},
                    headers={"Sec-Fetch-Site": "cross-site"})
    assert r.status_code == 403


def test_spoofed_host_is_rejected_even_on_get(client):
    """复现 F07①：Host 不是本机（DNS rebinding 的标志）一律 403，GET 也一样。"""
    r = client.get("/", headers={"Host": "attacker.example.com"})
    assert r.status_code == 403


@pytest.mark.parametrize("host", ["127.0.0.1:5000", "localhost", "localhost:5000",
                                  "[::1]:5000"])
def test_local_hosts_are_allowed(client, host):
    """本机各种 Host 写法都要照常放行，别把正常使用挡在门外。"""
    assert client.get("/", headers={"Host": host}).status_code == 200


def test_post_without_fetch_metadata_still_reaches_handler(client):
    """curl / 本机脚本没有 Sec-Fetch-Site 头，不受影响（仍是业务校验的 400）。"""
    r = client.post("/convert", data={})
    assert r.status_code == 400, "无跨站头的本机客户端不该被 403"

    r = client.post("/convert", data={"journal": "tjyj"},
                    headers={"Sec-Fetch-Site": "same-origin"})
    assert r.status_code == 400, "same-origin 的正常提交不该被 403"


def test_same_origin_settings_post_still_saves(client, ai_cfg_path):
    """守卫只拦跨站：本站自己的设置表单要能正常保存。"""
    r = client.post("/settings", data=_evil_settings(),
                    headers={"Sec-Fetch-Site": "same-origin"})
    assert r.status_code == 200
    assert ai_cfg_path.exists(), "same-origin 的保存没有落到配置文件"


# ---------------------------------------------------------------- F08 块序对位

def _doc_with_leading_bookmark(path):
    """body = [p, bookmarkStart, p, p]：书签位于块元素之前，正是 F08 的错位源。"""
    doc = Document()
    texts = ["这一段用于核验块序对位的第一个段落。",
             "这一段用于核验块序对位的第二个段落。",
             "这一段用于核验块序对位的第三个段落。"]
    for t in texts:
        doc.add_paragraph(t)
    bm = OxmlElement("w:bookmarkStart")
    bm.set(qn("w:id"), "9")
    bm.set(qn("w:name"), "_align_probe")
    doc.paragraphs[0]._p.addnext(bm)
    doc.save(path)
    return texts


def test_ir_idx_is_dense_over_body_children(tmp_path):
    """IR 的 idx 必须数 body 全部子元素：书签也要占一个号，消费端才不会错位。"""
    src = tmp_path / "manuscript.docx"
    texts = _doc_with_leading_bookmark(src)
    run_tool("parse_docx.py", src, "-o", tmp_path / "manuscript.ir.json")
    ir = json.loads((tmp_path / "manuscript.ir.json").read_text(encoding="utf-8"))

    idx_of = {b["text"]: b["idx"] for b in ir["blocks"] if b.get("text")}
    assert [idx_of[t] for t in texts] == [0, 2, 3], \
        f"idx 与 body 子元素位置不一致：{idx_of}"


def test_render_reaches_blocks_behind_bookmark(tmp_path):
    """parse → render 后文字与套版审计仍对位：书签后的段落不能被静默丢掉。"""
    src = tmp_path / "manuscript.docx"
    texts = _doc_with_leading_bookmark(src)
    run_tool("parse_docx.py", src, "-o", tmp_path / "manuscript.ir.json")
    out = tmp_path / "out.docx"
    run_tool("render_from_ir.py", tmp_path / "manuscript.ir.json",
             RULES / "tjyj.json", "-o", out)

    out_texts = body_texts(out)
    assert out_texts.get(2) == texts[1] and out_texts.get(3) == texts[2], \
        "书签之后的段落文字错位或丢失"

    # 套版触达：书签后的两段都被前置区套版改过对齐方式。
    # 旧代码下第 3 段（body 位置 3）拿不到元素，对齐保持 None。
    doc = Document(str(out))
    children = list(doc.element.body)
    # 三个段落 + 书签 + 结尾 sectPr：渲染不许丢掉书签、也不许增删段落
    assert len(children) == 5, "渲染不该丢掉或新增 body 子元素（书签要原样保留）"
    assert children[1].tag == qn("w:bookmarkStart")
    from docx.text.paragraph import Paragraph
    for pos in (2, 3):
        p = Paragraph(children[pos], doc)
        assert p.alignment is not None, f"body 位置 {pos} 的段落没被套版（消费端错位）"

    audit = json.loads((tmp_path / "out.audit.json").read_text(encoding="utf-8"))
    styled = {it["idx"] for it in audit["items"] if "套版" in it["action"]}
    assert styled == {0, 2, 3}, f"审计条目与块位置没对上：{sorted(styled)}"


# ---------------------------------------------------------------- F36 原子落位

@pytest.fixture()
def delivered_job(app_module, sample_docx, tmp_path, monkeypatch):
    """一个已交付过的任务：真实跑完 parse → render，成品与审计都在。"""
    work = tmp_path / "work"
    monkeypatch.setattr(app_module, "WORK", work)
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


def _break_ir(d):
    """把 IR 破坏成必败渲染：render_from_ir 一打开就崩，模拟一次失败的换刊/驳回。"""
    (d / "manuscript.ir.json").write_text("{ 这不是 JSON", encoding="utf-8")


def test_failed_rerender_keeps_deliverables(app_module, client, delivered_job, monkeypatch):
    """复现 F36：重渲染失败后，旧成品与审计必须原封不动，结果页仍可下载。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    audit_before = (d / "manuscript.tjyj.audit.json").read_bytes()

    def fail_render(_args):
        # 不能再靠破坏 IR 注入失败：details 现在会先识别截断 IR 并自动重解析。
        raise app_module.ConvertError("渲染失败", "测试注入的故障")

    monkeypatch.setattr(app_module, "run", fail_render)
    r = client.post(f"/job/{d.name}/details", data={})
    assert r.status_code == 200, "失败要走错误页，而不是 500"

    assert (d / "manuscript.tjyj.docx").read_bytes() == docx_before, "旧成品被删了"
    assert (d / "manuscript.tjyj.audit.json").read_bytes() == audit_before, "旧审计被删了"

    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta.get("docx") == "manuscript.tjyj.docx"
    assert meta.get("journal") == "tjyj"

    page = client.get(f"/result/{d.name}")
    assert page.status_code == 200, "已交付的结果页不能悬空"
    dl = client.get(f"/job/{d.name}/download/docx")
    assert dl.status_code == 200 and dl.data == docx_before, "成品下载不能 404"


def test_failed_retry_leaves_meta_on_old_journal(client, delivered_job):
    """复现 F36⑤：换刊失败的 retry 不得提前把 meta 的期刊写成新刊。"""
    d = delivered_job
    docx_before = (d / "manuscript.tjyj.docx").read_bytes()
    _break_ir(d)

    r = client.post(f"/job/{d.name}/retry", data={"journal": "cjyj"})
    assert r.status_code == 200

    meta = json.loads((d / "meta.json").read_text(encoding="utf-8"))
    assert meta.get("journal") == "tjyj", "失败的换刊把 meta 改成了新刊"
    assert (d / "manuscript.tjyj.docx").read_bytes() == docx_before
    assert not (d / "manuscript.cjyj.docx").exists(), "半成品不能落在正式文件名上"


def test_successful_rerender_replaces_deliverable(client, delivered_job):
    """正常重渲染仍要落位到正式文件名，且不留 .staging 半成品。"""
    d = delivered_job
    r = client.post(f"/job/{d.name}/details", data={})
    assert r.status_code == 302, "渲染成功应跳回结果页"
    assert (d / "manuscript.tjyj.docx").is_file()
    assert not list(d.glob(".staging.*")), "落位后不许残留临时文件"
    assert client.get(f"/job/{d.name}/download/docx").status_code == 200
