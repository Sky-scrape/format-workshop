"""验收层 A：一条主流程，从首页走到删除。

这些用例只从用户真正会点的入口进入（HTTP 路由 + 表单），底层仍走真实流水线，
不把 render 换成 mock——mock 只能证明路由会跳转，证明不了成品真的出来了。
"""

import json
import os
import time

import pytest

from acceptance_support import (
    audit_of,
    build_clean_docx,
    deliver,
    details_of,
    docx_contains,
    meta_of,
    para_texts,
    upload,
    upload_job,
)
from conftest import (
    Pipeline,
    RULES,
    change_items_without_rule,
    invariant_failures,
    layout_signature,
)
from rules_ref import resolvable


pytestmark = pytest.mark.acceptance


def _rules(journal="tjyj"):
    return json.loads((RULES / f"{journal}.json").read_text(encoding="utf-8"))


def _pipeline(work, job, journal="tjyj"):
    d = work / job
    audit = audit_of(work, job, journal)
    ir = json.loads((d / "manuscript.ir.json").read_text(encoding="utf-8"))
    rules = _rules(journal)
    kw_sep = next((b.get("separator") for b in rules.get("blocks", [])
                   if b.get("type") == "keywords" and b.get("lang") == "zh"), None)
    labels = {b["type"]: b["label"] for b in rules.get("blocks", [])
              if b.get("lang") == "zh" and b.get("type") in ("abstract", "keywords")
              and b.get("label")}
    return Pipeline(d, d / f"manuscript.{journal}.docx", audit, ir,
                    keyword_sep=kw_sep, labels=labels)


def test_end_to_end_upload_details_result_download_history_delete(
        acceptance_client, acceptance_env, sample_docx):
    """A-00：首页 → 工作台 → 上传 → 补录 → 结果 → 下载 → 历史 → 删除。"""
    client, work = acceptance_client, acceptance_env["work"]

    assert client.get("/").status_code == 200
    workspace = client.get("/workspace")
    assert workspace.status_code == 200
    assert 'name="journal"' in workspace.get_data(as_text=True)

    job = deliver(client, sample_docx)
    d = work / job
    assert (d / "manuscript.docx").is_file()
    assert (d / "manuscript.ir.json").is_file()
    assert (d / "triage.html").is_file()
    assert meta_of(work, job)["stage"] == "rendered"

    result = client.get(f"/result/{job}")
    assert result.status_code == 200
    html = result.get_data(as_text=True)
    for kind in ("docx", "docx-anon", "triage", "audit"):
        assert f"/job/{job}/download/{kind}" in html, f"结果页缺 {kind} 下载入口"

    # 四份交付真的能下载，且不是空文件
    for kind in ("docx", "docx-anon", "triage", "audit"):
        response = client.get(f"/job/{job}/download/{kind}")
        assert response.status_code == 200, kind
        assert len(response.data) > 100, f"{kind} 下载内容为空"
        assert response.headers["Content-Disposition"].startswith("attachment;")
        response.close()

    # 审计存在、条目可核、每条改动都有可解析的规则依据
    audit = audit_of(work, job)
    assert audit.get("items"), "交付了成品却没有审计条目"
    rules = _rules("tjyj")
    bad = [it for it in audit["items"]
           if it.get("rule") and not resolvable(rules, it["rule"])]
    assert not bad, f"审计里的 rule 解析不到规则集：{bad[:3]}"
    assert not change_items_without_rule(_pipeline(work, job), rules)

    # 正文四条不变式
    assert not invariant_failures(_pipeline(work, job))

    history = client.get("/history")
    assert history.status_code == 200
    assert "source.docx" in history.get_data(as_text=True)

    deleted = client.post(f"/job/{job}/delete")
    assert deleted.status_code == 302 and deleted.headers["Location"].endswith("/history")
    assert not d.exists(), "删除后任务目录还在"


def test_details_page_prefills_what_the_manuscript_has(
        acceptance_client, acceptance_env, tmp_path):
    """A-03：补录页显示稿件已有信息，并标出该刊缺的要件。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_clean_docx(tmp_path / "clean.docx")
    job = upload_job(client, source)
    page = client.get(f"/job/{job}/details")
    assert page.status_code == 200
    html = page.get_data(as_text=True)

    assert "数字金融与企业融资约束" in html
    assert "内容提要" in html
    assert "该刊必备" in html, "缺项没有标成必备"
    assert 'name="english_abstract"' in html
    assert 'name="funding"' in html
    assert "内容块" in html


def test_blank_details_delivers_without_placeholders(
        acceptance_client, acceptance_env, sample_docx):
    """A-04：补录全空也能出稿；成品里不能出现待填占位。"""
    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    out = work / job / "manuscript.tjyj.docx"

    text = "\n".join(para_texts(out))
    for marker in ("待填", "请填写", "（此处", "【待", "TODO", "XXX"):
        assert marker not in text, f"成品里出现了占位标记：{marker}"
    assert not invariant_failures(_pipeline(work, job))


def test_filled_details_are_injected_and_reported(
        acceptance_client, acceptance_env, tmp_path):
    """A-05：补录内容真的进成品，并在 meta 里留下注入记录。"""
    from docx import Document

    client, work = acceptance_client, acceptance_env["work"]
    source = build_clean_docx(tmp_path / "clean.docx")
    abstract = "This paper studies digital finance and financing constraints."
    funding = "国家自然科学基金青年项目（72003012）"
    job = deliver(client, source, details={
        "english_abstract": abstract,
        "english_title": "Digital Finance and Financing Constraints",
        "affiliation_zh": "某某大学经济学院",
        "funding": funding,
    })
    d = work / job
    saved = details_of(work, job)
    assert saved["english_abstract"] == abstract
    assert saved["funding"] == funding

    meta = meta_of(work, job)
    assert meta.get("injected"), "补录了内容却没有注入记录"
    assert docx_contains(d / "manuscript.tjyj.docx", abstract)
    assert docx_contains(d / "manuscript.tjyj.docx", "某某大学经济学院")

    doc = Document(str(d / "manuscript.tjyj.docx"))
    footer = "\n".join(p.text for p in doc.sections[0].first_page_footer.paragraphs)
    assert funding in footer, f"基金没有进首页脚注：{footer!r}"

    page = client.get(f"/result/{job}").get_data(as_text=True)
    assert "已注入" in page

    # 补录不是格式改动：审计不该把它记成 before/after 文字改动
    audit = audit_of(work, job)
    assert all(abstract not in (it.get("before") or "")
               and abstract not in (it.get("after") or "")
               for it in audit["items"])


def test_anonymized_and_signed_versions_differ_on_identity(
        acceptance_client, acceptance_env, tmp_path):
    """A-08：匿名稿删身份，署名稿保留；两份都过元数据检查。"""
    import docx_meta
    from docx import Document

    client, work = acceptance_client, acceptance_env["work"]
    source = tmp_path / "with_author.docx"
    doc = Document()
    doc.add_paragraph("数字金融与企业融资约束")
    doc.add_paragraph("张三　李四")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("关键词：数字金融；融资约束；企业创新")
    doc.add_paragraph("正文段落，用于确认匿名不会误删正文。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。")
    doc.save(str(source))

    job = deliver(client, source)
    d = work / job
    signed = d / "manuscript.tjyj.docx"
    anon = d / "manuscript.tjyj.anon.docx"
    assert signed.is_file() and anon.is_file()

    assert docx_contains(signed, "张三　李四")
    assert not docx_contains(anon, "张三　李四"), "匿名稿里还留着作者署名行"
    assert docx_contains(anon, "数字金融与企业融资约束"), "标题被误删"
    assert docx_contains(anon, "正文段落"), "正文被误删"
    docx_meta.assert_clean(str(signed))
    docx_meta.assert_clean(str(anon), strict_people=True)

    meta = meta_of(work, job)
    assert meta["anon_removed"] > 0


def test_switch_journal_through_result_page(
        acceptance_client, acceptance_env, sample_docx):
    """A-09：结果页换刊重排，成品、meta、审计一起换到新刊。"""
    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    d = work / job

    response = client.post(f"/job/{job}/retry", data={"journal": "cjyj"})
    assert response.status_code == 302
    assert response.headers["Location"].startswith(f"/result/{job}")

    meta = meta_of(work, job)
    assert meta["journal"] == "cjyj"
    assert meta["docx"] == "manuscript.cjyj.docx"
    assert (d / "manuscript.cjyj.docx").is_file()
    assert not (d / "manuscript.tjyj.docx").exists(), "换刊后旧刊成品没清理"
    assert client.get(f"/result/{job}").status_code == 200


def test_second_render_same_journal_is_stable(
        acceptance_client, acceptance_env, sample_docx):
    """A-12：同一份稿件同刊重排两次，版式与审计形状不变。"""
    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    d = work / job
    before_sig = layout_signature(d / "manuscript.tjyj.docx")
    before_audit = audit_of(work, job)

    response = client.post(f"/job/{job}/retry", data={"journal": "tjyj"})
    assert response.status_code == 302

    after_sig = layout_signature(d / "manuscript.tjyj.docx")
    after_audit = audit_of(work, job)
    assert after_sig == before_sig, "同输入重排后版式指纹变了"
    assert [(i["id"], i["action"]) for i in after_audit["items"]] == \
           [(i["id"], i["action"]) for i in before_audit["items"]]


def test_expired_job_is_not_served(acceptance_client, acceptance_env, sample_docx):
    """A-11：超过 24 小时的任务页不可再访问，目录被清掉。"""
    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    d = work / job
    old = time.time() - 25 * 3600
    os.utime(d, (old, old))

    assert client.get(f"/result/{job}").status_code == 404
    assert not d.exists(), "过期任务没有被清理"


def test_missing_report_blocks_delivery(
        acceptance_client, acceptance_env, sample_docx, monkeypatch):
    """B-12：报告生成失败时，成品不得交付（无报告视为交付失败）。"""
    import app as app_module

    client, work = acceptance_client, acceptance_env["work"]
    job = upload_job(client, sample_docx)
    d = work / job

    real_run = app_module.run

    def fake_run(cmd, *args, **kwargs):
        if any("render_audit.py" in str(part) for part in cmd):
            raise app_module.ConvertError("报告生成失败", "没有审计报告不得交付")
        return real_run(cmd, *args, **kwargs)

    monkeypatch.setattr(app_module, "run", fake_run)
    response = client.post(f"/job/{job}/details", data={})
    assert response.status_code == 200
    assert "报告生成失败" in response.get_data(as_text=True)
    assert not (d / "manuscript.tjyj.docx").exists(), "报告失败却落了成品"
    assert meta_of(work, job).get("stage") != "rendered"


def test_result_page_links_back_to_workspace(acceptance_client, sample_docx):
    """C-04：结果页的「转换另一份稿件」回到工作台。"""
    client = acceptance_client
    job = deliver(client, sample_docx)
    html = client.get(f"/result/{job}").get_data(as_text=True)
    assert 'href="/workspace"' in html
    assert client.get("/workspace").status_code == 200
