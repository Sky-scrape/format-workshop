"""验收层 D/E：空数据与大量数据。

空数据看的是「别 500、别假装成功」；大量数据看的是「别崩、别丢对象」。
真实 80MB 文件与 24 小时真实过期仍属于手动验收，这里用小上限与 mtime 模拟。
"""

import json

import pytest
from docx import Document

from acceptance_support import (
    build_clean_docx,
    build_empty_docx,
    build_large_docx,
    build_title_only_docx,
    deliver,
    upload,
)
from conftest import invariant_failures, render_into


pytestmark = pytest.mark.acceptance


def _only_job(work):
    dirs = [d for d in work.iterdir() if d.is_dir()]
    assert len(dirs) == 1, f"应当恰好一个任务目录，实际 {len(dirs)}"
    return dirs[0]


# ---------------------------------------------------------------- 空数据

def test_zero_byte_upload_lands_on_friendly_error(acceptance_client, acceptance_env):
    """D-01：0 字节文件给可读错误页，不产生可下载的成品。"""
    client, work = acceptance_client, acceptance_env["work"]
    response = upload(client, b"", "空稿.docx")
    assert response.status_code == 200
    assert "这份文件打不开" in response.get_data(as_text=True)

    job = _only_job(work)
    assert (job / "error.json").is_file()
    assert not list(job.glob("manuscript.*.docx")), "空文件却产出了成品"


def test_missing_file_or_journal_is_400(acceptance_client):
    """D-02：缺文件、缺期刊都 400，且报错页仍能继续选刊。"""
    client = acceptance_client

    no_file = client.post("/convert", data={"journal": "tjyj"})
    assert no_file.status_code == 400
    html = no_file.get_data(as_text=True)
    assert "请选择一份 DOCX 稿件" in html
    assert 'name="journal"' in html

    no_journal = upload(client, b"PK anything", "稿.docx", journal="")
    assert no_journal.status_code == 400
    assert "请选择一本目标期刊" in no_journal.get_data(as_text=True)


def test_empty_docx_is_refused_before_any_delivery(
        acceptance_client, acceptance_env, tmp_path):
    """D-03：0 段 docx 在解析阶段就报错，不产出 0 块成品。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_empty_docx(tmp_path / "empty.docx")
    response = upload(client, source, "空稿.docx")

    assert response.status_code == 200
    assert "没有可排版的内容" in response.get_data(as_text=True)
    job = _only_job(work)
    assert (job / "error.json").is_file()
    assert not list(job.glob("manuscript.*.docx")), "空稿却产出了成品"
    assert not (job / "meta.json").exists() or \
        json.loads((job / "meta.json").read_text(encoding="utf-8")).get("stage") != "rendered"


def test_title_only_manuscript_is_usable(acceptance_client, acceptance_env, tmp_path):
    """D-04/D-05：只有标题的稿子不崩；缺失项在报告里如实列出。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_title_only_docx(tmp_path / "title.docx")
    job = deliver(client, source)

    d = work / job
    assert (d / "triage.html").is_file()
    triage = (d / "triage.html").read_text(encoding="utf-8")
    assert "缺" in triage, "标题-only 稿件的缺失项没有进报告"
    assert client.get(f"/result/{job}").status_code == 200


def test_empty_history_and_empty_draft_list(acceptance_client):
    """D-07/D-08：没有历史、没有草稿时给空态，不 500。"""
    client = acceptance_client
    history = client.get("/history")
    assert history.status_code == 200
    assert "还没有转换记录" in history.get_data(as_text=True)

    journals = client.get("/journals")
    assert journals.status_code == 200
    assert "草稿" not in journals.get_data(as_text=True) or "没有" in journals.get_data(as_text=True)


def test_query_page_ignores_invalid_and_duplicate_ids(acceptance_client):
    """D-09：查询页不因非法 id、重复 id 或超过 3 本而 500。"""
    client = acceptance_client
    response = client.get("/query?journals=cjyj,tjyj,gjjryj,cjyj,zzz,..%5Csecret")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert "统计研究" in html
    assert "财经研究" in html


def test_empty_journal_library_does_not_500(acceptance_client, monkeypatch):
    """D-11：规则库为空时，首页/工作台/期刊页/查询页仍可打开。"""
    import app

    monkeypatch.setattr(app, "load_journals", lambda: [])
    client = acceptance_client
    for path in ("/", "/workspace", "/journals", "/query"):
        response = client.get(path)
        assert response.status_code == 200, path


def test_missing_download_is_404(acceptance_client, acceptance_env, sample_docx):
    """D-12：不存在的下载类别/产物一律 404，不能把目录里别的文件带出去。"""
    from acceptance_support import deliver

    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    assert client.get(f"/job/{job}/download/nope").status_code == 404
    (work / job / "audit.html").unlink()
    assert client.get(f"/job/{job}/download/audit").status_code == 404


# ---------------------------------------------------------------- 大量数据

@pytest.mark.slow
def test_large_manuscript_keeps_tables_figures_and_text(tmp_path):
    """E-01：300+ 段、12 表、60 文献、12 标题；对象不丢、文字守恒。"""
    source = build_large_docx(tmp_path / "large.docx")
    pipeline = render_into(tmp_path / "large-out", source, target="tjyj")

    doc = Document(str(pipeline.docx))
    assert len(doc.tables) == 12, f"表格数量变了：{len(doc.tables)}"
    assert len(doc.inline_shapes) == 3, f"图片数量变了：{len(doc.inline_shapes)}"
    refs = [t for t in pipeline.output_texts.values()
            if t.strip().startswith("[") and "]" in t]
    assert len(refs) >= 60, f"参考文献条目丢了：{len(refs)}"
    assert not invariant_failures(pipeline)


def test_history_shows_at_most_thirty_rows(acceptance_client, acceptance_env):
    """E-06：历史页只列最近 30 条，不把整目录铺出来。"""
    client, work = acceptance_client, acceptance_env["work"]
    for i in range(31):
        d = work / f"{i:012x}"
        d.mkdir()
        (d / "meta.json").write_text(json.dumps({
            "job": d.name, "original": f"稿件{i}.docx", "journal": "tjyj",
            "journal_name": "统计研究", "stage": "details", "created": "2026-11-04 10:00",
        }, ensure_ascii=False), encoding="utf-8")

    page = client.get("/history")
    assert page.status_code == 200
    assert page.get_data(as_text=True).count('<td class="name"') == 30


def test_long_chinese_filename_download_is_safe(acceptance_client, acceptance_env,
                                                tmp_path):
    """E-09：中文/空格/长文件名可上传、可下载；下载头带 RFC 5987 编码。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_clean_docx(tmp_path / "clean.docx")
    name = "这是一个很长的稿件名字 包含空格" * 3 + ".docx"
    job = deliver(client, source, name=name)

    response = client.get(f"/job/{job}/download/docx")
    assert response.status_code == 200
    disposition = response.headers["Content-Disposition"]
    assert "filename*=UTF-8''" in disposition
    assert "%" in disposition  # 中文名已编码
    response.close()


@pytest.mark.slow
def test_long_paragraph_and_unbroken_english_are_preserved(tmp_path):
    """E-10：超长段与无空格英文串不崩、不截断。"""
    from docx import Document as _Doc

    source = tmp_path / "long-text.docx"
    doc = _Doc()
    doc.add_paragraph("数字金融与企业融资约束")
    long_cn = "本文考察数字金融对企业融资约束的影响。" * 400
    long_en = "A" * 5000
    doc.add_paragraph(long_cn)
    doc.add_paragraph(long_en)
    doc.save(str(source))

    pipeline = render_into(tmp_path / "long-text-out", source, target="tjyj")
    assert long_cn in pipeline.output_texts.values()
    assert long_en in pipeline.output_texts.values()
