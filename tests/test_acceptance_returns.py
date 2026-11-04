"""验收层 C：返回、取消、删除。

这些路径最容易被「功能能跑通」掩盖：按钮在不在、点完回到哪、忙的时候
会不会假装成功。全部从 HTTP 入口验，不靠模板字符串猜。
"""

import pytest

from acceptance_support import (
    build_clean_docx,
    deliver,
    make_draft,
    upload,
)


pytestmark = pytest.mark.acceptance


def _only_job(work):
    dirs = [d for d in work.iterdir() if d.is_dir()]
    assert len(dirs) == 1, f"应当恰好一个任务目录，实际 {len(dirs)}"
    return dirs[0]


def test_failed_page_offers_workspace_and_home(acceptance_client, acceptance_env):
    """C-01/C-02：失败页给「回到工作台」；404 错误页另给「返回首页」。"""
    client, work = acceptance_client, acceptance_env["work"]
    response = upload(client, b"this is not a zip", "坏稿.docx")
    assert response.status_code == 200
    job = _only_job(work)
    assert (job / "error.json").is_file()

    page = client.get(f"/job/{job.name}/failed")
    assert page.status_code == 200
    html = page.get_data(as_text=True)
    assert 'href="/workspace"' in html
    assert "回到工作台" in html

    not_found = client.get("/this-page-does-not-exist")
    assert not_found.status_code == 404
    error_html = not_found.get_data(as_text=True)
    assert 'href="/workspace"' in error_html and 'href="/"' in error_html
    assert "回到工作台" in error_html and "返回首页" in error_html


def test_failed_retry_keeps_the_manuscript_for_reupload_or_retry(
        acceptance_client, acceptance_env):
    """C-03：失败页「再试一次」不要求重新上传；稿件文件仍在现场。"""
    client, work = acceptance_client, acceptance_env["work"]
    upload(client, b"this is not a zip", "坏稿.docx")
    job = _only_job(work)

    response = client.post(f"/job/{job.name}/retry", data={"journal": "tjyj"})
    assert response.status_code == 200
    assert "这份文件打不开" in response.get_data(as_text=True)
    assert (job / "manuscript.docx").is_file(), "重试失败后原稿被删了"


def test_diff_page_links_back_to_result(acceptance_client, acceptance_env, sample_docx):
    """C-05：差异页的「返回结果」指向原任务，不是工作台或空链接。"""
    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx)
    page = client.get(f"/job/{job}/diff?to=cjyj")
    assert page.status_code == 200
    html = page.get_data(as_text=True)
    assert f'href="/result/{job}"' in html
    assert "返回结果" in html


def test_draft_page_links_back_to_library_without_deleting(
        acceptance_client, acceptance_env, tmp_rules):
    """C-06：草稿页「返回规则库」只是链接，草稿必须还在。"""
    client = acceptance_client
    jid, draft = make_draft(tmp_rules)

    page = client.get(f"/journals/draft/{jid}")
    assert page.status_code == 200
    html = page.get_data(as_text=True)
    assert 'href="/journals"' in html
    assert "返回规则库" in html
    assert draft.is_dir(), "打开草稿页把草稿删了"


def test_new_journal_page_links_back_without_side_effects(acceptance_client):
    """C-07：新建期刊页「返回规则库」不产生草稿、不改配置。"""
    client = acceptance_client
    page = client.get("/journals/new")
    assert page.status_code == 200
    html = page.get_data(as_text=True)
    assert 'href="/journals"' in html
    assert "返回规则库" in html


def test_result_page_details_link_keeps_entered_values(
        acceptance_client, acceptance_env, tmp_path):
    """C-11：结果页「补充/修改作者信息」回到补录页，已填值仍在。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_clean_docx(tmp_path / "clean.docx")
    funding = "国家自然科学基金（72003012）"
    job = deliver(client, source, details={"funding": funding})

    result = client.get(f"/result/{job}").get_data(as_text=True)
    assert f'href="/job/{job}/details"' in result

    details = client.get(f"/job/{job}/details")
    assert details.status_code == 200
    assert funding in details.get_data(as_text=True)


def test_illegal_job_and_journal_ids_are_404(acceptance_client):
    """C-12：非法/不存在的 job、journal 一律 404，不进入文件系统路径。"""
    client = acceptance_client
    for path in (
        "/job/not-a-job/details",
        "/job/000000000000/details",
        "/journal/..%5Csecret/spec",
        "/journal/..%5Csecret/sample",
        "/result/not-a-job",
    ):
        assert client.get(path).status_code == 404, path
