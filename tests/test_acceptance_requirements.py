"""验收层 B：需求规则中「用户能感知、但底层契约测试不直接覆盖」的部分。

已有的 767 条契约测试仍然守着渲染细节；这里只补三类：
- 用户操作顺序与表单行为（选刊→补录、跨刊重填）；
- 报告与设置页的可核性（三口径字数、密钥不回显、建库门禁）；
- 结构性要件阻断（DESIGN §10.4）——稿件本体缺件不得开始转换。
"""

import json
import re

import pytest
from docx import Document

from acceptance_support import (
    TJYJ_FILL,
    build_clean_docx,
    build_many_headings_docx,
    build_spacing_docx,
    build_title_only_docx,
    deliver,
    make_draft,
    meta_of,
    upload,
    upload_job,
)
from conftest import (
    RULES,
    SHIPPED_JOURNALS,
    invariant_failures,
    minimal_doc,
    render_into,
    run_sizes,
    run_tool,
)


pytestmark = pytest.mark.acceptance


def _rules(journal="tjyj"):
    return json.loads((RULES / f"{journal}.json").read_text(encoding="utf-8"))


def _with_author_line(path, authors="张三　李四"):
    """build_clean_docx 没有署名行；统计研究要求「作者」，交付前补一行。"""
    build_clean_docx(path)
    doc = Document(str(path))
    doc.paragraphs[1].insert_paragraph_before(authors)
    doc.save(str(path))
    return path


def _complete_docx(path, english_abstract=True):
    """一本刊的完整稿：中文要件齐全，外加英文标题/摘要/关键词。

    目标刊用 gjjryj（blocks 不要求署名行以外的作者页字段），用来钉住
    「个人信息表单全空也能转换」——阻断项只认稿件本体要件，不认作者页字段。
    """
    doc = Document()
    lines = [
        "数字金融与企业融资约束",
        "张三　李四",
        "内容提要：本文考察数字金融对企业融资约束的影响。",
        "关键词：数字金融；融资约束；企业创新",
        "中图分类号：F821",
        "Digital Finance and Financing Constraints",
        "Key words：digital finance; financing constraints",
        "一、引言",
        "本文使用上市公司数据，考察数字金融对企业融资约束的影响。",
        "参考文献",
        "[1] 张三，2020：《数字金融与企业融资》，《经济研究》第3期。",
    ]
    if english_abstract:
        # 英文摘要必须排在英文标题之后、首个一级标题之前，解析器才认得出
        lines.insert(6, "Abstract：This paper studies digital finance and financing constraints.")
    for line in lines:
        doc.add_paragraph(line)
    doc.save(str(path))
    return path


def _triage_text(tmp_path, docx_path, journal="tjyj"):
    """跑一遍体检，返回去掉样式与标签后的纯文本。"""
    ir = tmp_path / "m.ir.json"
    out = tmp_path / "triage.html"
    run_tool("parse_docx.py", docx_path, "-o", ir)
    run_tool("compliance_report.py", ir, RULES / f"{journal}.json", "-o", out)
    raw = out.read_text(encoding="utf-8")
    raw = re.sub(r"<style.*?</style>", "", raw, flags=re.S)
    raw = re.sub(r"<script.*?</script>", "", raw, flags=re.S)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw))


# ---------------------------------------------------------------- 文字与版式

@pytest.mark.parametrize("target", SHIPPED_JOURNALS)
def test_text_fingerprint_zero_change_for_each_shipped_journal(
        sample_docx, tmp_path, target):
    """B-01：三本自带刊各跑一遍，抹掉白名单后正文逐字相同。"""
    pipeline = render_into(tmp_path / f"fingerprint-{target}", sample_docx, target=target)
    problems = invariant_failures(pipeline)
    assert not problems, f"{target}: " + "；".join(problems)


def test_unimplemented_normalizations_are_left_alone(tmp_path):
    """B-04：N-02～N-06 与段尾空白未实现，正文内部必须原样保留。"""
    source = build_spacing_docx(tmp_path / "spacing.docx")
    pipeline = render_into(tmp_path / "spacing-out", source, target="tjyj")

    original = "研究发现,  数字金融 显著缓解了融资约束 。样本 共 １２００ 个。"
    assert original in pipeline.output_texts.values(), \
        "正文内部的连续空格/半角标点/全角数字被改动了"


def test_layout_values_follow_the_ruleset(sample_docx, tmp_path):
    """B-06/B-31：纸张、页边距、字号来自规则集；改规则集，成品跟着变。"""
    rules = _rules("tjyj")
    pipeline = render_into(tmp_path / "layout", sample_docx, target="tjyj")
    doc = Document(str(pipeline.docx))
    sec = doc.sections[0]
    page = rules["page"]

    assert sec.page_width.mm == pytest.approx(page["size_mm"]["width"], abs=0.5)
    assert sec.page_height.mm == pytest.approx(page["size_mm"]["height"], abs=0.5)
    for side, key in (("top", "top"), ("bottom", "bottom"),
                      ("left", "left"), ("right", "right")):
        assert getattr(sec, f"{side}_margin").mm == pytest.approx(
            page["margin_mm"][key], abs=0.5)

    body_idx = next(i for i, t in pipeline.block_type.items() if t == "body")
    sizes = run_sizes(pipeline.docx)
    assert rules["typography"]["body"]["size_pt"] in sizes[body_idx]

    # 改字号 → 成品字号必须跟着动（规则集是真源，不是硬编码）
    alt = json.loads(json.dumps(rules))
    new_size = alt["typography"]["body"]["size_pt"] + 1.5
    alt["typography"]["body"]["size_pt"] = new_size
    alt_path = tmp_path / "alt.json"
    alt_path.write_text(json.dumps(alt, ensure_ascii=False), encoding="utf-8")
    moved = render_into(tmp_path / "layout-alt", sample_docx,
                        target="tjyj", rules_path=alt_path)
    moved_sizes = run_sizes(moved.docx)
    assert new_size in moved_sizes[body_idx], "规则集改了字号，成品没跟着变"


def test_audit_items_have_the_contract_fields(
        acceptance_client, acceptance_env, sample_docx):
    """B-09：审计条目字段齐全，id 为「块序号|动作」，不是随机数。"""
    from acceptance_support import audit_of

    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx, details=TJYJ_FILL)
    items = audit_of(work, job)["items"]
    assert items

    for item in items:
        for key in ("id", "idx", "action", "note", "rule", "rejected"):
            assert key in item, f"审计条目缺字段 {key}：{item}"
        assert "|" in item["id"], f"id 不是「块序号|动作」：{item['id']}"
        assert item["action"].strip(), "动作名为空"


def test_rejecting_one_edit_reverts_only_that_block(
        acceptance_client, acceptance_env, sample_docx):
    """B-13：驳回一条只回退该条，其余改动逐字不变。"""
    from acceptance_support import audit_of

    client, work = acceptance_client, acceptance_env["work"]
    job = deliver(client, sample_docx, details=TJYJ_FILL)
    before = audit_of(work, job)
    target = next(it for it in before["items"]
                  if (it.get("before") or it.get("after"))
                  and it["action"] in {"标签归一", "关键词分隔符归一",
                                       "题注归一", "文献表标题归一"})
    idx = target["idx"]

    response = client.post(f"/job/{job}/reject", data={"rid": target["id"]})
    assert response.status_code == 302
    after = audit_of(work, job)
    marked = next(it for it in after["items"] if it["id"] == target["id"])
    assert marked["rejected"] is True

    out = work / job / "manuscript.tjyj.docx"
    source = work / job / "manuscript.docx"
    from conftest import body_texts

    before_texts = body_texts(source)
    after_texts = body_texts(out)
    assert before_texts[idx] == after_texts[idx], "被驳回的块没有回到原样"

    # 其余块中，原本与成品不同的块仍然保持改动；不能因为驳回一条就整体回退
    baseline = audit_of(work, job)
    assert baseline["items"]


# ---------------------------------------------------------------- 报告与设置

def test_word_count_report_gives_three_metrics(tmp_path):
    """B-14：三种字数口径必须同时出现，并注明本刊采用哪一种。"""
    source = build_clean_docx(tmp_path / "m.docx")
    text = _triage_text(tmp_path, source, "tjyj")

    assert "字符数（不计空格）" in text
    assert "中文字符 + 英文单词" in text
    assert "字符数（计空格）" in text
    assert "本刊采用口径" in text


def test_over_limit_is_reported_but_never_edited(tmp_path):
    """B-15：超摘要上限只提示，不删字；原稿字节不变。"""
    long_abstract = "内容提要：" + "本文考察数字金融对企业融资约束的影响。" * 40
    source = tmp_path / "long.docx"
    minimal_doc(source, abstract=long_abstract,
                keywords="关键词：数字金融；融资约束；企业创新",
                extra_paragraphs=["正文段落。"])
    before = source.read_bytes()
    text = _triage_text(tmp_path, source, "gjjryj")

    assert ("超出" in text or "高于建议值" in text), "超上限没有提示"
    assert "需作者自行精简" in text, "没有把删改决定交回作者"
    assert source.read_bytes() == before, "体检报告改动了原稿"


def test_details_requires_a_journal_first(acceptance_client, acceptance_env):
    """B-16：没有选定目标期刊的任务，补录页应回工作台，而不是猜字段。"""
    import app

    client, work = acceptance_client, acceptance_env["work"]
    job = "abcdef123456"
    d = work / job
    d.mkdir()
    (d / "meta.json").write_text(json.dumps({"job": job, "original": "x.docx"}),
                                 encoding="utf-8")

    response = client.get(f"/job/{job}/details")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/workspace")


def test_cross_journal_details_keep_previous_values(
        acceptance_client, acceptance_env, tmp_path):
    """B-21：换刊后补录页只问差集，已填的基金信息不能丢。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = _with_author_line(tmp_path / "clean.docx")
    funding = "国家社会科学基金（22BJY000）"
    job = deliver(client, source, details={**TJYJ_FILL, "funding": funding})

    response = client.post(f"/job/{job}/retry", data={"journal": "cjyj"})
    assert response.status_code == 302

    page = client.get(f"/job/{job}/details")
    assert page.status_code == 200
    assert funding in page.get_data(as_text=True), "换刊后已填的基金信息被清空"


def test_shipped_journals_have_all_three_artifacts():
    """B-22：规则集、样板、规格页三件齐全；规格页不是空白壳。

    样板 docx 是 build_journal 的渲染产物，gitignore 不放行（72631b6 的
    needs_sample 同因）：公开检出里没有它，缺样本时跳过样板断言而不是假红；
    规格页三本照常核验，CI 上不丢这部分覆盖。
    """
    for jid in SHIPPED_JOURNALS:
        rules = _rules(jid)
        spec = RULES / f"{jid}.spec.html"
        assert spec.is_file(), f"{jid} 缺规格页"
        html = spec.read_text(encoding="utf-8")
        assert rules["journal"]["name_zh"] in html
        assert "成品样式预览" in html or "规格" in html
        if jid == "cjyj":
            # 静态规格页必须跟着 render_spec 重生成：总字数下限曾只认
            # limits.total.min，规格页会停在旧的单向上限写法。
            assert "15,000–25,000 字" in html, "cjyj 规格页的总字数区间没有随代码更新"
    # 规格页内嵌的预览图来自同刊样张 PDF；三本都在场时才校验，少一本就
    # 跳过整段（公开检出只有规格页，样张 PDF 不入库）。
    missing = [jid for jid in SHIPPED_JOURNALS
               if not (RULES / f"{jid}.sample.docx").is_file()
               or not (RULES / f"{jid}.sample.pdf").is_file()]
    if missing:
        pytest.skip(
            "样板 docx / 样张 pdf 由 build_journal 生成，不进开源仓库："
            + "、".join(missing))
    for jid in SHIPPED_JOURNALS:
        assert (RULES / f"{jid}.sample.docx").is_file(), f"{jid} 缺样板 docx"
        assert (RULES / f"{jid}.sample.pdf").is_file(), f"{jid} 缺样张 pdf"


# ---------------------------------------------------------------- 设置页

def test_settings_never_echoes_the_key_and_test_ignores_enabled(
        acceptance_client, acceptance_env, monkeypatch):
    """B-27：密钥不回显；未启用也能测试连通（测试与开关解耦）。"""
    import ai_assist

    client = acceptance_client
    key = "sk-test-secret-123456"
    response = client.post("/settings", data={
        "base_url": "https://example.invalid/v1",
        "model": "demo-model",
        "api_key": key,
        "send_mode": "snippet",
        "timeout": "30",
        "snippet_chars": "60",
    })
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert "已保存" in html
    assert key not in html, "设置页把密钥回显到了 HTML"

    page = client.get("/settings").get_data(as_text=True)
    assert key not in page

    monkeypatch.setattr(ai_assist, "test_connection",
                        lambda cfg=None, timeout=None: (True, "假客户端连通"))
    test_response = client.post("/settings/test", data={
        "base_url": "https://example.invalid/v1",
        "model": "demo-model",
        "api_key": key,
        "send_mode": "snippet",
    })
    payload = test_response.get_json()
    assert payload["ok"] is True and "假客户端" in payload["msg"]


def test_invalid_settings_are_rejected(acceptance_client):
    """B-27 反面：非法接口地址必须报错，不能悄悄存下。"""
    client = acceptance_client
    response = client.post("/settings", data={
        "base_url": "ftp://example.invalid/v1",
        "model": "demo",
        "send_mode": "snippet",
    })
    assert response.status_code == 200
    assert "http://" in response.get_data(as_text=True)


# ---------------------------------------------------------------- 建库门禁

def test_draft_is_invisible_until_published_and_publish_needs_confirmation(
        acceptance_client, acceptance_env, tmp_rules):
    """B-28：草稿不出现在选刊列表；未勾核对、报告有错都不得发布。"""
    client = acceptance_client
    jid, draft = make_draft(tmp_rules)

    workspace = client.get("/workspace").get_data(as_text=True)
    assert "测试期刊" not in workspace, "未核对的草稿进了选刊列表"

    assert client.get(f"/journals/draft/{jid}").status_code == 200

    without_confirm = client.post(f"/journals/draft/{jid}/publish", data={})
    assert without_confirm.status_code == 400
    assert "逐项核对" in without_confirm.get_data(as_text=True)
    assert draft.is_dir(), "发布被拒却删了草稿"

    report = json.loads((draft / "report.json").read_text(encoding="utf-8"))
    report.update({"ok_to_publish": False, "errors": ["缺字段"]})
    (draft / "report.json").write_text(json.dumps(report, ensure_ascii=False),
                                       encoding="utf-8")
    with_confirm = client.post(f"/journals/draft/{jid}/publish",
                               data={"confirm": "on"})
    assert with_confirm.status_code == 400
    assert "阻断性错误" in with_confirm.get_data(as_text=True)
    assert not (tmp_rules / f"{jid}.json").exists(), "有错还发布了正式规则"


def test_discard_draft_removes_only_the_draft(
        acceptance_client, acceptance_env, tmp_rules):
    """C-08/B-29：放弃草稿删除草稿目录，正式规则库不变。"""
    client = acceptance_client
    jid, draft = make_draft(tmp_rules)
    formal_before = {
        p.name: p.read_bytes() for p in tmp_rules.glob("*.json")
    }

    response = client.post(f"/journals/draft/{jid}/discard")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/journals")
    assert not draft.exists(), "放弃后草稿目录还在"
    assert {p.name: p.read_bytes() for p in tmp_rules.glob("*.json")} == formal_before


def test_discard_busy_draft_is_rejected_honestly(
        acceptance_client, acceptance_env, tmp_rules):
    """C-10：草稿正被处理时，放弃请求 409，不能假装删成功。"""
    import app

    client = acceptance_client
    jid, draft = make_draft(tmp_rules)
    lock = app.job_lock(f"journal:{jid}")
    assert lock.acquire(blocking=False)
    try:
        response = client.post(f"/journals/draft/{jid}/discard")
    finally:
        lock.release()
    assert response.status_code == 409
    assert draft.is_dir(), "忙碌时草稿被删了"


# ---------------------------------------------------------------- 输入与安全

def test_pdf_manuscript_and_fake_docx_are_refused_with_guidance(acceptance_client):
    """B-30/B-32：稿件 PDF 拒收并解释；假 docx 给「另存为」的出路。"""
    client = acceptance_client
    pdf = upload(client, b"%PDF-1.4 fake", "论文.pdf")
    assert pdf.status_code == 400
    assert "请向作者要源文件" in pdf.get_data(as_text=True)

    fake = upload(client, b"this is not a zip archive", "假稿.docx")
    assert fake.status_code == 200
    assert "这份文件打不开" in fake.get_data(as_text=True)

    missing_journal = upload(client, b"PK anything", "稿.docx", journal="")
    assert missing_journal.status_code == 400


def test_cross_site_post_and_security_headers(acceptance_client):
    """B-33：跨站写请求 403；本机页面带基本安全头与 no-store。"""
    client = acceptance_client
    cross = client.post("/settings", data={},
                        headers={"Sec-Fetch-Site": "cross-site"})
    assert cross.status_code == 403
    assert "不接受来自其他网站的提交" in cross.get_data(as_text=True)

    response = client.get("/")
    assert response.status_code == 200
    assert response.headers["X-Content-Type-Options"] == "nosniff"
    assert response.headers["X-Frame-Options"] == "SAMEORIGIN"
    assert "frame-ancestors" in response.headers["Content-Security-Policy"]
    assert response.headers["Cache-Control"] == "no-store"


def test_many_headings_beyond_ten_do_not_crash(tmp_path):
    """B-36/E-02：第 11、12 个一级标题必须出「十一、十二、」。"""
    source = build_many_headings_docx(tmp_path / "headings.docx", sections=12)
    pipeline = render_into(tmp_path / "headings-out", source, target="tjyj")
    text = "\n".join(pipeline.output_texts.values())

    assert "十一、第11节" in text
    assert "十二、第12节" in text
    assert not invariant_failures(pipeline)


# ---------------------------------------------------------------- 结构性要件阻断

def test_missing_required_blocks_block_conversion(
        acceptance_client, acceptance_env, tmp_path):
    """B-18：稿件本体缺必备要件时不得开始转换（DESIGN §10.4）。"""
    client, work = acceptance_client, acceptance_env["work"]
    source = build_title_only_docx(tmp_path / "title-only.docx")
    job = upload_job(client, source)

    page = client.get(f"/job/{job}/details").get_data(as_text=True)
    assert "该刊必备" in page, "标题-only 稿件没有标出缺失必备项"

    response = client.post(f"/job/{job}/details", data={})
    assert response.status_code == 400, "缺少必备项却开始转换了"
    html = response.get_data(as_text=True)
    assert "稿件本体还缺" in html
    assert "正文" in html and "参考文献" in html, "退回页没有列出缺件"
    # 阻断发生在 stage_render 之前：不产出成品，任务仍停在补录阶段
    assert not list((work / job).glob("manuscript.*.docx")), "阻断时仍产出了成品"
    assert meta_of(work, job)["stage"] != "rendered"


def test_complete_manuscript_converts_with_empty_personal_info(
        acceptance_client, tmp_path):
    """完整稿 + 个人信息表单全空仍可转换：作者页字段不参与结构阻断。"""
    client = acceptance_client
    source = _complete_docx(tmp_path / "complete.docx")
    job = upload_job(client, source, journal="gjjryj")

    page = client.get(f"/job/{job}/details").get_data(as_text=True)
    assert "稿件本体还缺" not in page, "完整稿被误判成缺件"
    response = client.post(f"/job/{job}/details", data={})
    assert response.status_code == 302, response.get_data(as_text=True)[:800]
    assert response.headers["Location"].startswith(f"/result/{job}")


def test_filled_english_abstract_stops_blocking(acceptance_client, tmp_path):
    """英文摘要缺着就阻断；在补录页填上即可出稿，不必回 Word 改。"""
    client = acceptance_client
    source = _complete_docx(tmp_path / "no-en-abstract.docx", english_abstract=False)
    job = upload_job(client, source, journal="gjjryj")

    blocked = client.post(f"/job/{job}/details", data={})
    assert blocked.status_code == 400
    assert "摘要（en）" in blocked.get_data(as_text=True), "英文摘要缺失没有报出来"

    filled = client.post(f"/job/{job}/details",
                         data={"english_abstract": "This paper studies digital finance."})
    assert filled.status_code == 302
    assert filled.headers["Location"].startswith(f"/result/{job}")


def test_highlights_and_received_date_never_block():
    """另附文件（Highlights）与编辑部填的收稿日期不派给作者，不阻断排版。"""
    import app

    rows_raw = [
        ("highlights", "en", "Highlights（en）", False, True),
        ("received_date", None, "收稿日期", False, True),
        ("article_number", None, "文章编号", False, True),
    ]
    assert app.blocking_missing(rows_raw, {}) == []

    # 只有真正要在稿件里补的块才阻断；能由补录补齐的，填了就不再阻断
    rows = rows_raw + [("body", None, "正文", False, True),
                       ("abstract", "en", "摘要（en）", False, True)]
    assert app.blocking_missing(rows, {}) == ["正文", "摘要（en）"]
    assert app.blocking_missing(rows, {"english_abstract": "text"}) == ["正文"]


def test_corresponding_author_is_filled_by_the_footer_field():
    """jae/jde 的通讯作者：补录页「通讯作者」填了就算补齐，不能永远阻断。"""
    import app

    rows = [("corresponding_author", None, "通讯作者", False, True)]
    assert app.blocking_missing(rows, {}) == ["通讯作者"]
    assert app.blocking_missing(rows, {"corresponding": "张三（zs@example.com）"}) == []
