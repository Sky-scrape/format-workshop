"""自定义格式建库：上传 Word/TXT 稿约 + 无 AI 手工建库 + 复核页逐项修正。

这组用例守的是新增能力的三条底线：

1. 稿约不再只收 PDF：.docx / .txt 必须真的取到正文（含 Word 表格）；
2. AI 关着也能建库：生成基线草稿，不把整条路堵死；
3. 手工修正只改被改过的字段：改完要从「沿用基线」清单移出并留痕，
   没改的字段继续如实披露，不能假装「都抽到了」。
"""

import io
import json
import sys
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import extract_journal as ej  # noqa: E402
import render_sample          # noqa: E402
import app as app_module      # noqa: E402


# ---------------------------------------------------------------- 取文

def test_read_text_source_docx_includes_tables(tmp_path):
    """Word 稿约常把字号写进表格；只读 paragraphs 会漏掉整张表。"""
    doc = Document()
    doc.add_paragraph("投稿须知")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "正文字号"
    table.cell(0, 1).text = "小五 10.5pt"
    table.cell(1, 0).text = "行距"
    table.cell(1, 1).text = "20 磅"
    path = tmp_path / "guide.docx"
    doc.save(str(path))

    text, meta = ej.read_source(path)
    assert "投稿须知" in text
    assert "正文字号" in text and "小五 10.5pt" in text, "Word 表格里的要求丢了"
    assert "行距" in text and "20 磅" in text
    assert meta["method"] == "docx"
    assert meta["truncated"] is False


def test_read_text_source_txt_accepts_gbk(tmp_path):
    """国内稿约 TXT 常见 GBK；不能让用户自己猜编码。"""
    path = tmp_path / "guide.txt"
    path.write_bytes("正文用宋体，字号 12。".encode("gb18030"))
    text, meta = ej.read_source(path)
    assert text == "正文用宋体，字号 12。"
    assert meta["method"] == "txt"


# ---------------------------------------------------------------- 路由

def _manual_ctx(tmp_path, monkeypatch, ai_ok=False):
    monkeypatch.setattr(app_module, "DRAFT_DIR", tmp_path)
    monkeypatch.setattr(app_module.ai_assist, "available", lambda cfg=None: ai_ok)
    app_module.app.config["TESTING"] = True
    return app_module.app.test_client()


def _upload(client, name="guide.txt", body="投稿须知：正文用五号宋体。", mode="manual",
            jid="customfmt"):
    payload = body.encode("utf-8") if isinstance(body, str) else body
    return client.post("/journals/extract", data={
        "jid": jid, "name": "自定义刊", "mode": mode,
        "file": (io.BytesIO(payload), name),
    }, content_type="multipart/form-data")


def test_manual_upload_creates_baseline_draft_without_ai(tmp_path, monkeypatch):
    """AI 不可用时上传 TXT：生成基线草稿，全部字段进「沿用基线」。"""
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    resp = _upload(client)

    assert resp.status_code == 302
    assert resp.headers["Location"].endswith("/journals/draft/customfmt")
    d = tmp_path / "customfmt"
    assert (d / "source.txt").is_file(), "上传的 TXT 原件要保留"
    assert (d / "source_text.txt").is_file()
    draft = json.loads((d / "draft.json").read_text(encoding="utf-8"))
    ext = draft["verification"]["extraction"]
    assert ext["mode"] == "manual"
    assert ext["text_method"] == "txt"
    assert ext["used"] == {}
    assert ext["manual"] == {}
    assert ext["inherited"], "手工草稿必须有完整的沿用基线披露"
    assert draft["verification"]["verified_by"] == "手工建库，未经稿约抽取"


def test_auto_mode_falls_back_to_manual_when_ai_off(tmp_path, monkeypatch):
    """用户点了「AI 抽取」但 AI 没配：不报死，改走手工建库。"""
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    resp = _upload(client, mode="auto")
    assert resp.status_code == 302
    draft = json.loads((tmp_path / "customfmt" / "draft.json").read_text(encoding="utf-8"))
    assert draft["verification"]["extraction"]["mode"] == "manual"


def test_docx_upload_reaches_draft_page(tmp_path, monkeypatch):
    """Word 稿约从上传到复核页全链路可走。"""
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    doc = Document()
    doc.add_paragraph("投稿须知：正文用小四号宋体。")
    buf = io.BytesIO()
    doc.save(buf)
    resp = _upload(client, name="guide.docx", body=buf.getvalue())
    assert resp.status_code == 302
    draft = json.loads((tmp_path / "customfmt" / "draft.json").read_text(encoding="utf-8"))
    assert draft["verification"]["extraction"]["text_method"] == "docx"


def test_unsupported_doc_suffix_is_rejected(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    resp = _upload(client, name="guide.doc")
    assert resp.status_code == 400
    assert "另存为 .docx" in resp.data.decode("utf-8")


def test_draft_page_has_manual_edit_form(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    _upload(client)
    page = client.get("/journals/draft/customfmt").data.decode("utf-8")
    assert "手工修正规则值" in page
    assert 'name="f_body_size_pt"' in page
    assert 'name="f_body_line_mode"' in page
    assert 'name="f_heading_levels"' in page
    assert "/journals/draft/customfmt/edit" in page


def test_new_page_offers_docx_txt_and_manual_mode(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    page = client.get("/journals/new").data.decode("utf-8")
    assert 'accept=".pdf,.docx,.txt,.md"' in page
    assert 'name="mode" value="manual"' in page, "AI 不可用时手工建库入口必须还在"
    assert 'name="baseline"' in page, "结构基线选择缺失"


# ---------------------------------------------------------------- 手工修正

def test_edit_applies_values_and_refreshes_disclosure(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    _upload(client)
    d = tmp_path / "customfmt"

    resp = client.post("/journals/draft/customfmt/edit", data={
        "f_journal_name_zh": "我的刊",
        "f_body_font_cn": "黑体",
        "f_body_size_pt": "12",
        "f_body_line_mode": "multiple",
        "f_body_line_value": "1.5",
        "f_heading_levels": "一、\n（一）\n1.",
        "f_citation_system": "author-year",
    })
    assert resp.status_code == 302

    draft = json.loads((d / "draft.json").read_text(encoding="utf-8"))
    ext = draft["verification"]["extraction"]
    assert draft["journal"]["name_zh"] == "我的刊"
    assert draft["typography"]["body"]["font_cn"] == "黑体"
    assert draft["typography"]["body"]["size_pt"] == 12
    assert draft["typography"]["body"]["line_spacing"] == {
        "mode": "multiple", "value": 1.5, "equivalent_exact_pt": 18.0}
    assert draft["headings"]["levels"] == ["一、", "（一）", "1."]
    assert draft["headings"]["max_level"] == 3
    assert draft["citation"]["system"] == "author-year"

    assert set(ext["manual"]) >= {"journal_name_zh", "body_font_cn",
                                  "body_size_pt", "body_line", "citation_system"}
    inherited = {it["key"] for it in ext["inherited"]}
    assert "typography.body.size_pt" not in inherited, "改过的字段还留在沿用基线清单"
    assert "typography.body.font_cn" not in inherited
    assert ext["unverified"] == []
    assert draft["verification"]["verified_by"].startswith("手工建库")
    assert any("手工修正" in a or "沿用基线值" in a for a in draft["assumptions"])


def test_edit_only_marks_changed_fields(tmp_path, monkeypatch):
    """表单提交全部字段，但只有真的改过的才应进 manual。"""
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    _upload(client)
    before = json.loads((tmp_path / "customfmt" / "draft.json").read_text(encoding="utf-8"))
    client.post("/journals/draft/customfmt/edit", data={"f_body_size_pt": "12"})
    after = json.loads((tmp_path / "customfmt" / "draft.json").read_text(encoding="utf-8"))
    manual = after["verification"]["extraction"]["manual"]
    assert list(manual) == ["body_size_pt"], f"未改动字段被记成手工修正：{list(manual)}"
    assert len(after["verification"]["extraction"]["inherited"]) == \
        len(before["verification"]["extraction"]["inherited"]) - 1


def test_edit_rejects_invalid_value_without_writing(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    _upload(client)
    d = tmp_path / "customfmt"
    before = (d / "draft.json").read_text(encoding="utf-8")

    resp = client.post("/journals/draft/customfmt/edit",
                       data={"f_body_size_pt": "abc"})
    assert resp.status_code == 400
    assert "正文字号" in resp.data.decode("utf-8")
    assert (d / "draft.json").read_text(encoding="utf-8") == before, \
        "有类型错误时仍写入了半份修改"


def test_out_of_range_value_is_saved_but_blocks_publish(tmp_path, monkeypatch):
    """超出结构范围的数值允许落草稿，但校验必须把它列成阻断项。

    与既有 AI 草稿口径一致：草稿可以先保存一份待修的值，发布闸认得出来；
    类型错误（如字母冒充字号）才整单拒收。
    """
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    _upload(client)
    resp = client.post("/journals/draft/customfmt/edit",
                       data={"f_body_size_pt": "999"})
    assert resp.status_code == 302
    report = json.loads((tmp_path / "customfmt" / "report.json")
                        .read_text(encoding="utf-8"))
    assert report["ok_to_publish"] is False
    assert any(e["field"] == "typography.body.size_pt" for e in report["errors"])


def test_manual_override_of_ai_field_keeps_original_evidence(tmp_path):
    """AI 抽到的字段被人工改掉：从 used 移出，但原文引文留在手工记录里。"""
    text = "投稿须知：正文用小五号宋体。"
    rules, _ = ej.build(
        {"body_size_pt": 10.5}, {"body_size_pt": "正文用小五号宋体"},
        text, jid="customai", name_zh="演示刊")
    assert "body_size_pt" in rules["verification"]["extraction"]["used"]

    out, report, errors = ej.apply_manual(rules, {"body_size_pt": "12"})
    assert errors == []
    ext = out["verification"]["extraction"]
    assert "body_size_pt" not in ext["used"], "被人工覆盖的 AI 值还算有原文依据"
    assert ext["manual"]["body_size_pt"]["value"] == 12
    assert ext["manual"]["body_size_pt"]["evidence"] == "正文用小五号宋体"
    assert report["stats"]["manual"] == 1
    assert report["ok_to_publish"] is True


def test_apply_manual_rejects_bad_heading_levels():
    rules, _ = ej.build({}, {}, "无关正文", jid="customh", name_zh="演示刊")
    out, report, errors = ej.apply_manual(rules, {"heading_levels": "  "})
    assert errors and "至少保留一级" in errors[0]
    assert out is rules and report is None


# ---------------------------------------------------------------- 结构基线

def test_build_can_choose_structural_baseline():
    """手工建库可从另一本已入库刊起步，而不是永远套默认刊的结构。"""
    rules, report = ej.build({}, {}, "无关正文", jid="customb",
                             name_zh="演示刊", baseline_id="tjyj")
    ext = rules["verification"]["extraction"]
    assert ext["baseline"] == "tjyj"
    assert report["stats"]["extracted"] == 0
    # 复核页「沿用基线」披露必须对回 tjyj，而不是默认的 gjjryj
    assert any("《tjyj》" in a for a in rules["assumptions"])
    out, report2, errors = ej.apply_manual(rules, {})
    assert errors == []
    assert out["verification"]["extraction"]["baseline"] == "tjyj"
    assert report2["stats"]["inherited"] == len(ext["inherited"])


def test_upload_route_passes_baseline(tmp_path, monkeypatch):
    client = _manual_ctx(tmp_path, monkeypatch, ai_ok=False)
    resp = client.post("/journals/extract", data={
        "jid": "custombase", "name": "演示刊", "mode": "manual",
        "baseline": "tjyj",
        "file": (io.BytesIO("正文用五号宋体".encode("utf-8")), "guide.txt"),
    }, content_type="multipart/form-data")
    assert resp.status_code == 302
    draft = json.loads((tmp_path / "custombase" / "draft.json")
                      .read_text(encoding="utf-8"))
    assert draft["verification"]["extraction"]["baseline"] == "tjyj"


def test_render_sample_handles_one_and_two_heading_levels(tmp_path):
    """自定义刊可能只声明 1–2 级标题；样板生成不能拿 levels[2] 裸下标。

    真实回归：两级标题的草稿在发布阶段 render_sample 抛 IndexError，
    页面只显示「稿件结构超出识别范围」，用户无法从报错知道是层级问题。
    """
    rules, _ = ej.build({}, {}, "稿约正文", jid="customl", name_zh="演示刊",
                        baseline_id="tjyj")
    for levels in (["一、"], ["一、", "（一）"]):
        one = json.loads(json.dumps(rules))
        one["headings"]["levels"] = levels
        one["headings"]["max_level"] = len(levels)
        out = tmp_path / f"sample-{len(levels)}.docx"
        render_sample.build(one, str(out))
        assert out.exists() and out.stat().st_size > 0
        text = "\n".join(p.text for p in Document(str(out)).paragraphs)
        if len(levels) < 2:
            assert "理论基础" not in text, "一级刊不该生成二级标题"
        if len(levels) < 3:
            assert "货币政策传导机制" not in text, "两级刊不该生成三级标题"
