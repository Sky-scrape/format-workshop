"""补录页按 personal_fields[].placement 分组（DESIGN §10.2）。

旧页面把 author_page 之外的一律归到「投稿系统里填」：署名行、首页脚注、
Highlights 这类另附文件全被说成系统字段，用户看不到稿子里真正会多出什么。
这里用 tjyj 规则组合成一份覆盖全部 placement 的规则集，钉住分组标题、
同义 placement 的合并与未见值的兜底。
"""

import json

import pytest

import app
from conftest import mutated_rules

pytestmark = pytest.mark.acceptance


def _fields():
    """覆盖规则库里出现过的全部 placement，外加一个未登记值。"""
    return [
        {"id": "paper_title", "label_zh": "文章标题",
         "placement": "author_page", "required": True},
        {"id": "author_name", "label_zh": "作者姓名",
         "placement": "byline", "required": True},
        {"id": "funding_info", "label_zh": "基金项目",
         "placement": "footnote", "required": False},
        {"id": "data_statement", "label_zh": "数据可得性声明",
         "placement": "before_references", "required": True},
        {"id": "author_bio", "label_zh": "作者简介",
         "placement": "back_matter", "required": True},
        {"id": "highlights", "label_zh": "Highlights（3–5 条）",
         "placement": "separate_file", "required": True},
        {"id": "cover_letter", "label_zh": "投稿信",
         "placement": "submission_system", "required": False},
        {"id": "reviewers", "label_zh": "推荐审稿人",
         "placement": "system", "required": False},
        {"id": "running_head", "label_zh": "页眉短标题",
         "placement": "title_page", "required": False},
        {"id": "odd_field", "label_zh": "未登记位置的字段",
         "placement": "mystery", "required": False},
    ]


@pytest.fixture()
def placed_rules(tmp_path, monkeypatch):
    """从入库的 tjyj 合成一份 personal_fields 覆盖全部 placement 的规则集。"""
    path = mutated_rules(tmp_path, "tjyj", name="tjyj.json",
                         personal_fields=_fields())
    monkeypatch.setattr(app, "RULES_DIR", tmp_path)
    return json.loads(path.read_text(encoding="utf-8"))


def _groups(rules):
    return {g["key"]: g for g in app.placement_groups(rules)}


def test_each_placement_gets_its_own_group_with_a_plain_description(placed_rules):
    """每个 placement 一组，组名按刊面用词，并带一句人话说明。"""
    groups = _groups(placed_rules)
    expected = {
        "author_page": "作者信息页",
        "byline": "署名行",
        "footnote": "首页脚注",
        "before_references": "参考文献之前",
        "back_matter": "文末",
        "separate_file": "另附文件",
        "submission_system": "投稿系统",
        "title_page": "独立标题页",
        "other": "其他位置",
    }
    assert {key: groups[key]["title"] for key in expected} == expected
    for group in groups.values():
        assert group["desc"].strip(), f"{group['key']} 分组没有一句人话说明"


def test_highlights_land_in_separate_file_not_submission_system(placed_rules):
    """Highlights 是另附文件，不能像旧页面那样说成「投稿系统里填」。"""
    groups = _groups(placed_rules)
    separate = [i["id"] for i in groups["separate_file"]["items"]]
    system = [i["id"] for i in groups["submission_system"]["items"]]
    assert "highlights" in separate
    assert "highlights" not in system


def test_system_spellings_share_one_group(placed_rules):
    """system 与 submission_system 同义（库内两种拼写并存），合成一组。"""
    groups = _groups(placed_rules)
    assert "system" not in groups
    assert {i["id"] for i in groups["submission_system"]["items"]} == {
        "cover_letter", "reviewers"}


def test_unknown_placement_does_not_claim_to_be_a_system_field(placed_rules):
    """未登记的 placement 不猜成投稿系统，归入「其他位置」并如实提示。"""
    groups = _groups(placed_rules)
    assert [i["id"] for i in groups["other"]["items"]] == ["odd_field"]
    assert "odd_field" not in [i["id"] for i in groups["submission_system"]["items"]]


def test_details_context_passes_grouped_fields_to_template(placed_rules, tmp_path):
    """details_context 把分组传给模板，而不是只给一坨非 author_page 字段。"""
    d = tmp_path / "job"
    d.mkdir()
    ctx = app.details_context(d, {"journal": "tjyj", "journal_name": "统计研究"})

    keys = [g["key"] for g in ctx["placement_groups"]]
    assert "separate_file" in keys and "submission_system" in keys
    assert "personal" not in ctx, "旧的一刀切字段列表还在"
    labels = [i["label"] for g in ctx["placement_groups"]
              if g["key"] == "separate_file" for i in g["items"]]
    assert labels == ["Highlights（3–5 条）"]


def test_details_page_renders_group_headings(placed_rules, tmp_path, monkeypatch):
    """页面真的按组渲染：每组标题与说明出现在补录页上。"""
    work = tmp_path / "work"
    work.mkdir()
    monkeypatch.setattr(app, "WORK", work)
    job = "abcdef123456"
    d = work / job
    d.mkdir()
    (d / "meta.json").write_text(json.dumps({
        "job": job, "journal": "tjyj", "journal_name": "统计研究",
        "original": "稿件.docx",
    }, ensure_ascii=False), encoding="utf-8")

    app.app.config["TESTING"] = True
    html = app.app.test_client().get(f"/job/{job}/details").get_data(as_text=True)
    for heading in ("署名行", "首页脚注", "参考文献之前", "文末",
                    "另附文件", "投稿系统", "独立标题页"):
        assert heading in html, f"补录页没有「{heading}」分组"
    assert "Highlights（3–5 条）" in html


def test_preflight_link_is_hidden_until_the_report_exists(
        placed_rules, tmp_path, monkeypatch):
    """「先不排版，仅下载体检报告」只在报告真的存在时给，避免死链接。"""
    work = tmp_path / "work"
    work.mkdir()
    monkeypatch.setattr(app, "WORK", work)
    job = "abcdef123456"
    d = work / job
    d.mkdir()
    (d / "meta.json").write_text(json.dumps({
        "job": job, "journal": "tjyj", "journal_name": "统计研究",
        "original": "稿件.docx",
    }, ensure_ascii=False), encoding="utf-8")
    client = app.app.test_client()

    assert "先不排版" not in client.get(f"/job/{job}/details").get_data(as_text=True)

    (d / "preflight.html").write_text("<html>版式体检</html>", encoding="utf-8")
    html = client.get(f"/job/{job}/details").get_data(as_text=True)
    assert "先不排版，仅下载体检报告" in html
    assert f"/job/{job}/download/preflight" in html
