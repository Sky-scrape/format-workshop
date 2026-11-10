"""F15 / F16 两个已确认问题的回归钉子，用例即问题单里的复现方法。

- F15  防编造主闸有旁路：模型没给引文的事实照常写入规则集、计入「有原文依据」
  （实测 stats={extracted:2,...}、notes 写「2 项有原文依据」，实际只有 1 项有引文；
  body_line / heading_levels 成组项连 unverified 都不记）。
  → 无引文事实一律 unverified：不写规则集正文、不计入 extracted，成组项同闸；
    复核页新增「未证实（需人工核）」分组。
- F16  继承披露只覆盖 FACTS 键：新刊草稿从基线 deepcopy 后，blocks 标签、
  statements、personal_fields、tables 边框等上百字段静默继承。
  → inherited 披露扩到「全部未被覆写的叶子字段」；build 后对基线做叶子级 diff，
    披露集合必须 == 实际继承集合（无静默继承）。
"""

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import extract_journal as ej  # noqa: E402

# 引文必须能被 norm() 后在原文里搜到；ISSN 故意不给引文（F15 的旁路入口）
TEXT = "投稿须知：正文用五号宋体排版，来稿请标明 ISSN。"
EXTRACTED = {"body_size_pt": 10.5, "journal_issn": "1000-9999"}
EVIDENCE = {"body_size_pt": "正文用五号宋体排版"}


@pytest.fixture(scope="module")
def draft():
    rules, report = ej.build(EXTRACTED, EVIDENCE, TEXT,
                             jid="demoj", name_zh="演示期刊甲")
    return rules, report


# ---------------------------------------------------------------- F15

def test_no_quote_fact_never_written(draft):
    """无引文的值不得落进规则集正文——当前值必须仍是基线值。"""
    rules, _ = draft
    base, _ = ej.baseline_rules()
    assert rules["journal"]["issn"] == base["journal"]["issn"], \
        "模型没给引文的 ISSN 竟然写进了规则集"


def test_stats_honest(draft):
    """stats 必须如实：只有带引文的 1 项算「有原文依据」，无引文项单独计 unverified。"""
    rules, report = draft
    s = report["stats"]
    assert s["extracted"] == 1, f"有原文依据应为 1 项（复现时是 2），实际 {s}"
    assert s["unverified"] == 1
    assert s["dropped"] == 0

    ext = rules["verification"]["extraction"]
    assert [u["key"] for u in ext["unverified"]] == ["journal_issn"]
    assert "journal_issn" not in ext["used"], "无引文项不得出现在 used（有原文依据）里"
    assert "1 项有原文依据" in rules["verification"]["notes"]


def test_grouped_items_go_through_same_gate():
    """body_line / heading_levels 成组项必须过同一道闸：没引文 → unverified，不落盘。"""
    rules, report = ej.build(
        {"body_line_mode": "exact", "body_line_value": 20,
         "heading_levels": ["一、", "（一）"]},
        {},  # 一个引文都不给
        TEXT, jid="demog", name_zh="演示期刊乙")
    ext = rules["verification"]["extraction"]
    assert {"body_line", "heading_levels"} <= {u["key"] for u in ext["unverified"]}, \
        "成组项连 unverified 都不记，就同时逃开了标红和丢弃告警"
    assert "body_line" not in ext["used"] and "heading_levels" not in ext["used"]

    base, _ = ej.baseline_rules()
    assert rules["typography"]["body"].get("line_spacing") == \
        base["typography"]["body"].get("line_spacing"), "无引文的行距竟然写进了规则集"
    assert rules["headings"]["levels"] == base["headings"]["levels"], "无引文的标题编号竟然写进了规则集"
    assert report["stats"]["extracted"] == 0


def test_grouped_fake_quote_dropped():
    """成组项给了引文但对不上原文 → 判编造丢弃（与单值字段同罪同罚）。"""
    rules, report = ej.build(
        {"body_line_mode": "exact", "body_line_value": 20},
        {"body_line_value": "行距固定为 20 磅"},  # TEXT 里没有这句话
        TEXT, jid="demod", name_zh="演示期刊丙")
    assert report["stats"]["dropped"] == 1
    assert report["stats"]["extracted"] == 0
    base, _ = ej.baseline_rules()
    assert rules["typography"]["body"].get("line_spacing") == \
        base["typography"]["body"].get("line_spacing")


def test_review_page_has_unverified_section(draft, tmp_path, monkeypatch):
    """复核页必须把未证实项单独亮出来，并标注「没有写入规则集」的风险。"""
    import app as app_module
    rules, report = draft
    d = tmp_path / "demoj"
    d.mkdir()
    (d / "draft.json").write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    (d / "report.json").write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(app_module, "DRAFT_DIR", tmp_path)

    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    page = client.get("/journals/draft/demoj").data.decode("utf-8")

    assert "未证实" in page, "复核页没有「未证实」分组——无引文项逃开了两套复核机制"
    assert "没有写入规则集" in page or "未写入" in page, "未证实分组必须标注风险"
    assert "ISSN" in page, "未证实清单里看不到具体字段"


# ---------------------------------------------------------------- F16

def test_inherited_disclosure_complete(draft):
    """build 后对基线做叶子级 diff：披露集合 == 实际继承集合（无静默继承）。"""
    rules, _ = draft
    base, _ = ej.baseline_rules()
    ext = rules["verification"]["extraction"]
    written = ej.written_paths(ext["used"], name_zh="演示期刊甲")

    actual = set()
    for path, val in ej.iter_leaves(base):
        if not path or path[0] in ("verification", "assumptions"):
            continue  # 草稿自己的核验区由本次重建
        if path in written:
            continue
        if ej._get_path(rules, path) == val:
            actual.add(".".join(str(p) for p in path))

    disclosed = {it["key"] for it in ext["inherited"]}
    assert disclosed == actual, (
        f"静默继承（实际继承但未披露）：{sorted(actual - disclosed)[:8]}\n"
        f"披露失实（披露了但值已不是基线）：{sorted(disclosed - actual)[:8]}")


def test_inherited_covers_structural_sections(draft):
    """blocks/statements/personal_fields/tables 等结构字段必须出现在披露里。"""
    rules, _ = draft
    base, _ = ej.baseline_rules()
    inherited = rules["verification"]["extraction"]["inherited"]
    sections = {it["section"] for it in inherited}
    for sec in ("blocks", "statements", "personal_fields", "tables", "page"):
        assert sec in sections, f"{sec} 区块被静默继承，没进披露"

    blocks_item = next(it for it in inherited if it["key"] == "blocks")
    assert blocks_item["value"] == base["blocks"], "blocks 清单披露的值与基线不符"


def test_inherited_excludes_overwritten(draft):
    """带引文的抽取值不算继承；无引文未写入的字段值仍是基线，属继承。"""
    rules, _ = draft
    inherited = {it["key"] for it in rules["verification"]["extraction"]["inherited"]}
    assert "typography.body.size_pt" not in inherited, "本次抽取覆写的字段被报成了继承"
    assert "journal.issn" in inherited, "未证实的字段没写入，值仍是基线，应属继承披露"


def test_assumptions_stay_readable(draft):
    """结构字段收成一条汇总；事实键仍逐条列出——assumptions 不能变成几百行转储。"""
    rules, _ = draft
    a = rules["assumptions"]
    assert any("结构字段" in x and "逐字段清单" in x for x in a), a[-1]
    fact_lines = [x for x in a if "未从稿约抽到，沿用基线值" in x]
    assert fact_lines, "事实键的逐条推定项丢了"
    assert all("blocks" not in x for x in fact_lines), "结构字段不该逐条塞进 assumptions"


def test_written_paths_covers_grouped_items():
    """written_paths 必须把成组项的落盘路径也算进去，否则叶子 diff 会误报继承。

    登记粒度必须与 iter_leaves 的叶子产出同口径：父路径匹配不上任何叶子，
    等于没登记（行距抽到基线同值时曾把叶子又报成「沿用基线」）。
    """
    used = {"body_line": {}, "heading_levels": {}}
    ps = ej.written_paths(used, name_zh="某刊")
    assert ("typography", "body", "line_spacing", "mode") in ps
    assert ("typography", "body", "line_spacing", "value_pt") in ps
    assert ("typography", "body", "line_spacing", "value") in ps
    assert ("typography", "body", "line_spacing", "equivalent_exact_pt") in ps
    assert ("headings", "levels") in ps and ("headings", "max_level") in ps
    assert ("journal", "id") in ps and ("journal", "name_zh") in ps


def test_extracted_line_spacing_never_double_counted():
    """行距抽到且与基线同值时，line_spacing 叶子不得再进 inherited 双计（F6-5）。"""
    rules, report = ej.build(
        {"body_line_mode": "exact", "body_line_value": 20},
        {"body_line_value": "行距固定 20 磅"},
        "投稿须知：行距固定 20 磅。", jid="demol", name_zh="演示期刊戊")
    assert report["stats"]["extracted"] == 1
    inherited_keys = {it["key"] for it in
                      rules["verification"]["extraction"]["inherited"]}
    assert "typography.body.line_spacing.mode" not in inherited_keys
    assert "typography.body.line_spacing.value_pt" not in inherited_keys


def test_used_reflects_grouped_override():
    """heading_levels 覆写 max_level 时，used 里的 heading_max_level 必须是生效值（F6-3）。"""
    rules, report = ej.build(
        {"heading_max_level": 4, "heading_levels": ["一、", "（一）"]},
        {"heading_max_level": "标题层级最多四级", "heading_levels": "一、（一）"},
        "投稿须知：标题层级最多四级，编号写作一、（一）。", jid="demom", name_zh="演示期刊己")
    assert rules["headings"]["max_level"] == 2
    ext = rules["verification"]["extraction"]
    assert ext["used"]["heading_max_level"]["value"] == 2
    assert "why" in ext["used"]["heading_max_level"]
    assert report["stats"]["extracted"] == 2
