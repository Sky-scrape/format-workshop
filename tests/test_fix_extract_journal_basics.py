"""F38 回归：extract_journal.py 的离线纯函数单测。

这个脚本产出「唯一真源」规则集的草稿，此前零测试；它的 bug 会顺着
build → draft.json → publish 直接污染四条承诺的源头。本文件全部离线
运行：不调 AI（ai_facts 不测）、不读 PDF（pdf_text 不测），只测结构
映射与闸门——norm / _cast / _set_path / _get_path / baseline_rules /
build / validate。引文闸与继承披露此处只钉最基本的行为，更细的回归
以 F15/F16 的专项文件为准（若有），不在此重复。
"""

import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parent.parent / "tools"
sys.path.insert(0, str(TOOLS))
import extract_journal as ej  # noqa: E402

# 与稿件原文逐字对应（含全角空格与换行），供引文闸比对
GAZE_TEXT = "来稿要求：正文用\u3000小五号宋体，A4 纸张。论文总字数不超过 20000 字。"


# ---------------------------------------------------------------- norm

def test_norm_kills_layout_whitespace_only():
    """去空白与全/不换空格：排版差异不许把真引文判成编造。"""
    assert ej.norm("正文\u3000用 小五号\xa0宋体\n") == "正文用小五号宋体"
    assert ej.norm(None) == ""
    # 引文里的空白抹平后能在原文里搜到——闸门靠它放行
    assert ej.norm("正文用小五号宋体") in ej.norm(GAZE_TEXT)


# ---------------------------------------------------------------- _cast

@pytest.mark.parametrize("raw,kind,expected", [
    ("10.5", "float", (10.5, "")),
    ("12pt", "float", (12.0, "")),          # 带 pt 单位也能剥掉
    ("abc", "float", (None, "类型不符")),
    (["一、"], "float", (None, "类型不符")),  # 列表冒充数字 → 拒收
    ("15000", "int", (15000, "")),
    ("66.9", "int", (66, "")),
    ("15,000", "int", (None, "类型不符")),   # 千分位逗号不是数字
    (True, "bool", (True, "")),
    ("是", "bool", (True, "")),
    ("no", "bool", (False, "")),
    ("maybe", "bool", (None, "不是布尔值")),
    (["需要"], "bool", (None, "不是布尔值")),
    ("zh", "enum:zh|en", ("zh", "")),
    ("fr", "enum:zh|en", (None, "不在 ['zh', 'en'] 内")),
    (["A4"], "enum:A4|B5|16开", (None, "不在 ['A4', 'B5', '16开'] 内")),
    (" 金融研究 ", "str", ("金融研究", "")),
    ("  ", "str", (None, "空串")),
    (None, "str", (None, "空值")),
])
def test_cast_gates_types(raw, kind, expected):
    """类型闸：字符串数字、布尔、列表与枚举各走各的判据。"""
    assert ej._cast(raw, kind) == expected


# ---------------------------------------------------------------- 路径读写

def test_set_path_creates_deep_keys():
    r = {}
    assert ej._set_path(r, ("typography", "body", "size_pt"), 10.5) is True
    assert r["typography"]["body"]["size_pt"] == 10.5


def test_set_path_refuses_scalar_intermediate():
    """中间层已是标量时必须失败，而不是把标量当字典硬塞。"""
    r = {"page": "A4"}
    assert ej._set_path(r, ("page", "margin_mm", "top"), 25) is False
    assert r["page"] == "A4"


def test_get_path_reads_and_misses_quietly():
    r = {"typography": {"body": {"size_pt": 10.5}}}
    assert ej._get_path(r, ("typography", "body", "size_pt")) == 10.5
    assert ej._get_path(r, ("typography", "title")) is None
    assert ej._get_path(r, ("typography", "body", "size_pt", "x")) is None


# ---------------------------------------------------------------- baseline

def test_baseline_rules_returns_complete_shipped_structure():
    """结构基线取仓库自带规则集：键齐、blocks 是列表、字号是数值。"""
    rules, name = ej.baseline_rules()
    assert isinstance(rules, dict) and isinstance(name, str)
    for key in ("journal", "page", "typography", "headings", "blocks"):
        assert key in rules, f"基线缺 {key} 区块"
    assert isinstance(rules["blocks"], list)
    assert isinstance(rules["typography"]["body"]["size_pt"], (int, float))


# ---------------------------------------------------------------- build

def test_build_applies_fact_backed_by_verbatim_quote():
    """引文逐字对得上 → 值写入规则集，其余沿基线并如实披露。"""
    rules, report = ej.build(
        {"body_size_pt": 10.5}, {"body_size_pt": "正文用小五号宋体"},
        GAZE_TEXT, jid="f38t", name_zh="测试刊")
    assert rules["typography"]["body"]["size_pt"] == 10.5
    assert rules["journal"]["id"] == "f38t"
    assert rules["journal"]["name_zh"] == "测试刊"
    assert rules["verification"]["status"] == "unverified", "AI 草稿必须标未核验"
    assert report["stats"]["extracted"] == 1
    assert report["ok_to_publish"] is True
    # 用户亲手填的刊名不算「沿用基线」，否则复核页出现假风险
    assert "journal_name_zh" not in [x["key"] for x in
                                     rules["verification"]["extraction"]["inherited"]]


def test_build_drops_fabricated_evidence():
    """复现 F38 主闸：引文在原文搜不到 = 编造，整条丢弃、基线值不动。"""
    rules, report = ej.build(
        {"body_size_pt": 9}, {"body_size_pt": "正文用四号字"},
        GAZE_TEXT, jid="f38t")
    assert rules["typography"]["body"]["size_pt"] == 10.5, "编造条目竟改掉了基线值"
    assert report["stats"]["dropped"] == 1
    dropped = rules["verification"]["extraction"]["dropped"]
    assert dropped[0]["why"].startswith("引文在原文中找不到")


def test_build_applies_but_flags_missing_evidence():
    """有值没引文：值照写，但必须进 unverified 清单等人工核。"""
    rules, report = ej.build(
        {"body_size_pt": 10.5}, {}, GAZE_TEXT, jid="f38t")
    assert rules["typography"]["body"]["size_pt"] == 10.5
    assert report["stats"]["unverified"] == 1
    uv = rules["verification"]["extraction"]["unverified"]
    assert uv[0]["key"] == "body_size_pt" and "无法核验" in uv[0]["why"]


def test_build_discloses_inherited_baseline():
    """复现 F38 继承披露：没抽到的键逐条进 inherited 与 assumptions。

    F16 之后披露扩到全部叶子（不止 FACTS 的 49 个键），总数必然大于
    len(FACTS)；这里钉 F38 的本意：49 个事实键一个不落——逐条进
    inherited（fact=True）并逐条进 assumptions（沿用基线值），结构字段
    另收一条汇总并指向 inherited 明细（更细的钉子见 F16 专项文件）。
    """
    rules, report = ej.build({}, {}, GAZE_TEXT, jid="f38t")
    ext = rules["verification"]["extraction"]
    assert report["stats"]["extracted"] == 0
    assert report["stats"]["inherited"] == len(ext["inherited"]) > len(ej.FACTS)
    assert report["needs_review"] == report["stats"]["inherited"]
    fact_paths = {".".join(map(str, meta[0])) for meta in ej.FACTS.values()}
    fact_items = [it for it in ext["inherited"] if it["fact"]]
    assert {it["key"] for it in fact_items} == fact_paths
    fact_lines = [a for a in rules["assumptions"] if "未从稿约抽到，沿用基线值" in a]
    assert len(fact_lines) == len(ej.FACTS)
    assert all("沿用基线值" in a for a in fact_lines)
    assert any("结构字段" in a and "逐字段清单" in a for a in rules["assumptions"]), \
        "结构字段必须收成一条汇总披露，否则 blocks 等上百字段成了静默继承"


def test_build_grouped_line_spacing_and_heading_levels():
    """成组项：倍数行距带折算磅值；标题层级数组同步 max_level。

    F15 之后成组项与单值字段过同一道引文闸：引文必须逐字在稿约原文里，
    且没引文的成组项不落盘，所以这里的稿约文本要真的含这句稿约，
    行距也要有自己的引文（无引文/假引文的淘汰路径见 F15 专项文件）。
    """
    text = GAZE_TEXT + "标题编号为 一、（一）1.，1.5 倍行距。"
    rules, report = ej.build(
        {"heading_levels": ["一、", "（一）", "1."], "body_line_mode": "multiple",
         "body_line_value": 1.5},
        {"heading_levels": "标题编号为 一、（一）1.，1.5 倍行距",
         "body_line_value": "1.5 倍行距"},
        text, jid="f38t")
    assert rules["headings"]["levels"] == ["一、", "（一）", "1."]
    assert rules["headings"]["max_level"] == 3
    ls = rules["typography"]["body"]["line_spacing"]
    assert ls["mode"] == "multiple" and ls["value"] == 1.5
    assert ls["equivalent_exact_pt"] == round(10.5 * 1.5, 1)  # 折算磅值保留 1 位小数


# ---------------------------------------------------------------- validate

def test_validate_passes_clean_baseline():
    v = ej.validate(ej.baseline_rules()[0])
    assert v["ok_to_publish"] is True and v["errors"] == []


def test_validate_aggregates_all_errors_and_only_warns_softly():
    """错误逐项汇总不短路；边距越界与丢弃项只进 warnings。"""
    rules, _ = ej.baseline_rules()
    bad = dict(rules)
    bad["journal"] = {"id": "ABC"}                       # 无刊名 + 代号违规
    bad["blocks"] = []
    bad["typography"] = {"body": {"size_pt": "large",
                                  "line_spacing": {"mode": "exact"}}}
    bad["page"] = {"margin_mm": {"top": 3, "bottom": 30, "left": 25, "right": 25}}
    v = ej.validate(bad, inherited=[{"key": "x"}],
                    dropped=[{"label": "正文字号", "why": "类型不符"}])

    fields = [e["field"] for e in v["errors"]]
    assert fields == ["journal.name_zh", "journal.id", "blocks",
                      "typography.body.size_pt", "typography.body.line_spacing"], fields
    assert v["ok_to_publish"] is False, "有错误项还允许发布"

    warn_fields = [w["field"] for w in v["warnings"]]
    assert "page.margin_mm.top" in warn_fields, "边距越界应只告警不阻断"
    assert "正文字号" in warn_fields, "丢弃项没有转成复核提示"
    assert v["needs_review"] == 1


def test_validate_flags_keyword_bounds_inversion():
    rules, _ = ej.baseline_rules()
    rules.setdefault("limits", {})["keywords"] = {"min": 8, "max": 3}
    v = ej.validate(rules)
    assert v["ok_to_publish"] is True, "下限>上限是软告警，不该阻断保存"
    assert any(w["field"] == "limits.keywords" for w in v["warnings"])
