"""第六批「文档与规则集」修复回归钉（F56/F60/F61/F62/F63/F64/F65/READMEIT）。

八个条目都在文档与规则集数据面，用例按各自的证据复现：

- F56  cjyj 页边距估算值只增补登记进 assumptions；体检 CLI 通过。
- F60  死键 intro_numbered（12 本）与拼写漂移 footnotes_scope:"page"（11 本）
       迁移：规则集零残留、jmsc 的活键 intro_number: 0 保留；渲染器只读
       intro_number——给规则集副本塞回 intro_numbered 版式签名必须不变，
       改 intro_number 必须变（这正是删除死键而语义不变的依据）。迁移前后
       各渲染 tjyj/cjyj 的版式签名零漂移已在本批会话实测（notes 留档），
       这里把「死键无效果」钉成永久守卫；audit_rulesets 对未知 headings.*
       键与非枚举 footnotes_scope 出 warning。
- F61  体检报告文末「2026-11-03 勘误」逐条对齐（cdxb 已闭合、jmsc 仅图注层
       待补且 [] 是合法编码、213/464 为补录前口径），且勘误依据的数据事实
       仍成立。
- F62  schema 结构图 verification 移到顶层；placement 枚举补三值并配含义。
- F63  覆盖面计数跨文档一致：会计单列（1 本）、管理学 6 本，桶合计以实际清单为准。
- F64  needs.md 同步 29 本：第七/八批入库记录与中文 22 / 英文 7 口径都在。
- F65  schema §3 block 类型总表覆盖规则集在用的每一个类型（含本批补的
       subtitle/doi/article_number/column_name/endnotes）。
- READMEIT  已知限制声明重排条目为纯文本、原条目内斜体不保留。

夹具姿势：渲染走子进程（conftest.run_tool）；体检工具照
test_fix_toolchain_isolation.py 的姿势在进程内 import。
"""

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import ROOT, RULES, TOOLS, layout_signature, run_tool

SCHEMA_MD = ROOT / "rules" / "journal-schema.md"
COVERAGE_MD = ROOT / "rules" / "coverage.md"
GAPS_MD = ROOT / "rules" / "gaps.md"
NEEDS_MD = ROOT / "rules" / "needs.md"
AUDIT_MD = ROOT / "rules" / "audit-2026-11-01.md"
README_MD = ROOT / "README.md"

# 本文件钉的是「台账 ↔ 全库磁盘状态」的一致性：coverage/gaps/needs 描述完整
# 规则库（29 本），开源检出只有 3 本示例（README「开源边界」），文件数对不上
# 是常态不是回归——全库一致性只在本地全库环境跑，公开检出整文件跳过。
pytestmark = pytest.mark.skipif(
    len(list(RULES.glob("*.json"))) < 10,
    reason="台账一致性以完整规则库为对象，公开检出（3 本示例）不适用")


def read(path):
    return path.read_text(encoding="utf-8")


# ---------------------------------------------------------------- F56

def test_f56_cjyj_margin_assumption_registered():
    """页边距是估算值，就必须显式登记（自家规范 §13.3、README 红线 2）。"""
    rules = json.loads(read(RULES / "cjyj.json"))
    hits = [a for a in rules["assumptions"] if "页边距" in a]
    assert hits, "cjyj 页边距四值为估算，assumptions 必须有登记条目"
    text = hits[0]
    for side in ("top", "bottom", "left", "right"):
        val = str(float(rules["page"]["margin_mm"][side]))
        assert val in text, f"登记条目缺实测原值 {val}（{side}）：{text}"
    assert "未见官方明文" in text and "2026-11-01 体检认定" in text


def test_f56_audit_cli_passes_with_zero_known_key_warnings():
    """体检 CLI 全量通过，且已知键告警列为 0（F60 迁移完成的旁证）。"""
    r = subprocess.run([sys.executable, str(TOOLS / "audit_rulesets.py")],
                       capture_output=True, text=True, cwd=str(ROOT))
    assert r.returncode == 0, f"体检失败：{r.stdout[-800:]}"
    assert "份失败" not in r.stdout
    warned = [ln for ln in r.stdout.splitlines()
              if ln.startswith("| ") and not ln.startswith("| 期刊")
              and not set(ln) <= {"|", "-", " ", ":"}]
    assert warned, "体检表没有数据行？"
    assert all(ln.rstrip().endswith("| 0 |") for ln in warned), \
        "有期刊带已知键告警：\n" + "\n".join(warned)


# ---------------------------------------------------------------- F60

def test_f60_no_dead_or_misspelled_keys_in_rulesets():
    """迁移后 grep 口径：规则集里无 intro_numbered、无 footnotes_scope:"page"。"""
    offenders = []
    for p in sorted(RULES.glob("*.json")):
        t = read(p)
        if "intro_numbered" in t:
            offenders.append(f"{p.name}: intro_numbered 仍在")
        if '"footnotes_scope": "page"' in t:
            offenders.append(f'{p.name}: footnotes_scope:"page" 仍在')
    assert not offenders, "；".join(offenders)


def test_f60_jmsc_intro_number_zero_kept():
    """唯一非 1 起编的 jmsc 走活键 intro_number: 0，迁移不得伤及。"""
    headings = json.loads(read(RULES / "jmsc.json"))["headings"]
    assert headings.get("intro_number") == 0


def _render_signature(tmp_path, name, rules):
    p = tmp_path / f"{name}.json"
    p.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / f"{name}.docx"
    run_tool("render_sample.py", p, "-o", out)
    return layout_signature(out)


@pytest.mark.parametrize("jid", ["tjyj", "cjyj"])
def test_f60_intro_numbered_is_dead_and_intro_number_is_live(tmp_path, jid):
    """渲染器只读 intro_number（缺省 1）：死键塞回去签名不变，活键一动就变。"""
    rules = json.loads(read(RULES / f"{jid}.json"))
    base = _render_signature(tmp_path, f"{jid}_base", rules)

    revived = json.loads(json.dumps(rules))
    revived["headings"]["intro_numbered"] = True   # 死键：迁移删掉的就是它
    assert _render_signature(tmp_path, f"{jid}_revived", revived) == base, \
        "intro_numbered 影响了输出——它就不是死键，删除它就是语义变更"

    zero = json.loads(json.dumps(rules))
    zero["headings"]["intro_number"] = 0           # 活键：jmsc 的用法
    assert _render_signature(tmp_path, f"{jid}_zero", zero) != base, \
        "intro_number=0 没有改变引言编号——活键失效了"

    with_path = tmp_path / f"{jid}_scope.json"
    scope = json.loads(json.dumps(rules))
    scope["numbering"]["footnotes_scope"] = "page"  # 同义拼写：无消费者
    with_path.write_text(json.dumps(scope, ensure_ascii=False), encoding="utf-8")
    assert _render_signature(tmp_path, f"{jid}_scope", scope) == base, \
        'footnotes_scope:"page" 影响了输出——与 per_page 同义的声明不成立'


def test_f60_audit_warns_on_unknown_headings_key(tmp_path):
    sys.path.insert(0, str(TOOLS))
    import audit_rulesets as ar

    rules = {
        "journal": {"id": "keycheck", "name_zh": "键检查刊", "language": "zh"},
        "headings": {"numbering": "cn-hierarchical", "intro_nubmered": True},
    }
    (tmp_path / "x.json").write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    rows, failures = ar.audit_all(tmp_path)
    assert not failures
    assert any("headings.intro_nubmered" in w for w in rows[0]["warnings"]), \
        f"未知 headings 键没有告警：{rows[0]['warnings']}"


@pytest.mark.parametrize("scope, expect", [
    ("paragraph", "不在枚举"),          # 拼写漂移出的第四种写法
    ("page", "同义"),                   # 历史同义拼写，建议改写 per_page
])
def test_f60_audit_warns_on_nonenum_scope(tmp_path, scope, expect):
    sys.path.insert(0, str(TOOLS))
    import audit_rulesets as ar

    rules = {
        "journal": {"id": "scopecheck", "name_zh": "拼写检查刊", "language": "zh"},
        "numbering": {"footnotes_scope": scope},
    }
    (tmp_path / "x.json").write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    rows, failures = ar.audit_all(tmp_path)
    assert not failures
    assert any(expect in w and "footnotes_scope" in w for w in rows[0]["warnings"]), \
        f"footnotes_scope={scope!r} 没有告警：{rows[0]['warnings']}"


def test_f60_audit_silent_on_migrated_scope():
    """迁移后的合法值不再告警。"""
    sys.path.insert(0, str(TOOLS))
    import audit_rulesets as ar

    for jid in ("tjyj", "cjyj", "cer"):
        rules = json.loads(read(RULES / f"{jid}.json"))
        assert ar.check_known_keys(rules) == [], f"{jid} 不应有已知键告警"


# ---------------------------------------------------------------- F61

def test_f61_erratum_section_present_and_aligned():
    t = read(AUDIT_MD)
    assert "## 七、2026-11-03 勘误" in t, "勘误小节缺失"
    tail = t[t.index("## 七、2026-11-03 勘误"):]
    # cdxb：三处旧文案对齐到第三节 7 的闭合状态
    assert "cdxb" in tail and "partial → verified" in tail and "第三节 7" in tail
    # jmsc：仅图注层待补，[] 是合法编码
    assert "figures.note_layers" in tail and "`[]`" in tail and "图注层" in tail
    # 213/464 是补录前口径，以 107/469 为准
    assert "213/464" in tail and "107/469" in tail and "旧口径" in tail


def test_f61_erratum_facts_hold_in_data():
    """勘误写下的数据事实必须仍然成立，否则勘误自己成了新的脱节。"""
    jmsc = json.loads(read(RULES / "jmsc.json"))
    assert jmsc["tables"]["note_layers"], "勘误依据：jmsc 表注层应有记录"
    assert jmsc["figures"]["note_layers"] == [], "勘误依据：jmsc 应仅图注层为空"

    sys.path.insert(0, str(TOOLS))
    import audit_rulesets as ar
    rows, failures = ar.audit_all(RULES)
    assert not failures
    unknown = sum(r["provenance"]["unknown"] for r in rows)
    total = sum(r["size_fields"] for r in rows)
    assert f"{unknown}/{total}" == "0/627", \
        f"出处计数漂成 {unknown}/{total}——勘误第 4 条的口径要跟着对齐"
    # 推定类单独成桶：纯推定项有出处，但必须与官方/实测区分开、可追踪复核
    assert sum(r["provenance"]["inferred"] for r in rows) == 52, \
        "inferred 桶漂移——勘误第 4 条的推定计数要跟着对齐"


# ---------------------------------------------------------------- F62

def _schema_between(start, end):
    t = read(SCHEMA_MD)
    return t[t.index(start):t.index(end)]


def test_f62_verification_is_top_level_in_diagram():
    sec2 = _schema_between("## 2. 顶层结构", "## 3. blocks")
    journal_line = next(ln for ln in sec2.splitlines() if '"journal"' in ln)
    assert "verification" not in journal_line, "verification 还画在 journal 对象内"
    assert re.search(r'^\s*"verification":\s*\{', sec2, re.M), \
        "结构图里没有顶层 verification 键"


def test_f62_placement_enum_has_three_new_values():
    sec9 = _schema_between("## 9. personal_fields", "## 10. ")
    row = next(ln for ln in sec9.splitlines() if "`placement` |" in ln)
    for val in ("byline", "footnote", "section", "author_page", "back_matter", "system"):
        assert f"`{val}`" in row, f"placement 枚举缺 {val}：{row}"
    assert "作者信息页构件" in row and "文末要件" in row and "投稿系统" in row, \
        f"三个新值缺一句含义：{row}"


# ---------------------------------------------------------------- F63

def test_f63_bucket_counts_match_actual_list():
    cov = read(COVERAGE_MD)
    # 覆盖面行随批次会增加方向（第七批加了「财经综合与人口资源环境」）。
    # 不把方向名写死，只要求「中文总数 = 各方向计数之和」且中英合计等于实际清单。
    m = re.search(r"覆盖面：中文 (\d+) 本，覆盖(.+?)；英文 (\d+) 本", cov)
    assert m, "coverage.md 覆盖面行格式变了"
    zh_total, buckets, en_total = int(m.group(1)), m.group(2), int(m.group(3))
    counts = [int(n) for n in re.findall(r"（(\d+)", buckets)]
    assert counts and sum(counts) == zh_total, \
        f"各方向计数之和 {sum(counts)} 与中文总数 {zh_total} 不一致"
    assert "管理学（7）" not in cov, "会计不应折入管理学"
    assert "管理学（6）" in cov and "会计（1）" in cov, "管理学/会计分桶变了"
    assert zh_total + en_total == len(list(RULES.glob("*.json"))), \
        "覆盖面总数与 rules/journal 实际清单不一致"


def test_f63_counts_agree_across_documents():
    gaps = read(GAPS_MD)
    mgmt_row = next(ln for ln in gaps.splitlines() if ln.startswith("| 管理学 |"))
    acct_row = next(ln for ln in gaps.splitlines() if ln.startswith("| 会计 |"))
    assert mgmt_row.rstrip().endswith("| 6 |"), mgmt_row
    assert acct_row.rstrip().endswith("| 1 |"), acct_row

    schema_tail = read(SCHEMA_MD)
    final = [ln for ln in schema_tail.splitlines() if ln.startswith("至此中文")][-1]
    assert "管理学方向 6 本" in final and "会计研究单列" in final, final


# ---------------------------------------------------------------- F64

def test_f64_needs_md_synced_to_current_library():
    t = read(NEEDS_MD)
    assert "均已凭正文 PDF 补齐，22 本" not in t, "needs.md 还停在 22 本"
    # 第八批（2026-11-12）入库后：中文 22 + 英文 7 = 29 本
    assert "共 29 本" in t, "needs.md 没有同步到 29 本"
    assert "第七批" in t, "第七批入库记录缺失"
    assert "第八批" in t, "第八批入库记录缺失"
    assert "**管理科学学报 / 系统工程理论与实践 / 管理工程学报**" not in t


# ---------------------------------------------------------------- F65

def test_f65_schema_block_table_covers_all_used_types():
    sec3 = _schema_between("## 3. blocks", "## 4. citation")
    documented = set(re.findall(r"^\| `([a-z_]+)` \|", sec3, re.M))
    used = set()
    for p in RULES.glob("*.json"):
        for b in json.loads(read(p)).get("blocks", []):
            used.add(b["type"])
    missing = used - documented
    assert not missing, f"§3 总表缺规则集在用类型：{sorted(missing)}"
    for name in ("subtitle", "doi", "article_number", "column_name", "endnotes"):
        assert name in documented, f"本批应补进 §3 的 {name} 不在总表"


# ---------------------------------------------------------------- READMEIT

def test_readmeit_italic_loss_disclosed():
    """已知限制必须说清斜体的真实边界（2026-11-13：可按 italic_fields 保留）。

    旧版声明「重排后条目一律纯文本」；现在规则集声明 italic_fields 就会
    保留斜体区间。限制条目仍要在场，且必须说明「没声明就不加斜体」——
    不能把可选的版式支持写成默认能力。
    """
    t = read(README_MD)
    known = t[t.index("## 已知限制"):]
    m = re.search(r"^- \*\*参考文献只改编号写法\*\*.*$", known, re.M | re.S)
    assert m, "已知限制里找不到参考文献条目那一条"
    bullet = m.group(0).split("\n- ")[0]
    assert "italic_fields" in bullet and "没声明就保持纯文本" in bullet, \
        f"已知限制未声明斜体的声明式边界：{bullet[-160:]}"


# ---------------------------------------------------------------- PROVCLS

def test_size_provenance_classes():
    """字号出处四级判定的契约（2026-11-04 P1 闭合时落地）。

    - official：出处文本含官方词（模板/体例/指南/…）且无实测词
    - official+measured：两者都在（互证或冲突并存，如「体例 9pt、成品 10.3pt」）
    - measured：measured 块 / *_source 写实测
    - inferred：纯「推定」标记（不含官方/实测词）也算有出处，但单独成桶
    - font_size_source 与 font_size_pt_source 一样被认（gjjryj 在用的拼写）
    """
    sys.path.insert(0, str(TOOLS))
    import audit_rulesets as ar

    R = {
        "typography": {
            "a": {"size_pt": 10.5, "source": "官方模板：正文五号"},
            "b": {"size_pt": 9.0, "note": "推定：按五号处理"},
            "c": {"size_pt": 12.0, "source": "体例明文五号；成品读数 11.8 相符"},
            "d": {"size_pt": 10.0, "measured": {"line_pt": 15.2}},
            "e": {"size_pt": 7.5, "size_source": "实测：表注读数 7.2"},
            "f": {"font_size_pt": 9.0, "font_size_source": "推定：暂按正文五号处理"},
            "g": {"size_pt": 11.0},
            "h": {"size_pt": 12.0, "source": "推定：官方未规定此项"},
        },
    }
    P = ar.size_provenance
    assert P(R, ("typography", "a", "size_pt")) == "official"
    assert P(R, ("typography", "b", "size_pt")) == "inferred"
    assert P(R, ("typography", "c", "size_pt")) == "official+measured"
    assert P(R, ("typography", "d", "size_pt")) == "measured"
    assert P(R, ("typography", "e", "size_pt")) == "measured"
    assert P(R, ("typography", "f", "font_size_pt")) == "inferred"
    assert P(R, ("typography", "g", "size_pt")) == "unknown"
    # 「推定」与官方词并存时按现有优先级归 official（历史行为，不悄悄翻类）
    assert P(R, ("typography", "h", "size_pt")) == "official"
