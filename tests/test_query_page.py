"""论文格式查询页：四个约定。

1. **取值只来自规则集**。每个非空单元格都要能经 rules_ref.resolve 解析回
   规则 JSON——查询页不许出现指不到出处的「规定」，同 §8.5 审计引用的口径。
2. **各刊行集合同构**。入库刊的分区/行键序列完全一致，否则对比模式的逐行
   对齐会静默错位：A 列的「行距」对上 B 列的「字号」。
3. **缺席要诚实**。该刊没规定的字段写「未规定/未设/不设上限」，推定项没
   登记就写「没有登记」，不许从同类刊猜值补齐。
4. **入参不可信**。id 先过 JID_RE 再过规则集存在性，路径穿透与不存在的
   id 直接忽略；最多并排 QUERY_MAX 本，超出的截掉。
"""

import re

import pytest
from conftest import SHIPPED_JOURNALS
from rules_ref import resolvable

CARDS_RE = re.compile(r'data-qid="([a-z0-9_]+)"')


@pytest.fixture()
def app_module():
    import app

    app.app.config["TESTING"] = True
    return app


@pytest.fixture()
def client(app_module):
    return app_module.app.test_client()


def test_nav_lists_query_beside_workspace(client):
    """「格式查询」与「格式转换」并列出现在导航里。"""
    html = client.get("/workspace").data.decode("utf-8")
    assert '<a href="/query"' in html
    assert "格式查询" in html


def test_empty_selection_lists_every_journal(app_module, client):
    """不带参数进入查询页：全部期刊以可选卡片出现。"""
    html = client.get("/query").data.decode("utf-8")
    listed = set(CARDS_RE.findall(html))
    assert listed == {j["id"] for j in app_module.load_journals()}


def test_single_view_shows_values_from_rules(client):
    """单刊查询：版式取值与该刊规则集一致。"""
    html = client.get("/query?journals=cjyj").data.decode("utf-8")
    assert "财经研究" in html
    assert "宋体" in html            # typography.body.font_cn
    assert "15.5pt" in html          # typography.body.line_spacing.value_pt
    assert "编号制" in html          # citation.system = numbered
    assert "一、" in html            # headings.levels
    assert "15,000–25,000" in html   # limits.total 的 min–value 区间


def test_missing_fields_are_reported_honestly(client, app_module, monkeypatch):
    """该刊没规定的字段如实写「未规定」；推定项没登记也照实说。"""
    html = client.get("/query?journals=gjjryj").data.decode("utf-8")
    assert "未规定" in html                    # gjjryj 没有 equations / typography.references
    assert "推定项披露" in html                # 披露区始终在场
    # 「没有登记推定项」的空态：出处登记（2026-11-04）后现实库已无零推定刊，
    # 注入一本合成刊把这个分支钉住。
    fake = {"journal": {"id": "zzfake", "name_zh": "零推定刊"}}
    monkeypatch.setattr(app_module, "rules_of",
                        lambda jid: fake if jid == "zzfake" else None)
    html = client.get("/query?journals=zzfake").data.decode("utf-8")
    assert "zzfake" in html                    # 合成刊进入了查询结果
    assert "没有登记推定项" in html            # assumptions 为空时不假装有


def test_assumptions_are_disclosed(client):
    """登记了推定项的刊，推定原文逐条出现在页面上。"""
    html = client.get("/query?journals=cjyj").data.decode("utf-8")
    assert "条推定" in html
    assert "按常见惯例处理" in html   # cjyj.assumptions 里表注层序那条的原文


def test_invalid_and_unknown_ids_are_ignored(client):
    """路径穿透与不存在的 id 被忽略，合法 id 照常渲染。"""
    html = client.get("/query?journals=../evil,nonexistent,cjyj").data.decode("utf-8")
    assert "财经研究" in html
    assert "nonexistent" not in html


def test_duplicate_ids_collapse_to_single_view(client):
    """同一本刊在 URL 里出现两次：去重成单刊视图，不进入对比。"""
    html = client.get("/query?journals=cjyj,cjyj").data.decode("utf-8")
    assert "只看这一本" not in html   # 对比模式列头才有的提示
    assert "用这本刊转换" in html     # 单刊刊名卡在场


def test_selection_is_capped_at_query_max(app_module, client, monkeypatch):
    """并排上限：超过 QUERY_MAX 的刊截掉（CI 只有 3 本入库刊，用调低上限验证）。"""
    monkeypatch.setattr(app_module, "QUERY_MAX", 2)
    html = client.get("/query?journals=cjyj,tjyj,gjjryj").data.decode("utf-8")
    assert html.count("只看这一本") == 2
    assert "添加期刊（2/2）" in html
    assert "国际金融研究<a" not in html  # 第三本没进 chips（卡片里的刊名带 <strong>）


def test_compare_mode_marks_differences(client):
    """两刊对比：出现差异计数与标黄行，数字与标黄行数一致。"""
    html = client.get("/query?journals=cjyj,tjyj").data.decode("utf-8")
    m = re.search(r"<strong>(\d+)</strong>\s*项规定不同", html)
    assert m, "对比模式没有差异计数"
    assert int(m.group(1)) > 0
    assert html.count('qrow qdiff') == int(m.group(1))


def test_compare_diffs_come_from_rules(client):
    """对比抓到的差异必须是规则集里的真实分歧（cjyj 10pt/4 级 vs tjyj 10.5pt/3 级）。"""
    html = client.get("/query?journals=cjyj,tjyj").data.decode("utf-8")
    assert "10.5pt" in html   # tjyj typography.body.size_pt
    assert "（3 级）" in html  # tjyj headings.levels
    assert "（4 级）" in html  # cjyj headings.levels


def test_style_rows_render_for_cjyj(app_module):
    """版式行的取值钉到规则集原文——fmt 函数坏掉时不得悄悄退成「未规定」。

    （曾有事例：_style_txt 的 f-string 笔误让所有版式行静默变 None，而旧断言
    的取值恰好在选择器卡片里也出现，测试照样绿。这里专挑只有分区单元格
    才会有的值。）
    """
    view = app_module.journal_query_view("cjyj")
    cells = {(s["key"], r["key"]): r["cells"][0]
             for s in view["sections"] for r in s["rows"]}
    assert cells[("titlepage", "doc_title")] == "楷体 · 20pt · 居中"
    assert cells[("abstract", "abs_style")] == "楷体 · 9.5pt · 标签「摘　要：」"
    assert cells[("body", "font")] == "宋体 / Times New Roman · 10pt · 两端对齐 · 首行缩进 2 字符"
    assert cells[("titlepage", "authors")] == "仿宋 · 13.5pt · 居中"
    assert cells[("citation", "ref_style")] is None  # cjyj 确实没有文献表版式字段
    assert cells[("table", "border")] == "顶 1.5 / 表头下 0.5 / 底 1.5pt"


def test_view_shapes_are_identical_across_journals(app_module):
    """各刊分区/行键序列完全一致，对比列才不会错位。"""
    views = [app_module.journal_query_view(jid) for jid in SHIPPED_JOURNALS]
    shape0 = [(s["key"], r["key"]) for s in views[0]["sections"] for r in s["rows"]]
    assert shape0
    for v in views[1:]:
        assert [(s["key"], r["key"]) for s in v["sections"] for r in s["rows"]] == shape0


def test_every_displayed_value_resolves_to_rules(app_module):
    """每个非空单元格的引用都能解析回规则集——查询页不许有编造的「规定」。"""
    ref_by_key = {(s["key"], r["key"]): r["ref"] for s in app_module.QUERY_SECTIONS
                  for r in s["rows"]}
    for jid in SHIPPED_JOURNALS:
        rules = app_module.rules_of(jid)
        view = app_module.journal_query_view(jid)
        for sec in view["sections"]:
            for row in sec["rows"]:
                cell = row["cells"][0]
                if cell is None:
                    continue
                ref = ref_by_key[(sec["key"], row["key"])]
                assert resolvable(rules, ref), \
                    f"{jid} 的「{row['label']}」展示了值 {cell!r}，但引用 {ref} 解析不到"


def test_self_compare_has_zero_diffs(app_module):
    """同一本刊与自己对比：差异计数必须是 0（没有规定不同）。"""
    views = [app_module.journal_query_view("cjyj") for _ in range(2)]
    assert app_module._mark_diffs(views) == 0
    assert all(r["same"] for v in views for s in v["sections"] for r in s["rows"])


def test_source_url_only_renders_as_link_when_http(app_module, client, monkeypatch):
    """官方来源是 http(s) 才渲染成链接；javascript: 伪协议降级为纯文本。"""
    html = client.get("/query?journals=cjyj").data.decode("utf-8")
    m = re.search(r"官方来源</div>\s*<div class=\"qval\">(.*?)</div>", html, re.S)
    assert m, "规则来源分区里没有「官方来源」行"
    assert m.group(1).strip().startswith('<a href="https://')

    fake = {"verification": {"source_url": "javascript:alert(1)", "status": "verified"}}
    monkeypatch.setattr(app_module, "rules_of", lambda jid: dict(fake))
    html = client.get("/query?journals=cjyj").data.decode("utf-8")
    assert "javascript:alert(1)" in html          # 值仍如实展示
    assert '<a href="javascript:' not in html     # 但绝不渲染成链接


def test_all_sections_render(app_module, client):
    """单刊页面把每个分区、每一行都渲染出来，一行不少。"""
    html = client.get("/query?journals=cjyj").data.decode("utf-8")
    for sec in app_module.QUERY_SECTIONS:
        assert sec["title"] in html, f"缺分区：{sec['title']}"
        for row in sec["rows"]:
            assert row["label"] in html, f"缺行：{sec['title']}/{row['label']}"


def test_sample_pdf_served_inline(client):
    """样张 PDF 内联出（浏览器可直接翻页缩放），不是附件。"""
    r = client.get("/journal/cjyj/sample.pdf")
    assert r.status_code == 200
    assert r.mimetype == "application/pdf"
    assert "attachment" not in (r.headers.get("Content-Disposition") or "")


def test_sample_docx_downloads_with_readable_name(client):
    """样板 docx 走附件下载，中文名经 RFC 5987 编码后可还原。"""
    from urllib.parse import unquote

    r = client.get("/journal/cjyj/sample")
    assert r.status_code == 200
    disp = r.headers.get("Content-Disposition", "")
    assert "attachment" in disp
    assert "版式样板" in unquote(disp)
    assert r.mimetype.endswith("wordprocessingml.document")


def test_sample_routes_reject_bad_ids(client):
    """样张路由与 spec 同款防护：穿越 id 404，不存在的刊 404。"""
    assert client.get("/journal/..%5Cjournal%5Ctjyj/sample.pdf").status_code == 404
    assert client.get("/journal/..%5Ctjyj/sample").status_code == 404
    assert client.get("/journal/nonexistent/sample.pdf").status_code == 404
    assert client.get("/journal/nonexistent/sample").status_code == 404


def test_query_page_embeds_real_sample_preview(client):
    """查询页嵌入的是真实渲染产物：对比模式两刊的样张都在场。"""
    html = client.get("/query?journals=cjyj,tjyj").data.decode("utf-8")
    assert '/journal/cjyj/sample.pdf' in html
    assert '/journal/tjyj/sample.pdf' in html
    assert '/journal/cjyj/sample"' in html or '/journal/cjyj/sample"' in html
    assert "下载" in html
