"""F20 迁移补全与 jryj 括号裁定的回归钉（2026-11-02 审读三问题）。

两条线，各自钉一件事：

- **F20 键名迁移补全**：首轮只迁了入库三本（tjyj / cjyj / mwm 反推），全库
  还有 7 本残留旧键——cmjj / ej / jqte / kjyj / sjjj / jryj（author-year，
  描述式 cn_formats / en_formats）、ciejournal（numbered，formats）。旧键
  使 `app.details_context` 的引用勾选闸（`bool(cite_rl.get("templates"))`，
  app.py details_context）读不到数据，闸永远关着，作者-年份转换在中文
  author-year 刊上不可达。本轮按 tjyj 逐字验证过的模式补齐：键迁到
  templates、格式串逐字未动、`templates_source` 注明出处。描述式著录写法
  不是占位符模板（citation_ref.format_entry 只认带 `{字段}` 的值），条目
  文字不会被拿去改写——开闸后可用的是正文括注转换，条目原样保留（fail-safe）。
  mwm 的 cn_formats / en_formats 是占位符模板的推导出处（其 templates_source
  明言「由 cn_formats / en_formats 的官方样例反推」），保留不删，删除等于
  销毁出处。
- **jryj 括号裁定**：jryj 曾声明 `bracket: "()"`（半角），与官方投稿指南
  的括注示例（「（Lee et al.，1993）」「已被广泛报道（Clinton，2003）」均
  全角）自相矛盾；指南的半角只见于叙述式（「Clinton(2003)就欧洲问题」），
  且同一句里括注仍全角；入库已刊样本 2 篇（rules/sources/jryj/paper1653、
  paper1655）正文引用 62 处全部全角、零半角。引用转换只产括注形态，故裁定
  为全角「（）」并在规则集留 bracket_source。原值不修正的话，F20 开闸即
  产出「(张三，2020)」半角形态，与该刊主要示例相反。

规则库 20 本不入库（见 .gitignore），涉及私有刊的用例在文件缺失时跳过，
与 conftest「夹具只能用入库刊」的守卫同一思路。
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document

from conftest import ROOT, RULES, TOOLS

sys.path.insert(0, str(TOOLS))

import citation_ref as cr  # noqa: E402

# 本轮补迁移的 7 本：author-year 六本（cn_/en_ 前缀键）+ 编号制一本（纯类型键）
MIGRATED_AUTHOR_YEAR = ("cmjj", "ej", "jqte", "kjyj", "sjjj", "jryj")
MIGRATED_NUMBERED = ("ciejournal",)

# 迁移前原值快照（迁移脚本逐字节校验过的输入，此处冻结为期望值）。
# 键名 = 迁移后 templates 里的键。
VERBATIM = {
    "cmjj": {
        "cn_journal": "作者：《文章标题》，《期刊名》××年第×期。",
        "cn_book": "作者：《书名》，出版社及出版年份。",
        "cn_translation": "[国别]作者：《书名》，译者，出版社及出版年份。",
        "cn_newspaper": "作者：《文章标题》，《报纸名》××年×月×日。",
        "cn_chapter": "文章作者：《文章标题》，载编著者《书名/文集名》，出版社及出版年份。",
        "cn_conference": "作者：《论文标题》，会议名称，年份月份。",
        "cn_working_paper": "作者：《论文标题》，发表单位，论文编号，年份。",
        "cn_web": "作者（责任者）：《文章标题》，网站名称，×年×月×日。",
        "en_journal": "Author, A., Title. Journal Name, Vol.xx, No.x, Month Year, pp.xx-xx.",
        "en_book": "Author, A., Book Title. City: Publisher, Year.",
        "en_chapter": "Author, A., Title. In Editor (eds.), Book Title. City: Publisher, Year, pp.xx-xx.",
        "en_working_paper": "Author, A., Title. NBER Working Paper, No.xxxxx, Year.",
        "en_conference": "Author, A., Title. Conference Name, Month Year.",
    },
    "ej": {
        "cn_journal": "作者1、作者2，年份：《文献题名》，《期刊名》第X期。",
        "cn_book": "作者，年份：《书名》，出版社。",
        "cn_newspaper": "作者，年份：《文章题名》，《报纸名》X月X日。",
        "en_journal": "Family name, Initials., Year, “Title,” Journal Name, Volume(Issue), pp. XX–XX.",
        "en_book": "Family name, Initials., Year, Book Title, Place: Publisher.",
    },
    "jqte": {
        "cn_journal": "作者. 题名［J］. 刊名，年，（期）：起页~止页.",
        "cn_book": "作者. 书名［M］. 出版地：出版社，年.",
        "cn_chapter": "作者. 题名［A］. 主编. 文集名［C］. 出版地：出版社，年.",
        "cn_thesis": "作者. 题名［D］. 学校学位论文类别，年.",
        "cn_newspaper": "作者. 题名［N］. 报纸名，年月日.",
        "en_journal": "Author A.， Author B.， Year， Title ［J］. Journal Name， Volume（Issue）， start~end.",
        "en_book": "Author A.， Year， Title ［M］. City： Publisher.",
        "en_chapter": "Author A.， Year， Title ［A］. Editor. Book Title ［C］. City： Publisher.",
        "en_working_paper": "Author A.， Year， Title ［R］. Working Paper Series， No. xxx.",
    },
    "kjyj": {
        "cn_journal": "著者．出版年．题（篇）名．刊名，卷（期）：页",
        "cn_book": "著者．出版年．书名．版本．出版地：出版者，页次",
        "cn_chapter": "著者．出版年．题（篇）名．见（in）：论文集编者．文集名．出版地：出版者，页次",
        "en_journal": "Author, A. Year. Title. Journal Name, Volume(Issue): pages",
        "en_book": "Author, A. Year. Book Title. Edition. Place: Publisher, pages",
    },
    "sjjj": {
        "cn_journal": "作者1、作者2（年代）：《文献名》，《期刊名》第几期。",
        "cn_translation": "（国别）作者（年代）：《文献名》（译者译），地名：出版社。",
        "en_journal": "Family name, M. First name. “The Title of the Article.” The Name of the Journal, the time, pp. XX-XX.",
        "en_book": "Family name, M. First name. The Title of the Book. Printing place: The Name of Pressing House, time, pp. XX-XX.",
        "en_chapter": "Family name, M. First name. “The Title of the Article,” in First name Family name, eds., The Title of the Book. Printing place: The Name of Pressing house, time.",
        "en_working_paper": "Family name, M. First name. “The Title of the Article.” XXX working paper No. XXX, the time.",
    },
    "jryj": {
        "cn_journal_one_author": "徐忠，2017，《中国稳健性货币政策的实践经验与货币政策理论的国际前沿》，《金融研究》第1期，第1~21页。",
        "cn_journal_two_authors": "王永进和施炳展，2014，《上游垄断与中国企业产品质量升级》，《经济研究》第4期，第116~129页。",
        "cn_journal_three_authors": "叶康涛、张然和徐浩萍，2010，《声誉、制度环境与债务融资——基于中国民营上市公司的证据》，《金融研究》第8期，第171~183页。",
        "cn_book": "白寿彝，1937，《中国交通史》，商务印书馆1937年9月第二版。",
        "cn_chapter": "佐藤宏，2004，《外出务工、谋职和城市劳动力市场——市场支撑机制的社会网络分析》，载李实、佐藤宏主编，《经济转型的代价》，北京中国财政经济出版社，第372~390页。",
        "cn_working_paper": "简军波，2016，《非洲与“海上丝绸之路”：挑战与前景》，复旦国际问题研究院系列工作论文，2016年第1期（总第7期）。",
        "en_journal_one_author": "Todaro, M. P., 1969, \"A Model of Labor Migration and Urban Unemployment in Less Developed Countries,\" The American Economic Review, 59(1), pp.138~148.",
        "en_journal_two_authors": "Banister, J. and J. R. Taylor, 1990, \"China: Surplus Labour and Migration,\" Asia-Pacific Population Journal, 4, pp.3~20.",
        "en_journal_three_authors": "Ghatak, S., P. Levine and S. W. Price, 1996, \"Migration Theories and Evidence: An Assessment,\" Journal of Economic Surveys, 10(2), pp.159~197.",
        "en_working_paper": "Krugman, P., 2006, \"Title of the Article,\" NBER Working Paper, No. 4567.",
        "en_book": "Family name, First name and First Name Family Name, Year, Title of the Monograph, Published by Harvard University Press.",
    },
    "ciejournal": {
        "journal": "作者.文献题名[J].期刊名,出版年,卷(期):页数.",
        "book": "作者.文献题名[M].出版地点：出版者，出版时间.",
        "chapter": "作者.析出文献题名[A].文集责任者.文集题名[C].出版地点：出版者，出版时间.",
        "working_paper": "作者.文献题名[R].机构论文,发表年.",
        "electronic": "作者.文献题名[EB/OL].获取和访问路径，引用年份.",
    },
}


def _rules(jid):
    path = RULES / f"{jid}.json"
    if not path.is_file():
        pytest.skip(f"{jid}.json 不在本机规则库（不入库的私有规则集）")
    return json.loads(path.read_text(encoding="utf-8"))


def _ref_list(jid):
    rl = _rules(jid)["citation"]["reference_list"]
    return rl


# ---------- F20 补全：旧键清零、模板在位、值逐字未动 ----------

@pytest.mark.parametrize("jid", MIGRATED_AUTHOR_YEAR + MIGRATED_NUMBERED)
def test_old_keys_migrated_to_templates(jid):
    """旧键不许残留；templates 必须在位（app 的勾选闸以此为据）。"""
    rl = _ref_list(jid)
    if jid in MIGRATED_AUTHOR_YEAR:
        assert "cn_formats" not in rl and "en_formats" not in rl, \
            f"{jid} 仍有旧键，勾选闸读不到模板"
    else:
        assert "formats" not in rl
    assert rl.get("templates"), f"{jid} 缺 templates，引用转换不可达"
    assert rl.get("templates_source"), f"{jid} 缺 templates_source，迁移出处必须留痕"


@pytest.mark.parametrize("jid", sorted(VERBATIM))
def test_migrated_strings_are_byte_identical(jid):
    """迁移只许动键名：每条格式串与迁移前逐字节一致（含全角括号与波浪线）。"""
    tpl = _ref_list(jid)["templates"]
    for key, expected in VERBATIM[jid].items():
        assert tpl[key] == expected, f"{jid}.templates.{key} 迁移时被改动"


def test_mwm_keeps_templates_and_provenance():
    """mwm 已有占位符模板（首轮迁入）；旧键是其推导出处，保留不删。"""
    rl = _ref_list("mwm")
    assert rl.get("templates"), "mwm 的占位符模板不许丢"
    assert "{" in rl["templates"]["cn_journal"], "mwm 必须保留占位符模板"
    assert "cn_formats" in rl and "en_formats" in rl, \
        "cn_formats/en_formats 是 templates_source 声明的推导出处，删除即销毁出处"


def test_citation_gate_opens_for_migrated_journals(tmp_path):
    """app 的引用勾选闸（details_context）对补迁移的刊必须开。"""
    import app

    for jid in MIGRATED_AUTHOR_YEAR[:2] + MIGRATED_NUMBERED + ("gjjryj",):
        if not (RULES / f"{jid}.json").is_file():
            pytest.skip(f"{jid}.json 不在本机规则库（不入库的私有规则集）")

    d = tmp_path
    ir = {"blocks": [{"idx": 0, "type": "body", "text": "引用了文献[1]和[2]。"}]}
    (d / "manuscript.ir.json").write_text(json.dumps(ir, ensure_ascii=False),
                                          encoding="utf-8")
    for jid in MIGRATED_AUTHOR_YEAR[:2]:  # 抽样作者-年份刊：迁移后闸必须开
        offered = app.details_context(d, {"journal": jid})["citation_offered"]
        assert offered is True, f"{jid} 迁移后勾选闸仍打不开"
    # 编号实引刊（ciejournal）：P4-1 之后渲染层硬闸整批拒绝转换，勾选闸加同款
    # 条件（阶段 D）后对它关——开着的勾选框永远落空，还暗示「能转编号实引」
    numbered = app.details_context(d, {"journal": MIGRATED_NUMBERED[0]})["citation_offered"]
    assert numbered is False
    # 无模板刊（gjjryj，入库）仍然关闸：模板缺席是如实状态，不许虚报
    denied = app.details_context(d, {"journal": "gjjryj"})["citation_offered"]
    assert denied is False


# ---------- 描述式模板不许改写条目（fail-safe 面） ----------

CN_ENTRY = cr.parse_entry("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")


@pytest.mark.parametrize("jid", sorted(VERBATIM))
def test_descriptive_templates_never_rewrite_entries(jid):
    """描述式著录写法没有 {字段}：format_entry 必须返回 None，条目一字不动。"""
    for key, tpl in _ref_list(jid)["templates"].items():
        assert "{" not in tpl, f"{jid}.templates.{key} 意外带占位符"
    assert cr.format_entry(CN_ENTRY, _rules(jid)) is None, \
        f"{jid} 的描述式模板被拿去改写条目了"


# ---------- jryj 括号裁定 ----------

def test_jryj_declares_fullwidth_bracket():
    """jryj 的 bracket 裁定为全角：in_text_style 消费后不得再出半角括注。"""
    rules = _rules("jryj")
    assert rules["citation"]["system"] == "author-year"
    it = rules["citation"]["in_text"]
    assert it["bracket"] == "（）", "jryj 括注裁定为全角，勿回退半角"
    assert "bracket_source" in it, "括号裁定必须留出处注记"
    assert cr.in_text_style(rules) == (("（", "）"), "；")


def test_jryj_examples_kept_verbatim_with_ruling_note():
    """官方指南三条示例逐字保留（含半角的叙述式简化写法），裁定写在 note 里。"""
    it = _rules("jryj")["citation"]["in_text"]
    assert it["examples"] == [
        "另外一种观点认为（Lee et al.，1993）",
        "Lee（1994）认为",
        "Clinton et al.(2003) 对此问题进行了分析",
    ]
    note = it["bracket_source"]
    assert "叙述式" in note and "全角" in note, "注记要说明半角例是叙述式形态"


def test_jryj_rewrite_produces_fullwidth_parens():
    """用 jryj 规则集跑引用改写：合并括注必须是（张三，2020；Smith，2021）。"""
    rules = _rules("jryj")
    index = cr.build_index([
        "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.",
        "[2] Smith J. Digital finance and constraints[J]. Journal of Finance, "
        "2021, 76(2): 20-40.",
    ])
    doc = Document()
    p = doc.add_paragraph("合并引用[1,2]。")
    from render_from_ir import rewrite_citations
    assert rewrite_citations(p, index, rules) == 1
    assert p.text == "合并引用（张三，2020；Smith，2021）。"
    assert "(" not in p.text and ")" not in p.text, "半角括号不得出现"


# ---------- 端到端：jryj 开闸后的产品核心场景 ----------

def _build_doc(path):
    doc = Document()
    doc.add_paragraph("数字金融与融资约束")
    doc.add_paragraph("内容提要：本文考察数字金融对企业融资约束的影响。")
    doc.add_paragraph("正文第一段，引用了文献一[1]和文献二[2]。")
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10.")
    doc.add_paragraph("[2] Smith J. Digital finance and constraints[J]. Journal of Finance, 2021, 76(2): 20-40.")
    doc.save(str(path))
    return path


def test_jryj_end_to_end_conversion(tmp_path):
    """jryj.json（私有库，本机）勾选后端到端：括注全角、条目原样、审计留痕。"""
    if not (RULES / "jryj.json").is_file():
        pytest.skip("jryj.json 不在本机规则库")
    src = _build_doc(tmp_path / "src.docx")
    ir = src.with_suffix(".ir.json")
    r = subprocess.run([sys.executable, str(TOOLS / "parse_docx.py"), str(src),
                        "-o", str(ir)], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=str(ROOT))
    assert r.returncode == 0, r.stderr[-800:]
    out = tmp_path / "out.docx"
    r = subprocess.run([sys.executable, str(TOOLS / "render_from_ir.py"), str(ir),
                        str(RULES / "jryj.json"), "-o", str(out), "--citation"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", cwd=str(ROOT))
    assert r.returncode == 0, r.stderr[-800:]

    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "正文第一段，引用了文献一（张三，2020）和文献二（Smith，2021）。" in texts, \
        "jryj 的正文括注必须是全角"
    assert not any("(张三" in t or "(Smith" in t for t in texts), \
        "半角括注违背裁定"
    # 描述式模板不改正文条目（fail-safe），也不许把格式说明写进成品
    assert "[1] 张三. 数字金融与企业融资[J]. 经济研究, 2020, 55(3): 1-10." in texts
    assert "徐忠，2017" not in "".join(texts), "格式说明泄漏进了成品条目"
    audit = json.loads(out.with_suffix(".audit.json").read_text(encoding="utf-8"))
    assert any(i["action"] == "正文引用改作者-年份" for i in audit["items"])
