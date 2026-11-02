"""测试公共夹具。

两个约定决定了夹具的形状：

1. **源稿用合成件**。`tools/render_sample.py` 从规则集渲染出的样板，正文是核验文本
   （「此段用于核验正文版式」这类），可以安全提交进仓库，也不会牵扯任何真实论文。
2. **跨刊渲染**。源稿用 tjyj 渲出来，目标刊是 cjyj——两刊版式差异大，能产生足够多的
   改动条目；同刊自渲染几乎全是「无变化」，测不出东西。
"""

import functools
import json
import re
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

import pytest
from docx import Document
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
RULES = ROOT / "rules" / "journal"

# 只能用仓库自带的期刊。完整 23 本规则库不入库，夹具若用了私有刊，
# 本机跑得通、CI 必挂（`test_fixtures_use_shipped_journals` 守着这一条）。
# gjjryj→tjyj 在开源的三本里改动最多（12 条审计、6 处文字改动）。
SOURCE_JOURNAL = "gjjryj"  # 国际金融研究
TARGET_JOURNAL = "tjyj"    # 统计研究


def run_tool(script, *args):
    """调 tools/ 下的命令行脚本；失败时把尾部输出带进异常，便于定位。"""
    cmd = [sys.executable, str(TOOLS / script), *[str(a) for a in args]]
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(ROOT))
    if r.returncode != 0:
        raise RuntimeError(
            f"{script} 执行失败（exit {r.returncode}）\n"
            f"--- stdout ---\n{r.stdout[-2000:]}\n--- stderr ---\n{r.stderr[-2000:]}"
        )
    return r.stdout


def body_texts(docx_path):
    """按 body 子元素下标取段落文字，下标与 IR 里的 idx 同源。"""
    doc = Document(str(docx_path))
    out = {}
    for i, child in enumerate(doc.element.body):
        if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p":
            out[i] = "".join(n.text or "" for n in child.iter()
                             if n.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        elif child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl":
            cells = []
            for tc in child.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
                cells.append("".join(n.text or "" for n in tc.iter()
                                     if n.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
            out[i] = "\x00".join(cells)
    return out


# ---- 白名单口径（DESIGN §4.1 的实现口径，2026-11-02）----
# 渲染器真正落地的字符级归一只有：段首缩进删除（N-01 段首＋N-08，只对
# body/abstract/keywords）；标题编号前缀按规则集重写、未编号节的标题文字整段
# 来自规则集 intro_title_text（「引　　言」的留白是版式）；题注的编号间隔归一；
# 文献条目编号前缀（N-07）。N-02～N-06（正文内部空格、全半角）与段尾空白未实现，
# 守卫逐字比对——「无条件删全部空白」的宽松口径已按 F31 收紧。
INDENT_LEADING_TYPES = {"body", "abstract", "keywords"}
HEADING_TYPES = {"heading1", "heading2", "heading3"}
CAPTION_TYPES = {"table_caption", "figure_caption"}

HEADING_NUMBER_PREFIXES = (re.compile(r"^[一二三四五六七八九十]{1,3}、"),
                           re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]"),
                           re.compile(r"^\d{1,2}(?:\.\d{1,2})*[.．]?"))
CAPTION_PREFIX = re.compile(r"^(表|图)\s*\d+")
REF_NUMBER_PREFIX = re.compile(
    r"^\s*[\[［〔【(（]?\s*\d{1,3}\s*[\]］〕】)）]?\s*[.．、]?\s*")


def normalize(text, btype=""):
    """把该块类型下白名单允许的字符级归一抹平，剩下的差异才算真改动。

    口径按块类型收紧（F31）：段首缩进空白只允许 body/abstract/keywords 删
    （N-01 段首＋N-08 的唯一落地处）；标题允许抹编号前缀与留白（未编号节的
    标题文字整段来自规则集 intro_title_text）；题注允许抹「表1/图2」前缀与
    两端空白（题注归一只改间隔）；文献条目允许抹编号前缀（N-07）。其余——
    包括全角↔半角、中文与拉丁之间的空格——一律不在白名单：normalize('数字
    金融 A 股') 不再等于 normalize('数字金融A股')。
    """
    s = text
    if btype in INDENT_LEADING_TYPES:
        s = s.lstrip("\t \u3000")
    if btype in HEADING_TYPES:
        s = re.sub(r"\s+", "", s.strip())
        for pat in HEADING_NUMBER_PREFIXES:
            s = pat.sub("", s)
    elif btype in CAPTION_TYPES:
        s = CAPTION_PREFIX.sub("", s.strip()).strip()
    elif btype == "reference_item":
        s = REF_NUMBER_PREFIX.sub("", s, count=1)
    return s


def content_of(text, btype):
    """标签块只比对冒号之后的正文。

    「内容摘要：」→「摘  要：」这类属于标签归一（B 档，规则集明确定义），改的是
    标签本身，不是内容。比对时把标签剥掉，剩下的差异才算真改动。

    冒号要同时认半角——真实稿件里「内容摘要:」很常见，标签归一顺带把它改成全角，
    这是允许的，不该被判成改字。

    文献表标题是纯标签（「参考文献」本身没有正文），一律比对为空——它改成
    「参考文献：」属于 `文献表标题归一`，有独立审计条目管着。
    """
    if btype not in ("abstract", "keywords", "references_heading"):
        return text
    m = re.search(r"[:：]", text)
    # 标签很短（"内容摘要:" 6 个字），指望号出现在前 12 个字内才算标签分隔
    if m and m.start() <= 12:
        return text[m.end():]
    return "" if btype == "references_heading" else text.rstrip("：:")


def keyword_terms(body, sep=None):
    """关键词串切词条。口径与渲染器 split_keywords 同一（tools/render_from_ir.py）。

    逗号、顿号、分号是明确的分隔符，总是切；空白分两等——目标分隔符本身是
    空白时全部当分隔符，否则只有「连续两个以上的空白」或全角空格才算
    （英文词组「digital finance」里的单词空格是词间空格）。sep 缺省按非空白
    分支（保守侧）：守卫欠切只是把词条连成大块来比，渲染器错把一个关键词
    拆成两个会被抓到；过切则会把这类错放行。
    """
    pattern = (r"[\s,，;；、]+" if sep is not None and not sep.strip()
               else r"(?:\s{2,}|\u3000|[,，;；、])+")
    return [t for t in re.split(pattern, body or "") if t.strip()]


def signature(text, btype, kw_sep=None):
    """块的可比对指纹。

    关键词单独处理：分隔符从全角空格换成「；」属于规则集定义的归一
    （`关键词分隔符归一`），所以按目标刊的分隔符口径拆成词条再比，词条
    本身逐字比（kw_sep 是目标刊 blocks:keywords/zh/separator 的值）。

    参考文献条目剥掉编号前缀（见 normalize）：`[1]`→`〔1〕` 这类改写法属
    N-07 白名单，但条目正文必须逐字不变，所以只比前缀之后的部分。
    """
    body = content_of(text, btype)
    if btype == "keywords":
        return tuple(normalize(t) for t in keyword_terms(body, kw_sep))
    return (normalize(body, btype),)


# 会真正改写文字的动作。只有这些条目的 before/after 是文字快照，
# 其余（页面设置、套版、悬挂缩进）存的是格式值。
TEXT_ACTIONS = {"标签归一", "关键词分隔符归一", "题注归一", "文献表标题归一",
                "标题重新编号", "文献条目编号归一"}


@pytest.fixture(scope="session")
def sample_docx(tmp_path_factory):
    """由规则集渲染出的合成源稿（tjyj 版式）。"""
    d = tmp_path_factory.mktemp("sample")
    out = d / "source.docx"
    run_tool("render_sample.py", RULES / f"{SOURCE_JOURNAL}.json", "-o", out)
    return out


class Pipeline:
    """一次「解析 → 渲染 → 审计」的结果，以及常用派生数据。"""

    def __init__(self, workdir, docx, audit, ir, keyword_sep=None):
        self.dir = workdir
        self.docx = docx
        self.audit = audit
        self.ir = ir
        self.keyword_sep = keyword_sep
        self.items = audit.get("items", [])
        self.source_texts = body_texts(workdir / "manuscript.docx")
        self.output_texts = body_texts(docx)
        self.block_type = {b["idx"]: b.get("type", "") for b in ir.get("blocks", [])}

    def item_by_id(self, item_id):
        return next((i for i in self.items if i["id"] == item_id), None)

    @property
    def changed_idx(self):
        return {i for i in self.output_texts
                if self.output_texts[i] != self.source_texts.get(i)}

    @property
    def text_editing_items(self):
        """真正改了文字的条目（B 档）；版式类条目的 before/after 都是空串。"""
        return [i for i in self.items if (i.get("before") or i.get("after"))]


@functools.lru_cache(maxsize=1)
def tracked_rulesets():
    """仓库里真正入库的规则集 id 集合。

    完整 23 本不入库，本机有全部——所以「本机跑得通」证明不了什么。
    每次用私有刊都会在这里拦下来，而不是等到 CI 上 FileNotFoundError。
    """
    r = subprocess.run(["git", "ls-files", "rules/journal"], capture_output=True, text=True)
    if r.returncode != 0:
        return None
    return {Path(line).stem for line in r.stdout.splitlines() if line.endswith(".json")}


def render_into(workdir, source_docx, target=TARGET_JOURNAL, rejects=None, rules_path=None,
                citation=False):
    """在指定目录完整跑一遍流水线，返回 Pipeline。

    citation=True 时带上 --citation：渲染器里唯一越线的开关（正文编号引用
    改作者-年份、文献条目按模板重排）。invariant_failures 对这类改写块有
    引用感知的豁免（F39），harness 因此能跑完 citation 路径。
    """
    d = Path(workdir)
    d.mkdir(parents=True, exist_ok=True)

    tracked = tracked_rulesets()
    if tracked is not None and rules_path is None and target not in tracked:
        raise AssertionError(
            f"测试用了未入库的期刊 {target!r}（已入库：{sorted(tracked)}）。"
            f"本机有全部 23 本所以跑得通，CI 上会直接找不到文件。")

    shutil.copy(source_docx, d / "manuscript.docx")
    run_tool("parse_docx.py", d / "manuscript.docx", "-o", d / "manuscript.ir.json")

    ruleset = Path(rules_path) if rules_path else RULES / f"{target}.json"
    outdocx = d / f"manuscript.{target}.docx"
    # 目标刊的关键词分隔符：signature 切词条的口径与渲染器 split_keywords
    # 同一（它读同一个字段），守卫才不会把规则集驱动的分隔符归一误判成改字
    rules = json.loads(Path(ruleset).read_text(encoding="utf-8"))
    kw_sep = next((b.get("separator") for b in rules.get("blocks", [])
                   if b.get("type") == "keywords" and b.get("lang") == "zh"), None)
    cmd = ["render_from_ir.py", d / "manuscript.ir.json", ruleset, "-o", outdocx]
    if rejects:
        (d / "reject.json").write_text(
            json.dumps({"ids": list(rejects)}, ensure_ascii=False), encoding="utf-8")
        cmd += ["--reject", d / "reject.json"]
    if citation:
        cmd += ["--citation"]
    run_tool(*cmd)

    audit_path = outdocx.with_suffix(".audit.json")
    return Pipeline(d, outdocx,
                    json.loads(audit_path.read_text(encoding="utf-8")),
                    json.loads((d / "manuscript.ir.json").read_text(encoding="utf-8")),
                    keyword_sep=kw_sep)


NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def note_reorder_blocks(p):
    """只换了位置、没换文字的表格注。

    目标刊要求「注」与「资料来源」的先后不同时，渲染器会交换这两段的内容
    （`表注／图注层序调整`）。逐块比对必然误判成「改了字」，所以先把这类纯换位
    识别出来：这些块各自的文字都变了，但整组文字守恒。

    只认 table_note——题注（表1 xxx）不参与层序调整，它们的空格归一是另一回事，
    混进来会让整组文字不守恒，反而识别不出真正的换位。
    """
    notes = {i for i in p.changed_idx if p.block_type.get(i) == "table_note"}
    if len(notes) < 2:
        return set()

    def sig(idx, texts):
        return signature(texts[idx], "table_note")

    changed = {i for i in notes if sig(i, p.source_texts) != sig(i, p.output_texts)}
    if not changed:
        return set()

    before = Counter(sig(i, p.source_texts) for i in notes)
    after = Counter(sig(i, p.output_texts) for i in notes)
    return changed if before == after else set()


def run_sizes(docx_path):
    """每个 body 子元素里出现过的字号（pt），下标口径与 IR 的 idx 一致。"""
    doc = Document(str(docx_path))
    out = {}
    for i, child in enumerate(doc.element.body):
        if child.tag != NS_W + "p":
            continue
        sizes = {round(r.font.size.pt, 2)
                 for r in Paragraph(child, doc).runs if r.font.size is not None}
        if sizes:
            out[i] = sizes
    return out


@pytest.fixture(scope="session")
def baseline(tmp_path_factory, sample_docx):
    """基准渲染结果，多个用例共用；流程只跑一次。"""
    return render_into(tmp_path_factory.mktemp("baseline"), sample_docx)


# ---- 共享判定（F31：invariant_failures 与 test_text_fingerprint 同源）----
# 四条承诺各自的不变式在这里各有一个判定函数；矩阵测试的 invariant_failures
# 和承诺测试（test_text_fingerprint.py）调用同一批函数，杜绝「同一不变式
# 两份实现」——抄本迟早和正本漂移，漂移的方向必然是更宽松。

# --citation 的豁免凭证（F39）：这两个动作会改写正文/条目文字，属产品唯一
# 越线开关。豁免以审计条目本身为准——条目存在且 rule 指对才豁免，改了字
# 没记录、或 rule 指错位置的，一律仍按白名单外改动处理。
CITATION_REWRITE_RULES = {
    "正文引用改作者-年份": "citation.in_text",
    "条目按目标刊体例重排": "citation.reference_list.templates",
}


def citation_rewrite_blocks(p):
    """凭引用改写获得豁免的块：动作与 rule 都对得上的审计条目所在块。"""
    return {it["idx"] for it in p.items
            if CITATION_REWRITE_RULES.get(it.get("action")) == it.get("rule")}


def unaudited_changes(p):
    """文字变了却没有审计记录的块（纯换位的表注/图注除外）。"""
    recorded = {it["idx"] for it in p.items}
    return sorted(p.changed_idx - recorded - note_reorder_blocks(p))


def whitelist_violations(p):
    """超出白名单的文字改动 [(idx, 块类型, before, after)]。

    引用改写块豁免逐块比对，但豁免只在条目（action+rule）对得上时成立：
    改了字而没有引用改写条目的，照旧算违规——豁免不豁免审计。
    """
    exempt = citation_rewrite_blocks(p)
    offenders = []
    for idx in sorted(p.changed_idx - note_reorder_blocks(p)):
        btype = p.block_type.get(idx, "")
        if (signature(p.source_texts[idx], btype, getattr(p, "keyword_sep", None))
                == signature(p.output_texts[idx], btype, getattr(p, "keyword_sep", None))):
            continue
        if idx in exempt:
            continue
        offenders.append((idx, btype,
                          p.source_texts[idx][:40], p.output_texts[idx][:40]))
    return offenders


def block_count_mismatch(p):
    """源稿与成品的 body 子元素是否还一一对应。"""
    return set(p.output_texts) != set(p.source_texts)


def fingerprint_delta(p, exclude=frozenset()):
    """（凭空新增, 丢失）文字指纹的多重集差；exclude 里的块两侧都不计。"""
    sep = getattr(p, "keyword_sep", None)

    def prints(texts):
        return Counter(signature(t, p.block_type.get(i, ""), sep)
                       for i, t in texts.items() if i not in exclude)

    before, after = prints(p.source_texts), prints(p.output_texts)
    return after - before, before - after


def change_items_without_rule(p, rules):
    """改动型条目里缺 rule、或 rule 解析不到实处的（DESIGN §8.5 的一般性强制，F41）。

    改动型 = 以下任一：文字快照真的变了（before≠after）；版式套版（动作名以
    「套版」结尾——这类 A 档条目的 before/after 是空串，说明在 note 里）；记
    录了格式目标值（after 非空，如悬挂缩进「2.3 字符」）。说明型（before==
    after 都是空串且非套版，如「编号样文本未转换」「引用体系转换（已跳过）」
    「公式（伪表格）原样保留」）只解释为什么没动，豁免。
    """
    from rules_ref import resolvable

    bad = []
    for it in p.items:
        before, after = it.get("before") or "", it.get("after") or ""
        action = it.get("action") or ""
        if before == after and not after and not action.endswith("套版"):
            continue
        if not it.get("rule") or not resolvable(rules, it["rule"]):
            bad.append((it["id"], action, it.get("rule") or "<缺 rule>"))
    return bad


def layout_signature(docx_path):
    """版式层的结构指纹：word/document.xml 与 word/styles.xml 的规范化序列化。

    docx 是 zip，条目头带时间戳，不能裸比字节（F40）。把承载正文与版式样式
    的这两个 part 解析后按文档顺序重新序列化，并剔除 rsid 修订会话属性
    （人在 Word 里每次编辑都会变，不属于版式），剩下的就是版式结构本身：
    同一输入两次渲染指纹相同，规则集变了指纹必然漂移。

    两个 part 是自家渲染器产出的 XML，解析前仍拒绝 DOCTYPE/ENTITY 声明
    （实体扩展这类解析器风险挡在门外）。
    """
    import xml.etree.ElementTree as ET
    import zipfile

    parts = []
    with zipfile.ZipFile(str(docx_path)) as z:
        for name in ("word/document.xml", "word/styles.xml"):
            raw = z.read(name)
            if b"<!DOCTYPE" in raw or b"<!ENTITY" in raw:
                raise ValueError(f"{name} 带有 DOCTYPE/ENTITY 声明，拒绝解析")
            root = ET.fromstring(raw)
            for el in root.iter():
                for attr in [a for a in el.attrib if "rsid" in a.lower()]:
                    del el.attrib[attr]
            parts.append(ET.tostring(root))
    return b"\n".join(parts)


def invariant_failures(p):
    """一次核对四条不变式，返回问题清单（空表示全部通过）。

    判定全部来自上面那批共享函数，与 test_text_fingerprint.py 的承诺测试
    同源（F31）——矩阵测试只是把同一批判定装配成消息。

    引用模式（--citation，F39）：引用改写块豁免「白名单外改动」与守恒比对，
    豁免凭证是 action 与 rule 都对得上的审计条目（见 CITATION_REWRITE_RULES）。
    """
    problems = []

    missing = unaudited_changes(p)
    if missing:
        problems.append(f"文字变了却没有审计记录：{missing[:5]}")

    for idx, btype, before, after in whitelist_violations(p):
        problems.append(f"块 {idx}（{btype}）超出白名单改动：{before[:20]} → {after[:20]}")

    if block_count_mismatch(p):
        problems.append("段落数量变了：源稿与成品的 body 子元素不再一一对应")

    exclude = note_reorder_blocks(p) | citation_rewrite_blocks(p)
    created, lost = fingerprint_delta(p, exclude=exclude)
    if lost:
        problems.append(f"成品里丢了文字：{[''.join(s)[:20] for s in lost][:3]}")
    if created:
        problems.append(f"成品里凭空多出文字：{[''.join(s)[:20] for s in created][:3]}")

    for it in p.items:
        if not (it.get("action") or "").strip():
            problems.append(f"审计条目缺动作名：{it.get('id')}")
            break

    return problems
