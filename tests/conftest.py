"""测试公共夹具。

两个约定决定了夹具的形状：

1. **源稿用合成件**。`tools/render_sample.py` 从规则集渲染出的样板，正文是核验文本
   （「此段用于核验正文版式」这类），可以安全提交进仓库，也不会牵扯任何真实论文。
2. **跨刊渲染**。源稿用 tjyj 渲出来，目标刊是 cjyj——两刊版式差异大，能产生足够多的
   改动条目；同刊自渲染几乎全是「无变化」，测不出东西。
"""

import functools
import json
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

import pytest
from docx import Document
from docx.text.paragraph import Paragraph

# 白名单判定只留一份实现（tools/text_guard.py，G 轮统一）。这里 import 后
# 转出同名 API：现有测试仍从 conftest 取，test_text_guard.py 的 parity 用例
# 盯着两边逐条相同——夹具里再抄一份实现，迟早一边先松。
from text_guard import (
    CAPTION_PREFIX, CAPTION_TYPES, CITATION_REWRITE_RULES, HEADING_NUMBER_PREFIXES,
    HEADING_TYPES, INDENT_LEADING_TYPES, REF_NUMBER_PREFIX, TEXT_ACTIONS,
    content_of, document_texts as body_texts, keyword_terms, normalize, signature,
)
from text_guard import (
    citation_rewrite_blocks as _tg_citation_blocks,
    note_reorder_blocks as _tg_note_reorder,
    whitelist_violations as _tg_whitelist_violations,
)

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
RULES = ROOT / "rules" / "journal"

# 只能用仓库自带的期刊。完整 23 本规则库不入库，夹具若用了私有刊，
# 本机跑得通、CI 必挂（`test_fixtures_use_shipped_journals` 守着这一条）。
# gjjryj→tjyj 在开源的三本里改动最多（12 条审计、6 处文字改动）。
SOURCE_JOURNAL = "gjjryj"  # 国际金融研究
TARGET_JOURNAL = "tjyj"    # 统计研究

# 入库示例刊（F58 收口：原先 test_audit_traceable / test_manuscript_matrix
# 各抄一份）。换刊矩阵、规则引用全库测试都只跑这三本——其余刊不入库，CI 没有。
SHIPPED_JOURNALS = ["cjyj", "tjyj", "gjjryj"]


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


def minimal_doc(path, title="数字金融与企业融资约束",
                abstract="内容提要：本文考察数字金融对企业融资约束的影响。",
                keywords=None, extra_paragraphs=()):
    """最小稿件的公共骨架（F58）：标题＋摘要（＋关键词），再逐段追加附加段落。

    各承诺测试起稿的两三段原先各抄一份，这里收成一个入口。文字全部由调用方
    逐字给定（本函数不悄悄换任何字）；keywords 为 None 表示这一版没有关键词段。
    需要拿到段落实体续造的（脚注引用、公式、图片），拿返回的 Document 接着
    add_paragraph、再 save 一次即可。骨架顺序固定为标题→摘要→关键词→附加段——
    匿名稿那类「作者行必须排第二段」的稿件不适用（front_matter 按位置判定，
    只有第一段保留，见 tools/anonymize.py 的 fm_seen 逻辑），保持手拼。
    """
    doc = Document()
    doc.add_paragraph(title)
    doc.add_paragraph(abstract)
    if keywords is not None:
        doc.add_paragraph(keywords)
    for line in extra_paragraphs:
        doc.add_paragraph(line)
    doc.save(str(path))
    return doc


def mutated_rules(tmp_path, base, name="mutated.json", **keys):
    """复制一份入库规则集、按点路径改键，落到 tmp 下（F58）。

    「复制 tjyj 改一键」在承诺测试里有六处，各自 load→改→dump。收成入口后：
    base 是 rules/journal 下的刊 id；keys 的键名是点路径（"citation.system"），
    数值段表示列表下标（"headings.levels.1"），值为 None 表示整键删除。
    只动点名的键，其余逐字保留原刊内容——值与出处都不碰。
    """
    rules = json.loads((RULES / f"{base}.json").read_text(encoding="utf-8"))
    for dotted, value in keys.items():
        parts = dotted.split(".")
        node = rules
        for part in parts[:-1]:
            node = node[int(part)] if isinstance(node, list) else node.setdefault(part, {})
        if isinstance(node, list):
            node[int(parts[-1])] = value
        elif value is None:
            node.pop(parts[-1], None)
        else:
            node[parts[-1]] = value
    path = Path(tmp_path) / name
    path.write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
    return path


# ---- 白名单口径：实现只有 tools/text_guard.py 一份（G 轮统一）----
# 别名已在文件头 import：normalize / content_of / keyword_terms / signature /
# body_texts 与各常量都直接来自 text_guard，夹具不再留第二份判定。口径说明
# 见 tools/text_guard.py 与 DESIGN §4.1。


@pytest.fixture(scope="session")
def sample_docx(tmp_path_factory):
    """由规则集渲染出的合成源稿（tjyj 版式）。"""
    d = tmp_path_factory.mktemp("sample")
    out = d / "source.docx"
    run_tool("render_sample.py", RULES / f"{SOURCE_JOURNAL}.json", "-o", out)
    return out


class Pipeline:
    """一次「解析 → 渲染 → 审计」的结果，以及常用派生数据。"""

    def __init__(self, workdir, docx, audit, ir, keyword_sep=None, labels=None):
        self.dir = workdir
        self.docx = docx
        self.audit = audit
        self.ir = ir
        self.keyword_sep = keyword_sep
        # 目标刊规则集的标签（blocks[type].zh.label，F43）：标签归一把源稿标签
        # 整段换成这个值，守卫比对时按它剥标签（content_of）。
        self.labels = labels or {}
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
    labels = {b["type"]: b["label"] for b in rules.get("blocks", [])
              if b.get("lang") == "zh" and b.get("type") in ("abstract", "keywords")
              and b.get("label")}
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
                    keyword_sep=kw_sep, labels=labels)


NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def note_reorder_blocks(p):
    """只换了位置、没换文字的表格注（判定在 tools/text_guard.py）。

    接口保持接收 Pipeline：现有断言按 p.source_texts / p.output_texts /
    p.block_type 调用，转一层比让几十处调用点改签名更稳。
    """
    return _tg_note_reorder(p.source_texts, p.output_texts, p.block_type)


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

# --citation 的豁免凭证（F39）：动作与 rule 都对得上才算越线开关（常量在
# tools/text_guard.py，见文件头 import）。


def citation_rewrite_blocks(p):
    """凭引用改写获得豁免的块（判定在 tools/text_guard.py）。"""
    return _tg_citation_blocks(p.items)


def unaudited_changes(p):
    """文字变了却没有审计记录的块（纯换位的表注/图注除外）。"""
    recorded = {it["idx"] for it in p.items}
    return sorted(p.changed_idx - recorded - note_reorder_blocks(p))


def text_snapshot_violations(p):
    """审计「文字快照」登记的双向对账（F54），返回问题清单。

    TEXT_ACTIONS 登记了会真正改写文字的动作，承诺测试对它们做逐字校验；旧的
    版式条目测试是同义反复（用 before/after 为空筛条目、再断言 before/after
    为空），真正的静默缺口在反方向：渲染器新增改文字动作而没登记进
    TEXT_ACTIONS 时，逐字校验整类漏掉。这里双向对账：

    - 已登记（TEXT_ACTIONS ∪ 引用改写，后者另有 rule 对账）的条目必须带
      非空 before/after；
    - 未登记的条目不得携带「与源稿/成品逐字一致」的快照对——新文字动作
      未登记即红，逼它先进 TEXT_ACTIONS 再谈逐字。

    版式条目记录的格式目标值（悬挂缩进的「2.3 字符」）与层序名词（表注／图注
    层序调整的「source / general」）不是文字快照，不在后一条之列。
    """
    registered = TEXT_ACTIONS | set(CITATION_REWRITE_RULES)
    bad = []
    for it in p.items:
        action = (it.get("action") or "").strip()
        before, after = it.get("before") or "", it.get("after") or ""
        idx = it.get("idx")
        if action in registered:
            if not before or not after:
                bad.append((it.get("id"), action, "登记为文字动作，快照却缺了半边"))
        elif (before and after
              and before == p.source_texts.get(idx)
              and after == p.output_texts.get(idx)):
            bad.append((it.get("id"), action, "携带逐字文字快照却未登记进 TEXT_ACTIONS"))
    return bad


def whitelist_violations(p):
    """超出白名单的文字改动 [(idx, 块类型, before, after)]。

    判定在 tools/text_guard.py（与 app 运行期闸同一实现）；这里只把
    Pipeline 字段摊成它的位置参数，保持旧调用点不变。
    """
    return _tg_whitelist_violations(
        p.source_texts, p.output_texts, p.block_type, p.items,
        labels=getattr(p, "labels", None),
        kw_sep=getattr(p, "keyword_sep", None))


def block_count_mismatch(p):
    """源稿与成品的 body 子元素是否还一一对应。"""
    return set(p.output_texts) != set(p.source_texts)


def fingerprint_delta(p, exclude=frozenset()):
    """（凭空新增, 丢失）文字指纹的多重集差；exclude 里的块两侧都不计。"""
    sep = getattr(p, "keyword_sep", None)

    def prints(texts):
        return Counter(signature(t, p.block_type.get(i, ""), sep,
                                 getattr(p, "labels", {}).get(p.block_type.get(i, "")))
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


# ---- 验收层夹具（非 autouse：只影响显式请求的验收用例）----------------
# 验收层不碰真实数据：任务目录、草稿目录、AI 配置全部指到 tmp_path，
# 并封住 socket。建库/发布用例用 tmp_rules 操作规则库副本。

@pytest.fixture()
def acceptance_env(tmp_path, monkeypatch):
    import socket

    import app
    import ai_assist

    work = tmp_path / "work"
    work.mkdir(parents=True)
    monkeypatch.setattr(app, "WORK", work)
    monkeypatch.setattr(app, "DRAFT_DIR", tmp_path / "drafts")
    monkeypatch.setattr(ai_assist, "CONFIG_PATH", tmp_path / "ai.json")

    def deny(*args, **kwargs):
        raise AssertionError("测试期间禁止联网：请把外部调用换成假客户端")

    monkeypatch.setattr(socket, "create_connection", deny)
    monkeypatch.setattr(socket.socket, "connect", deny)
    app.app.config["TESTING"] = True
    return {"work": work, "tmp": tmp_path}


@pytest.fixture()
def acceptance_client(acceptance_env):
    import app

    return app.app.test_client()


@pytest.fixture()
def tmp_rules(tmp_path, monkeypatch):
    """仓库自带 3 本刊的临时副本：发布/覆盖/回滚都只动这里。"""
    import app

    root = tmp_path / "rules"
    root.mkdir()
    for jid in SHIPPED_JOURNALS:
        for suffix in (".json", ".spec.html", ".sample.docx", ".sample.pdf"):
            src = RULES / f"{jid}{suffix}"
            if src.exists():
                shutil.copy2(src, root / src.name)
    monkeypatch.setattr(app, "RULES_DIR", root)
    monkeypatch.setattr(app, "DRAFT_DIR", root / "_drafts")
    return root
