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


def normalize(text):
    """把允许的字符级归一全部抹平，剩下的差异才算真改动。

    抹掉的东西都来自 DESIGN §4.1 的白名单（空格、全半角、标题编号），
    以及标题行内的多空格（「引　　言」是版式留白，不是文字）。
    """
    s = re.sub(r"\s+", "", text)
    s = "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c for c in s)
    for pat in (r"^[一二三四五六七八九十]{1,3}、",
                r"^[（(][一二三四五六七八九十]{1,3}[）)]",
                r"^\d{1,2}(?:\.\d{1,2})*[.．]?"):
        s = re.sub(pat, "", s)
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


def signature(text, btype):
    """块的可比对指纹。

    关键词单独处理：分隔符从全角空格换成「；」属于白名单内的归一
    （`关键词分隔符归一`），所以拆成词条再比，而不是整串比。

    参考文献条目剥掉编号前缀：`[1]`→`〔1]` 这类改写法属 N-07 白名单，
    但条目正文必须逐字不变，所以只比前缀之后的部分。
    """
    body = content_of(text, btype)
    if btype == "keywords":
        terms = [t for t in re.split(r"[；;、,，\s\u3000]+", body) if t.strip()]
        return tuple(normalize(t) for t in terms)
    if btype == "reference_item":
        body = re.sub(r"^\s*[\[［〔【(（]?\s*\d{1,3}\s*[\]］〕】)）]?\s*[.．、]?\s*", "", body)
    return (normalize(body),)


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

    def __init__(self, workdir, docx, audit, ir):
        self.dir = workdir
        self.docx = docx
        self.audit = audit
        self.ir = ir
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


def render_into(workdir, source_docx, target=TARGET_JOURNAL, rejects=None, rules_path=None):
    """在指定目录完整跑一遍流水线，返回 Pipeline。"""
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
    cmd = ["render_from_ir.py", d / "manuscript.ir.json", ruleset, "-o", outdocx]
    if rejects:
        (d / "reject.json").write_text(
            json.dumps({"ids": list(rejects)}, ensure_ascii=False), encoding="utf-8")
        cmd += ["--reject", d / "reject.json"]
    run_tool(*cmd)

    audit_path = outdocx.with_suffix(".audit.json")
    return Pipeline(d, outdocx,
                    json.loads(audit_path.read_text(encoding="utf-8")),
                    json.loads((d / "manuscript.ir.json").read_text(encoding="utf-8")))


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


def invariant_failures(p):
    """一次核对四条不变式，返回问题清单（空表示全部通过）。

    单独用例外面的矩阵测试靠它复用同一套判断——不然每换一个期刊组合，
    就得把同样的断言抄一遍，抄着抄着就漏了。
    """
    problems = []
    reordered = note_reorder_blocks(p)
    recorded = {it["idx"] for it in p.items}

    missing = sorted(p.changed_idx - recorded - reordered)
    if missing:
        problems.append(f"文字变了却没有审计记录：{missing[:5]}")

    for idx in sorted(p.changed_idx - reordered):
        btype = p.block_type.get(idx, "")
        if signature(p.source_texts[idx], btype) != signature(p.output_texts[idx], btype):
            problems.append(
                f"块 {idx}（{btype}）超出白名单改动："
                f"{p.source_texts[idx][:20]} → {p.output_texts[idx][:20]}")

    if set(p.output_texts) != set(p.source_texts):
        problems.append("段落数量变了：源稿与成品的 body 子元素不再一一对应")

    def prints(texts):
        return Counter(signature(t, p.block_type.get(i, "")) for i, t in texts.items())

    before, after = prints(p.source_texts), prints(p.output_texts)
    if before - after:
        problems.append(f"成品里丢了文字：{[''.join(s)[:20] for s in (before - after)][:3]}")
    if after - before:
        problems.append(f"成品里凭空多出文字：{[''.join(s)[:20] for s in (after - before)][:3]}")

    for it in p.items:
        if not (it.get("action") or "").strip():
            problems.append(f"审计条目缺动作名：{it.get('id')}")
            break

    return problems
