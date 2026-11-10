"""由成品稿派生一份「投稿用匿名稿」（DESIGN §11.4 的双版本交付）。

署名稿保持完整、供留档与校样；匿名稿只删身份信息，不动任何正文文字。
两份文件同源同版式，差别仅在身份信息，且删了什么逐条写进审计报告。

只删这几类东西：
  1. 前置区的作者署名行、单位行（标题与破折号副标题保留）
  2. 首页脚注 / 脚注 / 尾注里的作者简介、基金项目、通讯作者
  3. 页眉页脚里的身份信息（首页脚注整段清，那是补录注入的通道）
  4. 批注：正文连作者整删
  5. 「致谢」小节（若有）

用法：
    python tools/anonymize.py 成品.docx 稿件.ir.json 规则集.json -o 匿名稿.docx
"""

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from docx.table import Table
from docx.text.paragraph import Paragraph

from lxml import etree

import sys

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_meta import assert_clean, parse_part_xml, sanitize_docx

# 脚注/尾注的作者信息判定，三档（用在 apply 第 4 步）：
#   ① STRONG_RE 强信号——作者简介、通讯作者、ORCID、Email/邮箱/电子邮箱/@、
#     手机号、座机、基金。命中任一即删整条，不需要单位词佐证：「通讯作者：张三，
#     Email：…」这类没写单位的简介、「基金项目：国家自然科学基金（72003012）」
#     这类没写单位的基金注靠双条件会漏（漏删是身份泄露，不可挽回；基金项目
#     是 DESIGN §13.1 匿名口径里的必删五类之一）。
#   ② UNIT_RE×IDENT_RE 双命中——单位词与身份词同时出现才删，弱信号兜底。
#   ③ 两条都不命中——保留。
# 注意①是单条件判定，不看这条注是不是「内容注」：纯内容注只要含「电子邮箱」
# 「@」「手机号」这类字样（如「问卷经电子邮箱发放并回收。」）也会连内容
# 一起删。这是既定取舍：多删可从署名稿整份找回、删了什么逐条写进审计
# （removed），漏删一条身份信息则不可挽回。
UNIT_RE = re.compile(r"(大学|学院|研究院|研究所|商学院|经济学院|管理学院|系\b)")
IDENT_RE = re.compile(r"(邮政编码|邮编|通讯地址|通信地址|电子信箱|邮箱|E-?mail|@|"
                      r"教授|副教授|讲师|博士生|硕士生|研究方向|通讯作者|基金|项目编号|批准号)")
STRONG_RE = re.compile(
    r"作者简介|通讯作者|通信作者|Corresponding\s+author|ORCID|"
    r"E-?mail|电子邮件|电子邮箱|电子信箱|邮箱|@"
    r"|(?<!\d)1[3-9]\d{9}(?!\d)"                        # 手机号
    r"|(?<!\d)0\d{2,3}[-－]\d{3,4}[-－]?\d{4}(?!\d)"    # 座机（区号-号码）
    r"|基金",                                           # 基金项目注：必删，不要求单位词佐证
    re.IGNORECASE)
# 成品里标题可能被渲染器重新编号（「致谢」→「二、致谢」，cn-hierarchical
# 逐刊重写编号前缀）。核对文字时把编号前缀剥掉再比，剥法与渲染器同族。
NUM_PREFIX_RX = re.compile(r"^[（(]?[一二三四五六七八九十\d]{1,4}[、.．）)]?\s*")
ACK_HEAD_RE = re.compile(r"(致\s*谢|鸣\s*谢|Acknowledg(e)?ments?)[：:。]?\s*")


def _iter_children(doc):
    """按 body 子元素顺序产出 (序号, 元素)，序号与 IR 的 idx 对齐。"""
    for i, child in enumerate(doc.element.body):
        if child.tag == qn("w:p"):
            yield i, Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield i, Table(child, doc)


def _identity_hit(text):
    """作者地盘（前置区连字符行、页眉页脚）的身份判定。

    前置区、页眉页脚本来就属于作者署名区，出现单位词、联系方式、身份词
    任一即按身份信息处理：拿不准时删（进 removed，署名稿里还在），
    好过把单位或地址留在匿名稿里。
    """
    return bool(UNIT_RE.search(text) or IDENT_RE.search(text) or STRONG_RE.search(text))


def _norm_text(text):
    """比对口径：去掉全部空白。渲染器会把「致谢」写成「致  谢」，段内
    tab/换行在 IR 与 python-docx 里的记法也不一致，删前核对只看实际字符。"""
    return "".join((text or "").split())


def _norm_core(text):
    """比对口径②：再剥一个编号前缀。渲染器按目标刊重写标题编号
    （「致谢」→「二、致谢」），IR 里的文字还是解析时的原样，逐字比会对不上。"""
    stripped = NUM_PREFIX_RX.sub("", (text or "").strip())
    return _norm_text(stripped)


def _require_paired(by_idx, b, el):
    """删前置区身份行前的最后一道核对（F4-5）：待删元素必须真的坐在
    IR 声明的位置上，且文字与 IR 一致（或只差编号前缀——渲染器会重写
    标题编号）。错位时删的是无关段落、真身份行反而留下，两头都不可
    挽回，宁可拒绝产出请人工重跑。sdt 等定位不到的块在调用方已跳过，
    不会走到这里。"""
    if el is not by_idx.get(b["idx"]):
        raise SystemExit(
            f"idx 漂移：IR 第 {b['idx']} 块与待删元素不对应，"
            "IR 与成品不配套，请人工检查")
    want = _norm_text(b.get("text"))
    if want and want not in (_norm_text(el.text), _norm_core(el.text)):
        raise SystemExit(
            f"idx 漂移：IR 第 {b['idx']} 块「{(b.get('text') or '')[:20]}」与成品"
            f"对应位置「{el.text[:20]}」文字不一致，IR 与成品不配套，请人工检查")


def _droppable(by_idx, b, by_text_fallback=False):
    """确认 IR 块在成品里有位置、且与 IR 文字对得上，才允许按它删除。

    IR 是解析时的快照：sdt 内容控件里的块、渲染后搬过家的块在成品 body
    里可能找不到或对不上。找不到的静默跳过（sdt 段落不参与前置区推断是
    parse 端的既定行为，不在这里重复删除），对得上的返回元素——身份块的
    删除必须经过它核对，防止按错位 idx 删掉无关段落（F4-5）。核对不上
    （错位删错段落、真身份行留下）两头都不可挽回，拒绝产出请人工重跑。

    by_text_fallback：给致谢小节的成员块用。表注层序重排（渲染器既定
    行为）会整体搬动 table_note 元素，idx 对不上但文字还在——按文字
    全文定位，找不到才跳过。定位以「整段文字一致」为准，不会错删。
    """
    el = by_idx.get(b["idx"])
    if el is not None:
        if isinstance(el, Paragraph):
            have = el.text
        else:
            have = "".join(t.text or "" for t in el._element.iter(qn("w:t")))
        want = _norm_text(b.get("text"))
        have_n = _norm_text(have)
        if want == have_n or want == _norm_core(have):
            return el
        if not by_text_fallback:
            raise SystemExit(
                f"idx 漂移：IR 第 {b['idx']} 块「{(b.get('text') or '')[:20]}」与成品"
                f"对应位置「{have[:20]}」文字不一致，IR 与成品不配套，请人工检查")
    if not by_text_fallback:
        return None
    # 文字全文定位：只认完全一致的段落，错位或被改写的块都定位不到，
    # 定位不到就不删（宁漏勿错——漏删的致谢块会在 removed 的零命中
    # 点名里被人工检查兜住）
    want = _norm_text(b.get("text"))
    if not want:
        return None
    for cand in by_idx.values():
        cand_text = (cand.text if isinstance(cand, Paragraph)
                     else "".join(t.text or "" for t in cand._element.iter(qn("w:t"))))
        if _norm_text(cand_text) == want:
            return cand
    return None


def _part_root(part):
    """部件的 XML 根。python-docx 1.2 把脚注/尾注/批注部件注册成了 XmlPart，
    保存时按 _element 序列化——只改 _blob 会被旧内容悄悄覆盖。"""
    if hasattr(part, "_element"):
        return part._element
    return parse_part_xml(part.blob)


def _store_part_root(part, root):
    """把改完的部件写回：XmlPart 改的就是 _element，无需落盘动作。"""
    if hasattr(part, "_element"):
        return
    part._blob = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("ir")
    ap.add_argument("ruleset")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    rules = json.loads(Path(a.ruleset).read_text(encoding="utf-8"))
    ir = json.loads(Path(a.ir).read_text(encoding="utf-8"))
    sub = rules.get("submission") or {}
    if not sub.get("anonymous"):
        raise SystemExit("该刊不要求匿名，无需生成匿名稿")

    doc = Document(a.docx)
    by_idx = dict(_iter_children(doc))
    removed = []

    # ---- 0. 作者信息页：整页删除（该页段落带专用样式，不必猜范围）----
    page_paras = [p for p in doc.paragraphs if p.style is not None and p.style.name == "作者信息页"]
    if page_paras:
        removed.append(f"作者信息页（{len(page_paras)} 行）")
        for p in page_paras:
            p._p.getparent().remove(p._p)

    # ---- 1. 前置区：保留标题与副标题，其余（作者行、单位行）删掉 ----
    # 副标题的排版约定是破折号（——/—/-/–）打头；单位/地址行常被写成
    # 「- 北京大学经济学院，…」甚至「—北京大学经济学院，…」，当副标题静默
    # 保留会把单位留在匿名稿里（F10 修了半角连字符，F4-6 补上破折号）。
    # 所以身份扫描放在副标题豁免之前：命中身份词的破折号行也删。
    fm_seen = 0
    first_blk = ir["blocks"][0] if ir["blocks"] else None
    for b in ir["blocks"]:
        if b["type"] != "front_matter":
            continue
        fm_seen += 1
        el = by_idx.get(b["idx"])
        if el is None or not isinstance(el, Paragraph):
            continue
        txt = el.text.strip()
        if fm_seen == 1:
            if first_blk is None or first_blk.get("idx") == b["idx"]:
                continue  # 本块就是文档首块：正常情况下的标题，无条件保留
            if first_blk.get("sdt_pos") is None:
                # 不是文档首块、首块也不是内容控件里的段落（罕见版式：表格/
                # 书签排在标题之前）——标题身份存疑，保留并点名请人工检查
                if txt:
                    removed.append(f"未识别：前置区首段不在文档首位"
                                   f"（「{txt[:30]}」），请人工检查")
                continue
            # 标题藏在内容控件（sdt）里：作者行顶成了「第一个前置区段」，
            # 不能当标题放行（F4-4），落到下面按普通前置区行处理
        if _identity_hit(txt):
            if txt:
                note = ("（破折号行命中身份词）"
                        if txt.startswith(("——", "—", "-", "–")) else "")
                removed.append(f"作者署名{note}：{txt[:40]}")
            _require_paired(by_idx, b, el)   # 删前核对：idx 错位即拒绝产出（F4-5）
            el._p.getparent().remove(el._p)
            continue
        if txt.startswith(("——", "—", "-", "–")):
            continue  # 破折号打头但没有身份词：当副标题保留
        if txt:
            removed.append(f"作者署名：{txt[:40]}")
        _require_paired(by_idx, b, el)
        el._p.getparent().remove(el._p)

    # ---- 2. 补录注入的单位/作者简介/基金/致谢块 ----
    # 作者、单位、简介、基金四类本身就是身份信息，凡要求匿名一律删除；
    # 致谢是否删除由规则集声明（有的刊允许保留，只要不含致谢人姓名）
    #
    # 安全优先：规则集没声明"作者""单位"时也删。漏删是身份泄露（不可挽回），
    # 多删是基金信息缺失（可从署名稿拿回来，而且署名稿本来就完整保留）。
    # 但多删这件事必须在报告里说清，不能悄悄多删。
    listed = set(sub.get("anonymize_blocks") or [])
    SAFETY_TYPES = {"authors", "affiliations", "author_note", "funding_statement",
                    "corresponding_author",
                    # apply_details 补录注入的英文作者/单位块，与中文署名
                    # 同属身份信息：拼音名、英文单位同样能反查作者（F4-2）
                    "english_authors", "english_affiliations"}
    drop_types = SAFETY_TYPES | listed

    # 规则集声明了、工具却不认识的类型：记下来让用户人工检查，
    # 不能静默跳过——静默跳过等于身份信息留在稿子里
    unknown = sorted(listed - SAFETY_TYPES - {"acknowledgement"})
    for t in unknown:
        removed.append(f"未识别：规则集要求匿名化「{t}」，本工具不认识该类型，请人工检查")

    # 致谢小节：删标题还要删标题下面的正文，否则会留下一段孤立的「感谢……」——
    # 既泄露了致谢内容，又让成品多出一段没头没尾的文字。
    # parse_docx 从不产出 acknowledgement 类型（「致谢」标题被 UNNUMBERED_H1
    # 判成 heading1），所以标题只能按文本命中识别，不能再要求块类型在
    # drop_types 里——否则致谢小节永远删不掉（F4-1）。标题词必须独占一段
    # （可带编号前缀与标点）：正文里出现「致谢」字样的段落不能连坐。
    ack_idx = set()
    for pos, b in enumerate(ir["blocks"]):
        head = ACK_HEAD_RE.fullmatch((b.get("text") or "").strip())
        if not head:
            continue
        ack_idx.add(b["idx"])
        for nxt in ir["blocks"][pos + 1:]:
            if (nxt["type"].startswith("heading")
                    or nxt["type"] in ("references_heading", "appendix_heading",
                                       "acknowledgement")):
                break
            if nxt["type"] in ("body", "table", "table_note", "table_caption",
                               "figure", "figure_caption", "equation"):
                ack_idx.add(nxt["idx"])

    for b in ir["blocks"]:
        # 核对只对「要按它删东西」的块做：本工具只删前置区、安全类型与
        # 致谢小节，其余块（渲染器会归一标签/分隔符/编号、重排表注层序，
        # 元素文字与 IR 不同属既定行为）不经过核对，也不该被误伤（F4-5）
        # 的反面——把没动的块误判成漂移、拒绝整份交付。
        if (b["type"] != "front_matter" and b["type"] not in drop_types
                and b["idx"] not in ack_idx):
            continue
        # 致谢成员块可能被表注层序重排搬动：按文字全文兜底定位
        el = _droppable(by_idx, b, by_text_fallback=b["idx"] in ack_idx)
        if b["idx"] in ack_idx:
            if isinstance(el, Paragraph):
                removed.append(f"致谢小节：{el.text.strip()[:30]}")
                el._p.getparent().remove(el._p)
            continue
        if b["type"] not in drop_types or b["type"] == "front_matter":
            continue
        if isinstance(el, Paragraph):
            if b["type"] == "acknowledgement" or ACK_HEAD_RE.fullmatch(el.text.strip()):
                removed.append(f"致谢小节：{el.text.strip()[:30]}")
            else:
                # 规则集没要求、但为确保匿名一并删的，标注清楚
                mark = "" if b["type"] in listed else "（规则集未声明，为确保匿名一并移除）"
                removed.append(f"{b['type']}{mark}：{el.text.strip()[:40]}")
            el._p.getparent().remove(el._p)

    # 规则集声明了、IR 里却一个对应块都没有的类型：点名让用户人工检查。
    # parse_docx 现阶段产不出 funding_statement 这类块（基金在脚注里，
    # 由第 4 步按文本清），声明了却零命中不该静默——静默等于编辑部看到的
    # 匿名稿里可能还留着这类内容（F4-1）。
    for t in sorted(listed - {"acknowledgement"}):
        if not any(b["type"] == t for b in ir["blocks"]):
            removed.append(f"声明了 {t} 但全文无该类型块：本工具按类型找不到"
                           "可删对象，若稿件含此类内容请人工检查（基金/简介类"
                           "脚注已按文本另删，见上方「脚注」条目）")

    # ---- 3. 页眉页脚：全部节 × 默认/首页/偶数，含页眉页脚内的表格 ----
    # 首页脚注是补录注入作者信息的通道，整段清；其余页眉页脚里，单位行、
    # 邮箱、通讯作者这类命中身份词的段落清掉（运行头、页码等中性内容保留，
    # 页码是域代码没有 w:t，本来就不会命中）。
    for sec in doc.sections:
        stories = ((sec.header, "页眉", False),
                   (sec.first_page_header, "首页页眉", False),
                   (sec.even_page_header, "偶数页页眉", False),
                   (sec.footer, "页脚", False),
                   (sec.first_page_footer, "首页脚注", True),
                   (sec.even_page_footer, "偶数页页脚", False))
        for hf, label, blanket in stories:
            if hf.is_linked_to_previous:
                # 没有自己的部件：内容继承自前节，前节处理时已覆盖；
                # 此时取 _element 会凭空造一个空部件出来
                continue
            for p in hf._element.iter(qn("w:p")):
                text = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
                if not text:
                    continue
                if blanket or _identity_hit(text):
                    removed.append(f"{label}：{text[:40]}")
                    for t in p.iter(qn("w:t")):
                        t.text = ""

    # ---- 4. 正文脚注 / 尾注里的作者信息（原稿自带的作者简介常在这里）----
    # 脚注部件不在 python-docx 的常规对象模型里，按关系取到部件的 XML 再改
    for rel in list(doc.part.rels.values()):
        kind = ("footnote" if rel.reltype.endswith("/footnotes")
                else "endnote" if rel.reltype.endswith("/endnotes") else None)
        if kind is None:
            continue
        label = "脚注" if kind == "footnote" else "尾注"
        part = rel.target_part
        root = _part_root(part)
        killed = set()
        for note in root.findall(qn(f"w:{kind}")):
            text = "".join(t.text or "" for t in note.iter(qn("w:t")))
            if not text.strip():
                continue
            if STRONG_RE.search(text) or (UNIT_RE.search(text) and IDENT_RE.search(text)):
                removed.append(f"{label}：{text.strip()[:40]}")
                killed.add(note.get(qn("w:id")))
                for t in note.iter(qn("w:t")):
                    t.text = ""
        if killed:
            _store_part_root(part, root)
            # 正文里的引用标记一并摘掉，否则成品会留下一个空脚注号
            for ref in list(doc.element.body.iter(qn(f"w:{kind}Reference"))):
                if ref.get(qn("w:id")) in killed:
                    run = ref.getparent()
                    run.getparent().remove(run)

    # ---- 5. 批注：正文连作者整删，匿名稿不留任何批注 ----
    # 只把批注作者改成「匿名」挡不住正文泄露——批注里常写着
    # 「这个数据找张三要（zs@xxx.edu.cn）」。清掉部件内容，
    # 还要摘掉正文里的引用标记，否则成品留着指向空批注的痕迹。
    for rel in list(doc.part.rels.values()):
        if not rel.reltype.endswith("/comments"):
            continue
        part = rel.target_part
        root = _part_root(part)
        killed = False
        for c in root.findall(qn("w:comment")):
            removed.append(f"批注：{''.join(c.itertext()).strip()[:40]}")
            root.remove(c)
            killed = True
        if killed:
            _store_part_root(part, root)
    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
        for el in list(doc.element.body.iter(qn(tag))):
            el.getparent().remove(el)
    for ref in list(doc.element.body.iter(qn("w:commentReference"))):
        run = ref.getparent()
        run.getparent().remove(run)

    doc.save(a.out)
    # 正文删干净不等于交付安全：修订作者、批注作者、Word 属性和 WPS 账号
    # 都可能还带着身份。匿名稿一律连这些一起清，再做一次硬性检查。
    # 清理或闸门失败时输出文件必须撤下交付名（F4-7）：留着的就是「已保存
    # 但未清元数据」或「未过闸」的版本，CLI 直跑没有应用内 .staging 中转，
    # 谁拿到谁就当成品。失败即删除，重新跑一遍就是干净现场。
    try:
        meta_removed = sanitize_docx(a.out, scrub_people=True)
        assert_clean(a.out, strict_people=True)
    except BaseException:
        Path(a.out).unlink(missing_ok=True)
        raise
    print(json.dumps({"removed": removed, "metadata_removed": meta_removed},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
