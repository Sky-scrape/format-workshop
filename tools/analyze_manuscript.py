"""对真实稿件做结构普查，输出转换器必须处理的"脏"情况统计。

用途：在写解析器之前，先量清楚稿件的混乱程度，避免按理想文档设计。

用法：python tools/analyze_manuscript.py <file.docx> [...]
"""

import sys
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

from pathlib import Path

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

import xmlsafe

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

HEADING_STYLES = {"heading 1", "heading 2", "heading 3", "heading 4", "Title", "标题 1", "标题 2", "标题 3"}


def q(tag, ns=W):
    return ns + tag


def text_of(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def run_info(rpr):
    if rpr is None:
        return None
    rf = rpr.find(q("rFonts"))
    sz = rpr.find(q("sz"))
    b = rpr.find(q("b"))
    return (
        rf.get(q("eastAsia")) if rf is not None else None,
        int(sz.get(q("val"))) / 2 if sz is not None else None,
        b is not None,
    )


def analyze(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    doc = xmlsafe.fromstring(z.read("word/document.xml"))
    body = doc.find(q("body"))

    paras = list(body.iter(q("p")))
    tables = list(body.iter(q("tbl")))

    style_of = {}
    try:
        sty = xmlsafe.fromstring(z.read("word/styles.xml"))
        for s in sty.findall(q("style")):
            sid = s.get(q("styleId"))
            nm = s.find(q("name"))
            style_of[sid] = nm.get(q("val")) if nm is not None else sid
    except KeyError:
        pass

    print("=" * 68)
    print("文件:", path.split("/")[-1].split("\\")[-1])
    print("段落:", len(paras), "| 表格:", len(tables),
          "| 图片:", len(list(doc.iter(A + "blip"))),
          "| 公式对象:", len(list(doc.iter(M + "oMath"))))

    # 样式使用情况
    style_counter = Counter()
    no_style = 0
    for p in paras:
        ppr = p.find(q("pPr"))
        ps = ppr.find(q("pStyle")) if ppr is not None else None
        if ps is None:
            no_style += 1
            style_counter["（无样式，靠直接格式）"] += 1
        else:
            sid = ps.get(q("val"))
            style_counter[style_of.get(sid, sid)] += 1
    print("\n-- 段落样式分布 --")
    for name, cnt in style_counter.most_common():
        print(f"   {cnt:5d}  {name}")
    if paras:
        print(f"   → 无样式段落占比 {no_style / len(paras):.0%}")

    # 直接格式混乱度
    font_combo = Counter()
    spacing_combo = Counter()
    for p in paras:
        ppr = p.find(q("pPr"))
        rpr = p.find(q("rPr")) if ppr is not None else p.find(q("rPr"))
        info = run_info(rpr)
        if info and (info[0] or info[1]):
            font_combo[(info[0], info[1])] += 1
        sp = ppr.find(q("spacing")) if ppr is not None else None
        if sp is not None:
            line, rule = sp.get(q("line")), sp.get(q("lineRule"))
            if line:
                if rule == "exact":
                    spacing_combo[f"{int(line) / 20}pt 固定"] += 1
                elif rule == "atLeast":
                    spacing_combo[f"最小 {int(line) / 20}pt"] += 1
                else:
                    spacing_combo[f"{int(line) / 240:.2f} 倍"] += 1

    print("\n-- 段落级直接字体覆盖（中文字体, 字号） --")
    for (f, s), cnt in font_combo.most_common(8):
        print(f"   {cnt:5d}  字体={f} 字号={s}")
    print("\n-- 段落级行距取值 --")
    for k, cnt in spacing_combo.most_common():
        print(f"   {cnt:5d}  {k}")

    # 缩进取值
    indent = Counter()
    for p in paras:
        ppr = p.find(q("pPr"))
        ind = ppr.find(q("ind")) if ppr is not None else None
        if ind is not None:
            key = tuple(sorted((k.replace(W, ""), v) for k, v in ind.attrib.items()))
            indent[key] += 1
    print("\n-- 首行缩进取值 --")
    for k, cnt in indent.most_common(6):
        print(f"   {cnt:5d}  {dict(k)}")

    # 自动编号
    numpr = len(list(body.iter(q("numPr"))))
    print(f"\n-- 自动编号引用 numPr: {numpr} 处")

    # 章节标题识别（不依赖样式，按文本模式）
    import re
    pats = {
        "一、": re.compile(r"^[一二三四五六七八九十]+、"),
        "（一）": re.compile(r"^（[一二三四五六七八九十]+）"),
        "1.": re.compile(r"^\d+[.．]\s*\S"),
        "(1)": re.compile(r"^[（(]\d+[）)]"),
    }
    hits = Counter()
    samples = {}
    for p in paras:
        t = text_of(p).strip()
        if not t or len(t) > 40:
            continue
        for name, pat in pats.items():
            if pat.match(t):
                hits[name] += 1
                samples.setdefault(name, []).append(t[:28])
                break
    print("\n-- 按文本模式识别的标题（不依赖样式） --")
    for name, cnt in hits.items():
        print(f"   {name:6s} {cnt:3d} 处   例：{' / '.join(samples[name][:2])}")

    # 表格结构
    print("\n-- 表格 --")
    for i, tbl in enumerate(tables, 1):
        rows = tbl.findall(q("tr"))
        cols = len(rows[0].findall(q("tc"))) if rows else 0
        tp = tbl.find(q("tblPr"))
        b = tp.find(q("tblBorders")) if tp is not None else None
        border_desc = "未定义"
        if b is not None:
            edges = {e.tag.replace(W, ""): e.get(q("val")) for e in b}
            border_desc = ", ".join(f"{k}={v}" for k, v in edges.items() if v != "none") or "全无边框"
        first = text_of(rows[0])[:36] if rows else ""
        print(f"   #{i} {len(rows)}行×{cols}列 | {border_desc} | 首行：{first}")

    # 脚注尾注
    for f in ("word/footnotes.xml", "word/endnotes.xml"):
        if f in names:
            root = xmlsafe.fromstring(z.read(f))
            items = [x for x in root if x.tag in (q("footnote"), q("endnote"))]
            real = [x for x in items if x.get(q("type")) in (None, "normal")]
            print(f"\n-- {f}: 共 {len(real)} 条")
            for x in real[:3]:
                print("   例：", text_of(x)[:50])

    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    for p in sys.argv[1:]:
        try:
            analyze(p)
        except Exception as e:
            print("分析失败:", p, e)
