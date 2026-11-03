"""导出 docx 的版式参数（页面设置、样式定义、段落级格式），用于从官方投稿模板反推期刊规则集。

用法：python tools/inspect_docx.py <file.docx> [--max-para N]
"""

import sys
import zipfile
from xml.etree import ElementTree as ET

from pathlib import Path

# embeddable Python（_pth 模式）不会把脚本所在目录放进 sys.path——
# 这是它和常规解释器唯一的行为差异，绿色包里所有同目录 import 都会
# 因此失败。显式挂上，两种启动方式行为一致。
sys.path.insert(0, str(Path(__file__).resolve().parent))

import xmlsafe

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def q(tag):
    return W + tag


def brief(elem, keys):
    if elem is None:
        return None
    out = {}
    for k in keys:
        v = elem.get(q(k))
        if v is not None:
            out[k] = v
    return out or None


def dump_sect(doc):
    print("=== 页面设置 (sectPr) ===")
    for sect in doc.iter(q("sectPr")):
        pgsz = sect.find(q("pgSz"))
        pgmar = sect.find(q("pgMar"))
        if pgsz is not None:
            w = int(pgsz.get(q("w"), 0)) / 56.7
            h = int(pgsz.get(q("h"), 0)) / 56.7
            print(f"  纸张: {w:.1f} x {h:.1f} mm")
        if pgmar is not None:
            conv = lambda v: f"{int(v) / 56.7:.1f}mm"
            print(
                "  页边距: 上{} 下{} 左{} 右{} 装订线{}".format(
                    conv(pgmar.get(q("top"), 0)),
                    conv(pgmar.get(q("bottom"), 0)),
                    conv(pgmar.get(q("left"), 0)),
                    conv(pgmar.get(q("right"), 0)),
                    conv(pgmar.get(q("gutter"), 0)),
                )
            )
        break


def dump_styles(sty):
    print("=== 默认字体 (docDefaults) ===")
    dd = sty.find(q("docDefaults"))
    if dd is not None:
        rpr = dd.find(f"{q('rPrDefault')}/{q('rPr')}")
        if rpr is not None:
            rf = brief(rpr.find(q("rFonts")), ["ascii", "hAnsi", "eastAsia"])
            sz = rpr.find(q("sz"))
            print(f"  字体: {rf}  字号: {sz.get(q('val')) if sz is not None else None} (半磅)")
    print("=== 具名样式 ===")
    for s in sty.findall(q("style")):
        if s.get(q("type")) != "paragraph":
            continue
        sid = s.get(q("styleId"))
        nm = s.find(q("name"))
        name = nm.get(q("val")) if nm is not None else ""
        rpr = s.find(q("rPr"))
        ppr = s.find(q("pPr"))
        parts = []
        if rpr is not None:
            rf = brief(rpr.find(q("rFonts")), ["eastAsia", "ascii"])
            if rf:
                parts.append(f"font={rf}")
            sz = rpr.find(q("sz"))
            if sz is not None:
                parts.append(f"sz={int(sz.get(q('val'))) / 2}pt")
            if rpr.find(q("b")) is not None:
                parts.append("bold")
        if ppr is not None:
            sp = brief(ppr.find(q("spacing")), ["line", "lineRule", "before", "after"])
            if sp:
                line = sp.get("line")
                rule = sp.get("lineRule", "auto")
                if line and rule == "exact":
                    parts.append(f"行距={int(line) / 20}pt固定")
                elif line:
                    parts.append(f"行距={int(line) / 240:.2f}倍")
            jc = brief(ppr.find(q("jc")), ["val"])
            if jc:
                parts.append(f"对齐={jc['val']}")
            ind = brief(ppr.find(q("ind")), ["firstLine", "firstLineChars", "left"])
            if ind:
                parts.append(f"缩进={ind}")
        print(f"  [{sid}] {name} | " + " ".join(parts))


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def dump_paras(doc, max_para):
    print("=== 段落 ===")
    body = doc.find(q("body"))
    idx = 0
    for p in body.iter(q("p")):
        idx += 1
        if idx > max_para:
            print(f"  ... 其余段落省略，共 {sum(1 for _ in body.iter(q('p')))} 段")
            break
        ppr = p.find(q("pPr"))
        style = ""
        fmt = []
        if ppr is not None:
            ps = ppr.find(q("pStyle"))
            if ps is not None:
                style = ps.get(q("val"))
            sp = brief(ppr.find(q("spacing")), ["line", "lineRule", "before", "after"])
            if sp:
                line = sp.get("line")
                if line and sp.get("lineRule") == "exact":
                    fmt.append(f"行距{int(line) / 20}pt")
                elif line:
                    fmt.append(f"行距{int(line) / 240:.2f}x")
            jc = brief(ppr.find(q("jc")), ["val"])
            if jc:
                fmt.append(jc["val"])
            ind = brief(ppr.find(q("ind")), ["firstLineChars", "firstLine", "left"])
            if ind:
                fmt.append(f"ind{ind}")
        rpr = p.find(f"{q('pPr')}/{q('rPr')}") or p.find(q("rPr"))
        if rpr is not None:
            rf = brief(rpr.find(q("rFonts")), ["eastAsia"])
            if rf:
                fmt.append(f"{rf.get('eastAsia')}")
            sz = rpr.find(q("sz"))
            if sz is not None:
                fmt.append(f"{int(sz.get(q('val'))) / 2}pt")
            if rpr.find(q("b")) is not None:
                fmt.append("粗")
        txt = para_text(p).strip()
        if not txt:
            continue
        print(f"  {idx:3d} [{style or '-'}] {' '.join(fmt)} | {txt[:46]}")


def main():
    path = sys.argv[1]
    max_para = 80
    if "--max-para" in sys.argv:
        max_para = int(sys.argv[sys.argv.index("--max-para") + 1])
    z = zipfile.ZipFile(path)
    doc = xmlsafe.fromstring(z.read("word/document.xml"))
    sty = xmlsafe.fromstring(z.read("word/styles.xml"))
    dump_sect(doc)
    print()
    dump_styles(sty)
    print()
    dump_paras(doc, max_para)


if __name__ == "__main__":
    main()
