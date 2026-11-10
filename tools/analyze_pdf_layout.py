"""从已发表论文 PDF 反推期刊版式，并与规则集自动对照。

已发表论文是最好的版式证据——它展示的是编辑部实际执行的排版。
两个必须处理的问题：
1. 中文期刊 PDF 常被整体缩印，直接读出的字号偏小 → 用 --anchor-pt 做缩放校准
2. 官网投稿指南常滞后于实际排版 → 用 --compare 与规则集逐项对照，自动列出偏差

用法：
    python tools/analyze_pdf_layout.py paper.pdf
    python tools/analyze_pdf_layout.py paper.pdf --anchor-pt 12
    python tools/analyze_pdf_layout.py p1.pdf p2.pdf --compare rules/journal/ciejournal.json
"""

import argparse
import json
import re
import statistics
from collections import Counter

import fitz

CN_SIZES = [
    (42, "初号"), (36, "小初"), (26, "一号"), (24, "小一"), (22, "二号"),
    (18, "小二"), (16, "三号"), (15, "小三"), (14, "四号"), (12, "小四"),
    (10.5, "五号"), (9, "小五"), (7.5, "六号"), (6.5, "小六"), (5.5, "七号"),
]

PAPERS = {
    "A4": (210, 297), "A5": (148, 210), "B5": (176, 250),
    "16开": (185, 260), "大16开": (210, 285), "Letter": (216, 279),
}

PAT_H1 = re.compile(r"^[一二三四五六七八九十]{1,3}、")
PAT_NOTE = re.compile(r"^(注|资料来源|数据来源)\s*[:：]")


def nearest_cn_size(pt):
    best = min(CN_SIZES, key=lambda x: abs(x[0] - pt))
    return best[1] if abs(best[0] - pt) <= 0.45 else "—"


def match_paper(w, h, tol=4.0):
    for name, (pw, ph) in PAPERS.items():
        if (abs(w - pw) <= tol and abs(h - ph) <= tol) or \
           (abs(w - ph) <= tol and abs(h - pw) <= tol):
            return name
    return None


def measure(path, anchor=None):
    """返回单个 PDF 的版式指标。"""
    doc = fitz.open(path)
    r = doc[0].rect
    pw, ph = r.width / 72 * 25.4, r.height / 72 * 25.4

    size_chars = Counter()
    page_lines = []          # (页号, 基线y, 行字号)
    h1_sizes = []
    note_sizes = []
    x0 = 1e9
    x1 = -1e9
    readable = 0

    for pno, page in enumerate(doc):
        ph = page.rect.height
        for blk in page.get_text("dict").get("blocks", []):
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                # 版心宽只统计正文行：真实期刊的页眉/页脚常外凸出正文边界
                # （页码、页眉线），整页 bbox 一并累进会把 text_width_mm 撑宽
                # 数十毫米。排除整体落在每页上下 10% 带内的行。
                bb = line.get("bbox")
                if bb and (bb[3] > 0.10 * ph and bb[1] < 0.90 * ph):
                    x0, x1 = min(x0, bb[0]), max(x1, bb[2])
                txt = "".join(s.get("text", "") for s in line.get("spans", [])).strip()
                sz = max((s["size"] for s in line.get("spans", [])), default=0)
                if txt:
                    for s in line.get("spans", []):
                        if s.get("text", "").strip():
                            size_chars[round(s["size"], 1)] += len(s["text"])
                            readable += len(s["text"])
                    if PAT_H1.match(txt) and len(txt) < 30:
                        h1_sizes.append(round(sz, 1))
                    if PAT_NOTE.match(txt):
                        note_sizes.append(round(sz, 1))
                if line.get("spans"):
                    page_lines.append((pno, line["spans"][0]["origin"][1], sz))

    if readable < 200:
        return {"ok": False, "path": path, "reason": "字体编码损坏，可提取文字过少"}

    body_raw = size_chars.most_common(1)[0][0]
    scale = (body_raw / anchor) if anchor else 1.0

    # 行距：按页取「正文尺寸」行的基线，排序后取相邻差。
    # 不用块内相邻行——PDF 常把一段拆成多个块，块内采样不足会得出错值。
    by_page = {}
    for pno, origin, sz in page_lines:
        if abs(sz - body_raw) < 0.4:
            by_page.setdefault(pno, []).append(origin)
    gaps = []
    for origins in by_page.values():
        origins.sort()
        for a, b in zip(origins, origins[1:]):
            if 5 < b - a < 45:
                gaps.append(b - a)

    # 行距只统计正文尺寸的行
    return {
        "ok": True,
        "path": path,
        "page_mm": (pw, ph),
        "paper": match_paper(pw, ph),
        "scale": scale,
        "body_pt": body_raw / scale,
        "line_pt": (statistics.median(gaps) / scale) if gaps else None,
        "level1_pt": (Counter(h1_sizes).most_common(1)[0][0] / scale) if h1_sizes else None,
        "note_pt": (Counter(note_sizes).most_common(1)[0][0] / scale) if note_sizes else None,
        "text_width_mm": (x1 - x0) / scale * 25.4 / 72,
        "body_font": Counter({s: n for s, n in size_chars.items()}).most_common(1)[0][0],
    }


def summarize(metrics):
    """多篇取众数，降低单篇异常的影响。"""
    ok = [m for m in metrics if m["ok"]]
    if not ok:
        return None
    out = {"n": len(ok)}
    for key in ("body_pt", "line_pt", "level1_pt", "note_pt", "text_width_mm", "scale"):
        vals = [round(m[key], 1) for m in ok if m.get(key)]
        out[key] = Counter(vals).most_common(1)[0][0] if vals else None
    out["paper"] = Counter(m["paper"] for m in ok).most_common(1)[0][0]
    out["page_mm"] = ok[0]["page_mm"]
    out["spread"] = {k: sorted({round(m[k], 1) for m in ok if m.get(k)}) for k in ("body_pt", "line_pt")}
    return out


def compare(s, rules):
    """把实测值与规则集逐项对照，返回行列表。"""
    ty = rules["typography"]
    hd = rules["headings"]
    rows = []
    # 规则集取值的权威来源：template（官方模板）> guide（官方指南）> published（实物实测）
    authority = rules.get("verification", {}).get("typography_authority", "published")

    def add(item, guide, real, tol, unit="pt"):
        if real is None:
            rows.append((item, guide, "未测到", "—", "跳过"))
            return
        if guide is None:
            rows.append((item, "未规定", f"{real}{unit}", "—", "新增"))
            return
        diff = real - guide
        if abs(diff) <= tol:
            verdict = "一致"
        elif authority in ("template", "guide"):
            # 该刊有官方模板/明文规定，规则集以官方为准；
            # 实物差异属于编辑部排版环节的处理，不作为规则集错误
            verdict = "参考"
        else:
            verdict = "**不符**"
        sign = "+" if diff > 0 else ""
        rows.append((item, f"{guide}{unit}", f"{real}{unit}", f"{sign}{round(diff, 1)}", verdict))

    b = ty["body"]
    add("正文字号", b.get("size_pt"), s.get("body_pt"), 0.6)
    add("正文行距", b.get("line_spacing", {}).get("value_pt"), s.get("line_pt"), 1.0)
    add("一级标题", hd.get("level1", {}).get("size_pt"), s.get("level1_pt"), 0.8)
    note_guide = rules["tables"]["note_layers"][0].get("size_pt") if rules["tables"].get("note_layers") else None
    add("表注字号", note_guide, s.get("note_pt"), 0.8)

    pg = rules["page"]
    guide_w = pg["size_mm"]["width"] - pg["margin_mm"]["left"] - pg["margin_mm"]["right"]
    if s.get("text_width_mm"):
        diff = s["text_width_mm"] - guide_w
        ok = abs(diff) <= 3
        verdict = "一致" if ok else ("参考" if authority in ("template", "guide") else "**不符**")
        rows.append(("版心宽", f"{guide_w:.1f}mm", f"{s['text_width_mm']:.1f}mm",
                     f"{'+' if diff > 0 else ''}{diff:.1f}", verdict))
    paper_label = s.get("paper") or "非标准"
    if s.get("page_mm"):
        same = s["paper"] == pg["size"]
        # 纸张不同不算缺陷：规则集按 A4 渲染以便投稿，原刊成品开本可能不同
        rows.append(("纸张", pg["size"],
                     f"{paper_label} {s['page_mm'][0]:.0f}×{s['page_mm'][1]:.0f}mm",
                     "—", "一致" if same else "参考"))
    return rows


def print_compare(s, rules, rows):
    j = rules["journal"]
    print(f"\n-- 与规则集对照：{j['name_zh']}（{rules['verification']['status']}） --")
    print(f"   样本 {s['n']} 篇，缩放系数 {s['scale']:.3f}")
    if s.get("spread") and len(s["spread"].get("body_pt", [])) > 1:
        print(f"   ⚠ 各篇正文字号不一致：{s['spread']['body_pt']}")
    w = [12, 14, 20, 8, 8]
    head = ["项目", "规则集/指南", "实测", "偏差", "结论"]
    print("   " + "".join(h.ljust(x) for h, x in zip(head, w)))
    for row in rows:
        print("   " + "".join(str(v).ljust(x) for v, x in zip(row, w)))
    bad = [r[0] for r in rows if "不符" in r[4]]
    ref = [r[0] for r in rows if r[4] == "参考" and r[1] != "A4"]
    print()
    if bad:
        print(f"   共 {len(bad)} 项不符：{'、'.join(bad)}")
        print("   → 建议以实测值为准修改规则集，并在 verification.cross_check 中记录")
    else:
        print("   规则集与实物排版吻合（无不符项）")
    if ref:
        authority = rules.get("verification", {}).get("typography_authority", "published")
        src = {"template": "官方模板", "guide": "官方指南"}.get(authority, "官方规定")
        print(f"   其中 {'、'.join(ref)} 与实物有差异，但规则集取自{src}，以官方为准；")
        print("   该差异属编辑部排版环节的处理，不影响投稿要求")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--anchor-pt", type=float, default=None,
                    help="官方已知的正文字号，用于反推 PDF 缩放系数")
    ap.add_argument("--compare", default=None, metavar="RULESET",
                    help="与该期刊规则集 JSON 对照，自动列出不符项")
    args = ap.parse_args()

    metrics = []
    for p in args.files:
        m = measure(p, anchor=args.anchor_pt)
        if not m["ok"]:
            print(f"跳过 {p.split('/')[-1][:26]}：{m['reason']}")
            continue
        metrics.append(m)
        print(f"{p.split('/')[-1][:30]:32s} 正文 {m['body_pt']:5.1f}pt  "
              f"行距 {(m['line_pt'] or 0):5.1f}  一级 {(m['level1_pt'] or 0):5.1f}  "
              f"{m['paper'] or '非标准'}")

    if not metrics:
        print("无可用样本\n")
        return
    s = summarize(metrics)
    if args.compare:
        rules = json.load(open(args.compare, encoding="utf-8"))
        print_compare(s, rules, compare(s, rules))
    print()


if __name__ == "__main__":
    main()
