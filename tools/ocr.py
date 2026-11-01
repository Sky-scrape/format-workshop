"""PDF 取文：优先用文字层，没有文字层才走 OCR。

为什么要先判文字层
--------------------------------------------------------------------
文字层是 PDF 里原生的字，逐字精确，还能拿到字号——用它零误差，且不花时间。
OCR 是"看图猜字"，一定有错字，只是错多错少。先取文字层，是让大多数情况
根本不需要猜。

什么时候算"没有文字层"
--------------------------------------------------------------------
不是"一个字符都没有"才算。真实扫描件常有页眉页码是后加的（带文字层），
正文却是图片。所以按**每页有效字符数**判：低于阈值就单独 OCR 那一页，
其余页仍用文字层。混排 PDF 很常见，整本交给 OCR 是浪费也是降准。

行合并
--------------------------------------------------------------------
RapidOCR 返回的是一堆检测框，同一视觉行常被切成 3–4 段。不合并就得到
"正文采用 / 小五号 / 宋体" 这种碎句，正则匹配全废、模型也容易读串。
这里按纵坐标聚行、按横坐标排序，再按标点判断是否续段。

用法
----
    from ocr import extract_pdf_text
    text, meta = extract_pdf_text("稿约.pdf")

meta 里 method 为 text_layer / ocr / mixed，供上游在报告里如实标注。
"""

import re
import sys
from pathlib import Path

DEFAULT_DPI = 300
# 一页有效字符（去空白）少于此数，判为"正文是图"，该页走 OCR。
# 20 是经验值：页眉页脚通常十几个字，不足以拼出稿约要求。
MIN_CHARS_PER_PAGE = 20

_ENGINE = None
_ENGINE_TRIED = False

# 全角拉丁字母与数字转半角。RapidOCR 在混排时常吐出 "Rｅvｉｅｗ" 这种，
# 不转的话后面所有正则都匹配不上。中文标点不动。
_FW_MAP = {**{i: i - 0xFEE0 for i in range(0xFF10, 0xFF1A)},   # ０-９
           **{i: i - 0xFEE0 for i in range(0xFF21, 0xFF3B)},   # Ａ-Ｚ
           **{i: i - 0xFEE0 for i in range(0xFF41, 0xFF5B)},   # ａ-ｚ
           0xFF0D: 0x2D,                                        # － → -
           0x3000: 0x20}                                        # 全角空格

# OCR 会把「300」拆成「3 0 0」、「300 字」读成两个词。中文排版里数字与
# 汉字之间本来也不留空格，所以这几类空格一律压掉——只在 OCR 路径用，
# 不能碰原生文字层，那是作者原本的排版。
_SQUASH = [
    (re.compile(r"(?<=[0-9])[ \t]+(?=[0-9])"), ""),
    (re.compile(r"(?<=[0-9])[ \t]+(?=[\u4e00-\u9fff])"), ""),
    (re.compile(r"(?<=[\u4e00-\u9fff])[ \t]+(?=[0-9])"), ""),
    (re.compile(r"(?<=[\u4e00-\u9fff])[ \t]+(?=[\u4e00-\u9fff])"), ""),
]


def normalize(s):
    return (s or "").translate(_FW_MAP)


def squash(s):
    for pat, rep in _SQUASH:
        s = pat.sub(rep, s)
    return re.sub(r"[ \t]{2,}", " ", s).strip()


def engine():
    """加载 OCR 引擎。加载失败返回 None，由调用方决定怎么提示。"""
    global _ENGINE, _ENGINE_TRIED
    if _ENGINE is not None or _ENGINE_TRIED:
        return _ENGINE
    _ENGINE_TRIED = True
    try:
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    except Exception:  # noqa: BLE001
        _ENGINE = None
    return _ENGINE


def available():
    return engine() is not None


def engine_name():
    return "RapidOCR（本地 ONNX）" if engine() is not None else "未安装"


# ---------------------------------------------------------------- 行重建

def _bbox(points):
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def group_lines(boxes):
    """把检测框聚成视觉行，返回按阅读顺序排好的行文本列表。"""
    items = []
    for b in boxes or []:
        try:
            x0, y0, x1, y1 = _bbox(b[0])
            txt = str(b[1])
        except (IndexError, TypeError):
            continue
        if not txt.strip():
            continue
        items.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1,
                      "cy": (y0 + y1) / 2, "h": max(1.0, y1 - y0), "text": txt})
    if not items:
        return []

    items.sort(key=lambda d: d["cy"])
    lines, cur = [], [items[0]]
    for it in items[1:]:
        avg_h = sum(x["h"] for x in cur) / len(cur)
        if abs(it["cy"] - cur[0]["cy"]) <= 0.6 * avg_h:
            cur.append(it)
        else:
            lines.append(cur)
            cur = [it]
    lines.append(cur)

    out = []
    for ln in lines:
        ln.sort(key=lambda d: d["x0"])
        buf = ""
        prev = None
        for it in ln:
            if prev is not None:
                gap = it["x0"] - prev["x1"]
                # 中文之间不补空格；间隔明显（超过半个字高）才当空格
                if gap > 0.5 * prev["h"] and not _is_cjk_edge(prev["text"], it["text"]):
                    buf += " "
            buf += it["text"]
            prev = it
        out.append(normalize(buf).strip())
    return [x for x in out if x]


def _is_cjk_edge(left, right):
    a = left.rstrip()[-1:] or ""
    b = right.lstrip()[:1] or ""
    return bool(a and b and _is_cjk(a) and _is_cjk(b))


def _is_cjk(ch):
    return "\u4e00" <= ch <= "\u9fff"


_TERM = "。！？；：!?;:"
_LEAD_MARK = re.compile(r"^\s*([（(]?\d+[）).、]|[一二三四五六七八九十]+、|[（(][一二三四五六七八九十]+[）)]|[-•·*])")


def merge_paragraphs(lines):
    """把同一段被折行切开的视觉行接回去。

    只在"上一行没有句末标点、下一行不是新的编号项"时接。稿约里几乎每行
    都是独立条目，接错了会把两条要求粘成一条，所以判据宁可保守。
    """
    out = []
    for ln in lines:
        if out and out[-1] and out[-1][-1] not in _TERM and not _LEAD_MARK.match(ln):
            out[-1] += ln
        else:
            out.append(ln)
    return out


# ---------------------------------------------------------------- 主流程

def _page_needs_ocr(text):
    return len(re.sub(r"\s", "", text or "")) < MIN_CHARS_PER_PAGE


def _ocr_page(page, dpi):
    eng = engine()
    if eng is None:
        return "", 0
    pix = page.get_pixmap(dpi=dpi)
    img = pix.tobytes("png")
    try:
        result, _ = eng(img)
    except Exception:  # noqa: BLE001
        return "", 0
    lines = [squash(x) for x in merge_paragraphs(group_lines(result))]
    return "\n".join(x for x in lines if x), len(result or [])


def extract_pdf_text(path, *, dpi=DEFAULT_DPI, max_chars=24000, log=None):
    """返回 (text, meta)。meta.method: text_layer / ocr / mixed。"""
    import fitz

    doc = fitz.open(str(path))
    pages, ocr_pages, chars = [], [], 0
    try:
        for i, page in enumerate(doc):
            raw = (page.get_text("text") or "").strip()
            used = "text_layer"
            if _page_needs_ocr(raw):
                txt, n = _ocr_page(page, dpi)
                if txt.strip():
                    raw, used = txt, "ocr"
                    ocr_pages.append(i + 1)
                elif not raw.strip():
                    raw = ""
            pages.append(f"\n\n【第{i + 1}页】\n{raw}" if raw else "")
            chars += len(re.sub(r"\s", "", raw))
    finally:
        doc.close()

    text = re.sub(r"[ \t]+\n", "\n", "".join(pages))
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    trunc = len(text) > max_chars

    if not text:
        raise RuntimeError(
            "PDF 里既没有文字层，OCR 也没读出内容。"
            "如果是扫描件，请确认图像清晰、不是整页手写或表格截图。")

    if not ocr_pages:
        method = "text_layer"
    elif len(ocr_pages) == len(pages):
        method = "ocr"
    else:
        method = "mixed"

    meta = {"method": method, "pages": len(pages), "chars": chars,
            "ocr_pages": ocr_pages, "truncated": trunc, "dpi": dpi,
            "engine": engine_name() if ocr_pages else "文字层"}
    if log and ocr_pages:
        log(f"OCR：第 {ocr_pages} 页无文字层，已识别")
    return text[:max_chars], meta


def main():
    import argparse
    import json
    ap = argparse.ArgumentParser(description="PDF 取文（文字层优先，缺失时 OCR）")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    ap.add_argument("--chars", type=int, default=24000)
    a = ap.parse_args()

    try:
        text, meta = extract_pdf_text(a.pdf, dpi=a.dpi, max_chars=a.chars, log=print)
    except Exception as e:  # noqa: BLE001
        print(f"失败：{e}")
        return 1

    if a.out:
        Path(a.out).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)
    print("\n---")
    print(json.dumps(meta, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
