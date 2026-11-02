"""把 23 本规则集按「官方明文优先、实测校准次之」重排一遍字号。

三条规则，按顺序执行：
1. 官方文件有明文的项，值改回官方明文（本轮：中国工业经济、财贸经济表题）
2. 实测得到的带小数字号，先按 0.98 反算（PDF 排版缩放系数），再落到 0.5 磅刻度
3. 有官方依据的字号补上 source，出处不清的标明是推定

只改数据，不动渲染器。运行前已有 /tmp/rules_before 快照可回溯。
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JDIR = ROOT / "rules" / "journal"
SCALE = 0.98          # PDF 排版缩放系数：实测值 × 0.98 ≈ 标称值
LOG = []
TYPE_WARNINGS = []    # 命中 SIZE_KEYS 但值是字符串的字段——没法归整，必须点名报警


def snap(v: float) -> float:
    """反算缩放后落到 0.5 磅刻度（Word 字号的真实精度）。"""
    return round(round(v / SCALE * 2) / 2, 1)


def log(jid, path, old, new, why):
    LOG.append((jid, path, old, new, why))


SIZE_KEYS = {"size_pt", "font_size_pt", "caption_size_pt", "value_pt"}


def walk(node, path, fn):
    """对每个字号字段调用 fn(container, key, value, path)。"""
    if isinstance(node, dict):
        for k, v in list(node.items()):
            if k in ("measured", "measured_raw", "size_pt_raw"):
                continue                        # 实测原值必须原样保留
            if k in SIZE_KEYS and isinstance(v, (int, float)) and not isinstance(v, bool):
                fn(node, k, float(v), path + (k,))
            else:
                if k in SIZE_KEYS and isinstance(v, str):
                    TYPE_WARNINGS.append((".".join(str(p) for p in path + (k,)), v))
                walk(v, path + (k,), fn)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            walk(v, path + (i,), fn)


# ---------------------------------------------------------------- 1. 官方明文
def official_ciejournal(r):
    """《中国工业经济》2016 投稿指南「形式范式」：39 行 × 42 字、五号宋体；
    一级标题四号仿宋居中，二级五号黑体缩进 2 字独占行，三级五号宋体缩进 2 字接排；
    表名在上、图名在下，均为五号黑体；资料来源六号宋体。
    上一版把这三项换成了实物实测值（12pt / 16pt / 12pt），与官方指南冲突，本轮改回。
    """
    t = r["typography"]
    guide = "官方投稿指南「形式范式」"
    fixes = [
        (t["body"], "size_pt", 10.5, f"{guide}：正文五号宋体"),
        (t["body"], "font_cn", "宋体", f"{guide}：五号宋体"),
        (t["body"], "chars_per_line", 42, f"{guide}：每页 39 行 × 42 字"),
        (t["body"], "lines_per_page", 39, f"{guide}：每页 39 行 × 42 字"),
        (t["abstract"], "size_pt", 10.5, f"{guide}：全文五号宋体"),
        (t["level1"], "size_pt", 14, f"{guide}：一级标题四号仿宋居中"),
        (t["level2"], "size_pt", 10.5, f"{guide}：二级标题五号黑体"),
        (t["level3"], "size_pt", 10.5, f"{guide}：三级标题五号宋体"),
        (t["level3"], "font_cn", "宋体", f"{guide}：三级标题五号宋体"),
        (r["headings"]["level1"], "size_pt", 14, f"{guide}：一级标题四号仿宋居中"),
        (r["headings"]["level2"], "size_pt", 10.5, f"{guide}：二级标题五号黑体"),
        (r["headings"]["level3"], "size_pt", 10.5, f"{guide}：三级标题五号宋体"),
        (r["tables"], "caption_size_pt", 10.5, f"{guide}：表名在上，五号黑体"),
        (r["figures"], "caption_size_pt", 10.5, f"{guide}：图名在下，五号黑体"),
        (r["tables"]["note_layers"][0], "size_pt", 7.5, f"{guide}：表注六号宋体"),
        (r["tables"]["note_layers"][1], "size_pt", 7.5, f"{guide}：资料来源六号宋体"),
        (r["figures"]["note_layers"][0], "size_pt", 7.5, f"{guide}：图注六号宋体"),
        (r["figures"]["note_layers"][1], "size_pt", 7.5, f"{guide}：资料来源六号宋体"),
    ]
    for container, key, new, why in fixes:
        old = container.get(key)
        if old != new:
            container[key] = new
            log("ciejournal", why, old, new, "官方明文优先")

    for lv in ("level1", "level2", "level3"):
        for node in (r["headings"][lv], t[lv]):
            node.pop("measured", None)
            node["source"] = f"{guide}：与本条规则一致"
    t["body"]["source"] = f"{guide}：正文五号宋体、每页 39 行 × 42 字"
    t["body"]["measured"] = {
        "size_pt": 11.76, "line_pt": 18.0, "chars_per_line": 36, "lines_per_page": 37,
        "samples": 3,
        "note": "实物为 12pt/36 字×37 行（行宽 155.9mm 恰为 42 个五号字，说明版心未变、字号被放大）。"
                "投稿按官方指南的五号排，不用实物值。",
    }
    t["abstract"]["source"] = f"{guide}：全文五号宋体"
    r["verification"]["conflict"] = {
        "kind": "official-vs-published",
        "official": "指南（2016）：正文五号宋体 10.5pt，每页 39 行 × 42 字；一级标题四号仿宋 14pt",
        "published": "2026 年实物：正文 12pt（11.76 实测），每页 37 行 × 36 字；一级标题 16pt",
        "resolution": "以官方指南为准（用户口径：官方投稿格式优先）。实物值保留在 typography.body.measured 与 verification.cross_check 中备查。",
        "risk": "若编辑部实际已按新版排版（12pt）验收，投稿稿字号会小一档；投稿前建议向编辑部确认一次。",
    }
    r["assumptions"] = [a for a in r["assumptions"] if "按实物实测值" not in a]
    r["assumptions"].append(
        "字号以 2016 版官方投稿指南为准（五号宋体、39 行 × 42 字）；实物排版已放大到 12pt，"
        "属刊物改版后指南未同步，差异见 verification.conflict"
    )


def official_cmjj(r):
    """《财贸经济》发稿排版示例：表题五号黑体，图题小五号黑体。
    上一版把表题写成小五（9pt），与官方示例相反，本轮改回五号。
    """
    tb = r["tables"]
    if tb.get("caption_size_pt") != 10.5:
        log("cmjj", "表题字号", tb.get("caption_size_pt"), 10.5, "官方发稿排版示例：表题五号黑体")
        tb["caption_size_pt"] = 10.5
    tb["caption_size_source"] = "官方《发稿排版示例》：表题「中国国内生产总值历年统计（五号黑体）」"
    r["figures"]["caption_size_source"] = "官方《发稿排版示例》：图题「电信业务总量及其占GDP的比重（小五号黑体）」"


# ---------------------------------------------------------------- 2. 刻度归整
def normalize(r, jid):
    def fn(container, key, val, path):
        if abs(val * 2 - round(val * 2)) < 1e-9:
            return                      # 已在半磅刻度上，不动
        new = snap(val)
        container[key] = new
        if "measured" not in container and "size_pt_raw" not in container:
            container["size_pt_raw"] = val
        log(jid, ".".join(str(p) for p in path), val, new, "0.98 反算 + 半磅刻度")

    for sec in ("typography", "headings", "tables", "figures"):
        walk(r.get(sec, {}), (sec,), fn)


# ---------------------------------------------------------------- 3. 补出处
GJJRYJ_SRC = {
    "title": "官方投稿模板：文章标题黑体二号（22pt）、居中、1.5 倍行距",
    "subtitle": "官方投稿模板：副标题随文章标题格式，居中",
    "body": "官方投稿模板：中文五号宋体、英文与数字五号 Times New Roman、行距 20 磅",
    "abstract": "官方投稿模板：摘要、关键词与正文同为五号宋体",
    "title_en": "官方投稿模板：英文标题四号 Times New Roman",
    "footnote": "官方稿约：脚注小五号宋体（模板未单列，按稿约）",
}
KJYJ_SRC = {
    "title": "官方《投稿格式样板》：文题小三黑体（15pt）、居中、1.2 倍行距",
    "subtitle": "官方《投稿格式样板》：副标题同文题格式",
    "body": "官方投稿须知：正文五号宋体、1.2 倍行距（与样本 docx 一致）",
    "abstract": "官方投稿须知：摘要五号宋体",
    "references": "官方《投稿格式样板》：参考文献九号宋体、1.2 倍行距",
}
EN_SRC = ("官方指南未规定字体、字号与行距；按 Elsevier/Wiley 投稿惯例取 12pt、双倍行距，"
          "属推定（见 assumptions 第 1 条）")


def add_source(jid, mapping, jdir=None):
    jdir = Path(jdir) if jdir else JDIR
    r = json.loads((jdir / f"{jid}.json").read_text(encoding="utf-8"))
    for field, text in mapping.items():
        node = r["typography"].get(field)
        if isinstance(node, dict) and "source" not in node:
            node["source"] = text
    r["verification"]["typography_provenance"] = (
        "字号逐项出处见 typography.*.source；无 source 者见 verification.cross_check 与 assumptions"
    )
    (jdir / f"{jid}.json").write_text(json.dumps(r, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main(jdir=None):
    jdir = Path(jdir) if jdir else JDIR
    failures = []
    # 1 + 2 一起走，官方改完再归整余下的小数。逐文件隔离：坏文件点名报错，
    # 其余照常处理，末尾汇总失败清单（退出码非零）。
    for p in sorted(jdir.glob("*.json")):
        try:
            r = json.loads(p.read_text(encoding="utf-8"))
            jid = r["journal"]["id"]
            if jid == "ciejournal":
                official_ciejournal(r)
            if jid == "cmjj":
                official_cmjj(r)
            normalize(r, jid)
            r.setdefault("verification", {})["size_normalization"] = (
                "字号一律落在 0.5 磅刻度上（Word 字号的真实精度）；实物实测值先按 0.98 反算排版缩放再取刻度，"
                "原始读数保留在 typography.*.size_pt_raw 或 measured 中"
            )
            p.write_text(json.dumps(r, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            failures.append((p.name, f"{type(e).__name__}: {e}"))

    for jid, mapping in (("gjjryj", GJJRYJ_SRC), ("kjyj", KJYJ_SRC)):
        if not (jdir / f"{jid}.json").exists():
            continue
        try:
            add_source(jid, mapping, jdir)
        except Exception as e:  # noqa: BLE001
            failures.append((f"{jid}.json", f"出处补写失败：{type(e).__name__}: {e}"))
    for jid in ("cer", "cwe", "ecmo", "jae", "jce", "jde"):
        if not (jdir / f"{jid}.json").exists():
            continue
        try:
            add_source(jid, {"body": EN_SRC, "title": EN_SRC}, jdir)
        except Exception as e:  # noqa: BLE001
            failures.append((f"{jid}.json", f"出处补写失败：{type(e).__name__}: {e}"))

    print(f"共 {len(LOG)} 处改动：")
    for jid, path, old, new, why in LOG:
        print(f"  {jid:10s} {path:38s} {old} → {new}   [{why}]")
    if TYPE_WARNINGS:
        print(f"\n{len(TYPE_WARNINGS)} 处字号是字符串，未归整（无法参与刻度检查）：")
        for path, val in TYPE_WARNINGS:
            print(f"  {path} = {val!r}")
    if failures:
        print(f"\n{len(failures)} 份处理失败：")
        for name, why in failures:
            print(f"  失败：文件 {name}：{why}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
