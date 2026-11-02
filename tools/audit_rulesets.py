"""规则集体检：把 23 本期刊的规则集逐项过一遍，输出可复核的体检报告。

只做检查，不改数据。检查项分四类：

1. 权威顺序（用户的硬要求）——官方文件有明文的项，不得被实物实测值覆盖
2. 字号可取性——Word 的字号是半磅刻度，带小数的值交付时会被截断
3. 出处完整性——字号字段有没有 source / measured / 推定标记
4. 产物新鲜度——docx/pdf/html 是否比 json 旧

用法：
    python tools/audit_rulesets.py                    # 打印摘要
    python tools/audit_rulesets.py --json out.json    # 落盘结构化结果
    python tools/audit_rulesets.py --md out.md        # 落盘 markdown 表
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
JOURNAL_DIR = ROOT / "rules" / "journal"
SOURCE_DIR = ROOT / "rules" / "sources"

# 字号类字段名（值单位 pt）。不在此列的数值字段不参与刻度检查。
SIZE_KEYS = {
    "size_pt", "font_size_pt", "caption_size_pt", "value_pt",
    "references_size_pt", "note_size_pt",
}

# 中文经管刊常备构件。缺失不算错，但要在报告里出现。
COMMON_BLOCKS = [
    "title", "authors", "affiliations", "corresponding_author",
    "abstract", "keywords", "classification_code", "document_code",
    "received_date", "funding_statement", "author_note",
    "body", "endnotes", "appendix", "acknowledgement",
    "references", "data_availability", "conflict_of_interest",
]

OFFICIAL_MARKERS = ("官方", "明文", "模板", "体例", "指南", "须知", "规范", "稿约")
MEASURED_MARKERS = ("实测", "measured", "成品", "实物")


def walk_sizes(node, path=()):
    """产出 (路径, 值) —— 只含字号类字段。"""
    if isinstance(node, dict):
        for k, v in node.items():
            if k in ("measured", "measured_raw", "size_pt_raw"):
                continue                    # 实测原值只作证据保留，不参与刻度检查
            if k in SIZE_KEYS and isinstance(v, (int, float)) and not isinstance(v, bool):
                yield path + (k,), float(v)
            else:
                yield from walk_sizes(v, path + (k,))
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from walk_sizes(v, path + (i,))


def walk_string_sizes(node, path=()):
    """产出 (路径, 字符串值) —— 命中 SIZE_KEYS 但值是字符串的字段。

    这类值既不参与刻度检查也没法归整，静默跳过等于假装没看见，必须点名报警。
    """
    if isinstance(node, dict):
        for k, v in node.items():
            if k in ("measured", "measured_raw", "size_pt_raw"):
                continue
            if k in SIZE_KEYS and isinstance(v, str):
                yield path + (k,), v
            else:
                yield from walk_string_sizes(v, path + (k,))
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from walk_string_sizes(v, path + (i,))


def get_path(node, path):
    for p in path:
        node = node[p]
    return node


def size_provenance(rules, path):
    """判断某个字号值的出处等级：official / measured / inferred / unknown。"""
    parent_path = path[:-1]
    try:
        parent = get_path(rules, parent_path)
    except (KeyError, IndexError, TypeError):
        return "unknown"
    if not isinstance(parent, dict):
        return "unknown"
    text = " ".join(
        str(parent.get(k) or "") for k in ("source", "note", "size_source",
                                           "caption_size_source", "font_size_pt_source")
    )
    has_official = any(m in text for m in OFFICIAL_MARKERS)
    has_measured = bool(parent.get("measured")) or any(m in text for m in MEASURED_MARKERS)
    if has_official and not has_measured:
        return "official"
    if has_official and has_measured:
        return "official+measured"
    if has_measured:
        return "measured"
    return "unknown"


def audit_one(path: Path):
    rules = json.loads(path.read_text(encoding="utf-8"))
    jid = rules["journal"]["id"]
    typo = rules.get("typography", {})
    ver = rules.get("verification", {})

    sizes = list(walk_sizes({k: rules.get(k) for k in
                             ("typography", "headings", "tables", "figures", "numbering")}))
    str_sizes = [(".".join(str(p) for p in sp), v) for sp, v in
                 walk_string_sizes({k: rules.get(k) for k in
                                    ("typography", "headings", "tables", "figures", "numbering")})]
    off_grid = []
    prov = {"official": 0, "official+measured": 0, "measured": 0, "unknown": 0}
    for sub_path, val in sizes:
        if abs(val * 2 - round(val * 2)) > 1e-9:
            off_grid.append((".".join(str(p) for p in sub_path), val))
        prov[size_provenance({"typography": rules.get("typography"),
                              "headings": rules.get("headings"),
                              "tables": rules.get("tables"),
                              "figures": rules.get("figures"),
                              "numbering": rules.get("numbering")}, sub_path)] += 1

    block_types = [b.get("type") for b in rules.get("blocks", [])]
    # 正文里没有作者行，有三种正常情形，都不能算缺构件：
    #   ① 走独立作者页（separate_page = true）
    #   ② 作者信息只在投稿系统里填（personal_fields 的 placement = "system"）
    #   ③ 只收集不落版（placement = "none"）
    sep_page = bool((rules.get("submission") or {}).get("author_page", {}).get("separate_page"))
    author_pf = [f for f in (rules.get("personal_fields") or [])
                 if f.get("id") in ("author_name", "affiliation", "corresponding_author", "email")]
    placements = {f.get("placement") for f in author_pf}
    byline_by_design = sep_page or (bool(author_pf) and placements <= {"system", "none"})
    byline = {"authors", "affiliations", "corresponding_author"}
    missing_blocks = [
        b for b in COMMON_BLOCKS
        if b not in block_types
        and not (byline_by_design and b in byline)
        and not (rules["journal"].get("language") != "zh" and b in ("classification_code", "document_code", "received_date"))
    ]

    limits = rules.get("limits", {})
    null_limits = [k for k, v in limits.items()
                   if isinstance(v, dict) and v.get("value") is None]

    stale = []
    for ext in ("sample.docx", "sample.pdf", "spec.html"):
        art = path.with_name(f"{jid}.{ext}")
        if art.exists() and art.stat().st_mtime < path.stat().st_mtime:
            stale.append(ext)

    return {
        "id": jid,
        "name": rules["journal"].get("name_zh") or rules["journal"].get("name_en"),
        "language": rules["journal"].get("language"),
        "status": ver.get("status"),
        "authority": ver.get("typography_authority"),
        "has_official_source": (SOURCE_DIR / jid).is_dir(),
        "source_url": ver.get("source_url"),
        "size_fields": len(sizes),
        "string_sizes": str_sizes,
        "off_grid": off_grid,
        "provenance": prov,
        "missing_blocks": missing_blocks,
        "null_limits": null_limits,
        "stale_artifacts": stale,
        "byline_by_design": byline_by_design,
        "has_conflict": bool(ver.get("cross_check", {}).get("discrepancies")
                             or ver.get("conflict")),
        "assumptions": len(rules.get("assumptions", []) or []),
    }


def render_md(rows):
    out = ["| 期刊 | 状态 | 字号权威 | 官方原件 | 字号字段 | 非半磅 | 字符串字号 | 有出处 | 实测/推定 | 缺构件 | 产物过期 |",
           "|---|---|---|---|---|---|---|---|---|---|---|"]
    for r in rows:
        prov = r["provenance"]
        known = prov["official"] + prov["official+measured"] + prov["measured"]
        out.append(
            f"| {r['name']}（{r['id']}） | {r['status']} | {r['authority'] or '—'} | "
            f"{'有' if r['has_official_source'] else '无'} | {r['size_fields']} | "
            f"{len(r['off_grid'])} | {len(r['string_sizes'])} | {known}/{r['size_fields']} | "
            f"{prov['measured']} | {len(r['missing_blocks'])} | "
            f"{len(r['stale_artifacts'])} |"
        )
    return "\n".join(out)


def audit_all(jdir):
    """体检目录下全部规则集。坏文件逐个点名进 failures，不拖垮其余文件。"""
    rows, failures = [], []
    for p in sorted(Path(jdir).glob("*.json")):
        try:
            rows.append(audit_one(p))
        except Exception as e:  # noqa: BLE001 —— 一份手写坏 JSON 不能让体检整体崩溃
            failures.append((p.name, f"{type(e).__name__}: {e}"))
    return rows, failures


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json")
    ap.add_argument("--md")
    args = ap.parse_args()

    rows, failures = audit_all(JOURNAL_DIR)
    print(f"规则集 {len(rows)} 份，体检时间 {datetime.now():%Y-%m-%d %H:%M}"
          + (f"，{len(failures)} 份失败" if failures else ""))
    for name, why in failures:
        print(f"  失败：文件 {name}：{why}")
    print(render_md(rows))

    total_off = sum(len(r["off_grid"]) for r in rows)
    total_unknown = sum(r["provenance"]["unknown"] for r in rows)
    total_fields = sum(r["size_fields"] for r in rows)
    stale = [r["id"] for r in rows if r["stale_artifacts"]]
    print(f"\n字号字段合计 {total_fields}，非半磅 {total_off}，无出处 {total_unknown}/{total_fields}")
    bad_type = [(r["id"], s) for r in rows for s in r["string_sizes"]]
    if bad_type:
        print(f"\n{len(bad_type)} 处字号是字符串（无法参与刻度检查）：")
        for jid, (path, val) in bad_type:
            print(f"  {jid}: {path} = {val!r}")
    if stale:
        print("产物过期：", "、".join(stale))

    if args.json:
        Path(args.json).write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    if args.md:
        Path(args.md).write_text(render_md(rows) + "\n", encoding="utf-8")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
