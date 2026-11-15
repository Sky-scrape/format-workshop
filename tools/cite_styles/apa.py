"""APA 7th edition 文献条目格式化器。

只处理著者-年份制最常见的三类：期刊论文、专著、其他（网络/报告）。
字段取自 citation_ref.parse_entry，不跨越原文另造内容。
"""

from .common import (Build, add_italic_title, doi_of, initials_text, period,
                     split_name, year_of)


def _author(name):
    sur, initials = split_name(name)
    ini = initials_text(initials, spaced=True)
    return f"{sur}, {ini}".strip().rstrip(",") if ini else sur


def _authors(names):
    names = [n for n in names if n]
    if not names:
        return ""
    parts = [_author(n) for n in names]
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]}, & {parts[1]}"
    return ", ".join(parts[:-1]) + f", & {parts[-1]}"


def _journal(entry, b):
    year = year_of(entry)
    b.add(f"{_authors(entry.get('authors') or [])} ({year}). ")
    b.add(period(entry.get("title") or "") + " ")
    journal = entry.get("journal") or ""
    if journal:
        b.add(journal, italic=True)
        vol = entry.get("volume") or ""
        issue = entry.get("issue") or ""
        pages = entry.get("pages") or ""
        if vol:
            b.add(", ")
            b.add(vol, italic=True)
        if issue:
            b.add(f"({issue})")
        if pages:
            b.add(f", {pages}")
        b.add(".")
    doi = doi_of(entry)
    if doi:
        b.add(f" https://doi.org/{doi}")
    return b.done()


def _book(entry, b):
    year = year_of(entry)
    b.add(f"{_authors(entry.get('authors') or [])} ({year}). ")
    add_italic_title(b, entry.get("title") or "")
    pub = entry.get("publisher") or ""
    if pub:
        b.add(" " + period(pub))
    return b.done()


def _other(entry, b):
    year = year_of(entry)
    b.add(f"{_authors(entry.get('authors') or [])} ({year}). ")
    b.add(period(entry.get("title") or "") + " ")
    src = entry.get("source") or entry.get("publisher") or ""
    if src:
        b.add(period(src) + " ")
    doi = doi_of(entry)
    if doi:
        b.add(f"https://doi.org/{doi}")
    return b.done()


def format_entry(entry, style_id=""):
    b = Build()
    slot = entry.get("slot") or "other"
    if slot == "journal":
        return _journal(entry, b)
    if slot == "book":
        return _book(entry, b)
    if slot in ("other",):
        return _other(entry, b)
    # 章节、学位论文、标准等：字段不足，交给模板或保持原样
    return None, []
