"""Harvard 文献条目格式化器（Elsevier / Wiley 两种常见变体）。

规则集 style_id 写 HARVARD-ELSEVIER 或 HARVARD-WILEY；其他 HARVARD*
按 Elsevier 变体处理（两者都是作者-年份制，差异只在标点与年份位置）。
"""

from .common import (Build, add_italic_title, doi_of, initials_text, period,
                     split_name, year_of)


def _author(name, spaced=True):
    sur, initials = split_name(name)
    ini = initials_text(initials, spaced=spaced)
    return f"{sur}, {ini}".strip().rstrip(",") if ini else sur


def _authors(names, variant):
    names = [n for n in names if n]
    if not names:
        return ""
    if variant == "wiley":
        parts = [_author(n) for n in names]
        if len(parts) == 1:
            return parts[0]
        if len(parts) == 2:
            return f"{parts[0]} and {parts[1]}"
        return ", ".join(parts[:-1]) + " and " + parts[-1]
    parts = [_author(n, spaced=False) for n in names]
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]}, {parts[1]}"
    return ", ".join(parts[:-1]) + " and " + parts[-1]


def _journal(entry, b, variant):
    year = year_of(entry)
    authors = _authors(entry.get("authors") or [], variant)
    title = period(entry.get("title") or "")
    journal = entry.get("journal") or ""
    vol = entry.get("volume") or ""
    issue = entry.get("issue") or ""
    pages = entry.get("pages") or ""
    doi = doi_of(entry)

    if variant == "wiley":
        b.add(f"{authors} {year}, ")
        b.add(f"'{title.rstrip('.')}', ")
        if journal:
            b.add(journal, italic=True)
            if vol:
                b.add(", ")
                b.add(vol, italic=True)
            if issue:
                b.add(f"({issue})")
            if pages:
                b.add(f", pp. {pages}")
            b.add(".")
        if doi:
            b.add(f" https://doi.org/{doi}")
        return b.done()

    b.add(f"{authors.rstrip(',')}, {year}. ")
    b.add(title + " ")
    if journal:
        b.add(journal, italic=True)
        if vol:
            b.add(" ")
            b.add(vol, italic=True)
        if issue:
            b.add(f" ({issue})")
        if pages:
            b.add(f", {pages}")
        b.add(".")
    if doi:
        b.add(f" https://doi.org/{doi}")
    return b.done()


def _book(entry, b, variant):
    year = year_of(entry)
    authors = _authors(entry.get("authors") or [], variant)
    if variant == "wiley":
        b.add(f"{authors}, {year}. ")
    else:
        b.add(f"{authors.rstrip(',')}, {year}. ")
    add_italic_title(b, entry.get("title") or "")
    pub = entry.get("publisher") or ""
    if pub:
        b.add(" " + period(pub))
    return b.done()


def format_entry(entry, style_id=""):
    variant = "wiley" if "WILEY" in (style_id or "").upper() else "elsevier"
    b = Build()
    slot = entry.get("slot") or "other"
    if slot == "journal":
        return _journal(entry, b, variant)
    if slot == "book":
        return _book(entry, b, variant)
    return None, []
