"""Chicago author-date 文献条目格式化器（杂志/期刊与专著两类）。"""

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
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


def _journal(entry, b):
    year = year_of(entry)
    b.add(f"{_authors(entry.get('authors') or [])} {year}. ")
    b.add(f'"{period(entry.get("title") or "").rstrip(".")}." ')
    journal = entry.get("journal") or ""
    vol = entry.get("volume") or ""
    issue = entry.get("issue") or ""
    pages = entry.get("pages") or ""
    if journal:
        b.add(journal, italic=True)
        if vol:
            b.add(" ")
            b.add(vol, italic=True)
        if issue:
            b.add(f" ({issue})")
        if pages:
            b.add(f": {pages}")
        b.add(".")
    doi = doi_of(entry)
    if doi:
        b.add(f" https://doi.org/{doi}")
    return b.done()


def _book(entry, b):
    year = year_of(entry)
    b.add(f"{_authors(entry.get('authors') or [])} {year}. ")
    add_italic_title(b, entry.get("title") or "")
    pub = entry.get("publisher") or ""
    if pub:
        b.add(" " + period(pub))
    return b.done()


def format_entry(entry, style_id=""):
    b = Build()
    slot = entry.get("slot") or "other"
    if slot == "journal":
        return _journal(entry, b)
    if slot == "book":
        return _book(entry, b)
    return None, []
