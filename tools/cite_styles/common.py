"""样式插件共用的作者名、年份、DOI 与斜体区间工具。"""

import re

CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")


def is_cn(text):
    return bool(CJK.search(text or ""))


def split_name(name):
    """「Wang H」/「Weiss, D.」→ (姓, 名缩写列表)。中文名原样作姓。"""
    name = (name or "").strip().strip(".")
    if is_cn(name):
        return name, []
    if "," in name:
        sur, rest = name.split(",", 1)
        initials = [x for x in re.split(r"[\s.]+", rest) if x]
        return sur.strip(), initials
    parts = name.split()
    if len(parts) == 1:
        return parts[0], []
    return parts[0], parts[1:]


def initials_text(initials, spaced=True):
    if not initials:
        return ""
    sep = " " if spaced else ""
    return sep.join(i[0].upper() + "." for i in initials if i)


def surname(name):
    return split_name(name)[0]


def year_of(entry):
    return entry.get("year") or "n.d."


def doi_of(entry):
    """从原始条目里取 DOI；没有就返回空串。不编造。"""
    raw = entry.get("raw") or ""
    m = re.search(r"(?:doi\s*[:：]?\s*|doi\.org/)(10\.\d{4,9}/[^\s,;，；]+)",
                  raw, re.IGNORECASE)
    if not m:
        return ""
    return m.group(1).rstrip(".")


def period(text):
    text = (text or "").strip()
    return text if not text or text.endswith((".", "?", "!")) else text + "."


def add_italic_title(b, text):
    """书名/刊名斜体只包文字，句末标点不算斜体。"""
    t = period(text)
    if t:
        b.add(t[:-1], italic=True)
        b.add(t[-1])


class Build:
    """拼一条文献文本，同时记下需要斜体的区间。"""

    def __init__(self):
        self.parts = []
        self.length = 0
        self.spans = []

    def add(self, text, italic=False):
        text = text or ""
        start = self.length
        self.parts.append(text)
        self.length += len(text)
        if italic and text:
            self.spans.append((start, self.length))
        return self

    def done(self):
        merged = []
        for s, e in self.spans:
            if s >= e:
                continue
            if merged and s <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))
            else:
                merged.append((s, e))
        return "".join(self.parts), merged
