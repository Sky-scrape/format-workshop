"""引用体系重排：编号制（[1] / 上标）↔ 作者-年份制。

**这是一个会改写文字的能力，默认关闭。** 用户在界面上勾选后才执行，
理由见 DESIGN：引用映射一旦出错是语义级错误（引错文献），比排版错误严重。

两道闸门，任一不过就整批不落盘：
1. 每条正文引用编号都要能唯一定位到文献表某一条
2. 每条文献都要被正文引用过（规则集 must_require_bidirectional_match）
做不到的位置原样保留，并在审计里写清楚"这里没转、为什么"。

体例模板由规则集给出（citation.reference_list.templates），代码不内置任何
刊物的写法——换刊只改规则集。
"""

import re

CN_CHAR = re.compile(r"[\u4e00-\u9fa5]")
YEAR = re.compile(r"(?:19|20)\d{2}")

# 正文里的编号引用：[1]、[1,2]、[3-5]、[1，2]
CITE = re.compile(r"\[\s*(\d{1,3}(?:\s*[,，\-–—~]\s*\d{1,3})*)\s*\]")

# 数学语境里同样有「[0, 1]」「[90, 100]」这样的数字括号：取值区间、百分比
# 区间。它们与编号引用同形，一旦括号里的数字恰好都是已收录且被引的条目，
# 就会被静默改写成作者-年份——那是语义级错误。拦截分三道，各有适用面：
# 1. 文献表从 [1] 起编（GB/T 7714），括号里出现 0 就不可能是引用——
#    [0]、[0, 1] 只会是取值区间或笔误，任何语境都拦下。
# 2. 多数字括号（[90, 100]、[1-5]）左邻（可隔空白）是数字、百分号、±、
#    比较号这类字符，或右邻紧跟百分号，按取值区间处理。
# 3. ∈、≈、≡ 这类关系符对所有形式生效（「属于集合 [1]」是数学写法）；
#    其余数学邻接符号只拦多数字形式——「10%[12]」「图3[4]」「5‰[13]」里
#    括号前的 %、数字、‰ 是前一个记号的一部分，单数字括号仍是真引用。
#    左邻只排除数学字符、不排除拉丁字母：「see [12]」是英文稿的正常引用，
#    按词排除会把英文稿的转换整体堵死（闸门全量失败、一处都转不了）。
MATH_BEFORE = re.compile(r"[\d.%‰±×÷=<>≥≤∈≈≡∩∪∑√∞^]")
MATH_RELATION = "∈≈≡"
MATH_AFTER = "%‰±"

# 数量/区间语境（fail-closed 的最后一道）：这些样文本与编号引用同形，却
# 几乎总是数量区间——「调研对象为[90, 100]家企业」「[90, 100]为样本量的
# 上下界」。判据从严取「紧邻」：区间词必须紧贴括号（中间只隔空白），远处
# 出现不算——「另有[5, 6]支持这一区间估计」里的「区间」属于前半句的
# 「区间估计」，不归 [5, 6] 管；按 6 字窗口找子串会把这种真引用整批误杀
# （双向体检随之判「文献未被引用」→ 功能整体关闭）。命中者不改写，也不算
# 进「已被引用」，由调用方逐处写进审计让人看得见、判得了。
INTERVAL_WORDS = ("取值", "区间", "范围", "介于", "落入", "位于")
INTERVAL_TAIL = "是为在约"          # 「取值在」「范围是」这类「词+系词」
QUANT_AFTER_CHARS = set("家个种份名篇项元人年月日次点分秒只条处辆台页倍棵株位")
QUANT_AFTER_WORDS = ("为", "左右", "上下", "分别是", "分别为",
                     "之间", "以内", "以上", "以下")


def _math_context(text, m):
    """数学邻接：任何形式都不该当引用。"""
    multi = re.search(r"\d\s*[,，\-–—~]\s*\d", m.group(1)) is not None
    i = m.start() - 1
    while i >= 0 and text[i].isspace():
        i -= 1
    if i >= 0 and (text[i] in MATH_RELATION
                   or (multi and MATH_BEFORE.match(text[i]))):
        return True
    return m.end() < len(text) and text[m.end()] in MATH_AFTER


def _interval_before(text, m):
    """多数字括号左边是否紧贴着「取值/区间/范围…」（含「取值在」这类词+系词）。"""
    before = text[max(0, m.start() - 8):m.start()].rstrip()
    return any(before.endswith(w) or before.endswith(w + t)
               for w in INTERVAL_WORDS for t in INTERVAL_TAIL)


def _quantity_after(text, m):
    """右边是否紧贴量词或「为/之间」这类数量接续词。"""
    after = text[m.end():m.end() + 8].lstrip()
    if after[:1] in QUANT_AFTER_CHARS:
        return True
    return any(after.startswith(w)
               for w in QUANT_AFTER_WORDS + INTERVAL_WORDS)


def citation_spans(text):
    """真引用：CITE 命中剥掉数学语境后的部分，返回 match 列表。

    体检（check）与改写（render_from_ir.rewrite_citations）共用这里，
    两端认到的引用永远一致。
    """
    return classify_spans(text)[0]


def ambiguous_spans(text):
    """数量语境的可疑样文本：不改写、不计入「已被引用」，调用方写进审计。"""
    return classify_spans(text)[1]


def classify_spans(text):
    """把 CITE 命中分成（真引用，数量语境的可疑样文本）两类。

    可疑样文本不改写、不计入「已被引用」，调用方逐处写进审计。
    """
    text = text or ""
    cites, ambig = [], []
    for m in CITE.finditer(text):
        spec = m.group(1)
        if any(n == 0 for n in expand_numbers(spec)):
            continue                  # 编号从 [1] 起编，含 0 是区间或笔误
        if _math_context(text, m):
            continue
        multi = re.search(r"\d\s*[,，\-–—~]\s*\d", spec) is not None
        if multi and _interval_before(text, m):
            ambig.append(m)
            continue
        if _quantity_after(text, m):
            ambig.append(m)
            continue
        cites.append(m)
    return cites, ambig

ITEM = re.compile(
    r"^(?:\[\s*\d+\s*\]\s*)?"
    r"(?P<authors>.+?)[.．]\s*"
    r"(?P<title>.+?)"
    r"\[(?P<kind>[A-Z]+(?:/[A-Z]+)?)\]"
    r"[.．]?\s*(?P<rest>.*)$", re.S)

# 期刊出处：刊名, 年, 卷(期): 起-止
JOURNAL = re.compile(
    r"^(?P<journal>.+?)[,，]\s*(?P<year>(?:19|20)\d{2})\s*"
    r"(?:[,，]\s*(?P<tail>.*))?$", re.S)

VOLUME = re.compile(
    r"^(?P<vol>\d+)?\s*(?:\((?P<issue>[^)]+)\))?\s*(?::\s*(?P<pages>[\d\-–—~]+))?")

# 类型标识 → 模板键
KIND = {
    "J": "journal", "M": "book", "D": "thesis", "C": "chapter",
    "S": "standard", "N": "other", "R": "other", "P": "other",
    "EB/OL": "other", "Z": "other", "A": "other",
}


def is_cn(s):
    return bool(CN_CHAR.search(s or ""))


def split_authors(text):
    """把作者段切成名字列表，丢掉「et al」「等」这类占位。"""
    raw = re.split(r"[,，、;；]", text or "")
    out = []
    for a in raw:
        a = a.strip().strip(".")
        if not a:
            continue
        if a.lower() in ("et al", "etal", "等", "and"):
            continue
        out.append(a)
    return out


def pick_year(text):
    m = YEAR.search(text or "")
    return m.group(0) if m else ""


def parse_volume(tail):
    """解析「30(3): 286-297」或「620: 47-60」。"""
    tail = (tail or "").strip().rstrip(".")
    m = VOLUME.match(tail)
    if not m:
        return "", "", ""
    return (m.group("vol") or "").strip(), (m.group("issue") or "").strip(), \
        (m.group("pages") or "").strip()


def parse_entry(text):
    """把一条 GB/T 7714 式条目拆成字段。拆不出来返回 None。"""
    m = ITEM.match((text or "").strip())
    if not m:
        return None
    authors = split_authors(m.group("authors"))
    if not authors:
        return None
    kind = m.group("kind").upper()
    rest = (m.group("rest") or "").strip()
    entry = {
        "authors": authors,
        "title": m.group("title").strip(),
        "kind": kind,
        "slot": KIND.get(kind, "other"),
        "lang": "cn" if is_cn(m.group("authors") + m.group("title")) else "en",
        "year": "", "journal": "", "volume": "", "issue": "", "pages": "",
        "publisher": "", "source": "",
    }

    if kind == "J":
        j = JOURNAL.match(rest)
        if j:
            entry["journal"] = j.group("journal").strip().rstrip(".")
            entry["year"] = j.group("year")
            vol, issue, pages = parse_volume(j.group("tail") or "")
            entry.update(volume=vol, issue=issue, pages=pages)
    if not entry["year"]:
        entry["year"] = pick_year(rest)

    if kind in ("M", "S", "R"):
        # 「Thousand Oaks: Sage, 2014」「Warrendale: SAE International, 2021」
        pub = re.split(r"[,，]", rest)[-2:] if rest else []
        entry["publisher"] = ""
        m2 = re.search(r"([^:：,，]+)[:：]\s*([^,，]+?)\s*[,，]\s*(?:19|20)\d{2}", rest)
        if m2:
            entry["publisher"] = m2.group(2).strip().rstrip(".")
        elif pub:
            entry["publisher"] = pub[0].strip().rstrip(".")
    if kind in ("EB/OL", "N", "Z"):
        # 先摘掉 URL——它的「https:」会被当成「出版地: 出版者」的冒号，
        # 把整条网址切出来当来源
        clean = re.sub(r"https?://\S+", "", rest)
        head = re.split(r"[,，]\s*(?:\(|\[|19|20)", clean)[0].strip()
        if head.startswith(("(", "[")):
            head = ""
        else:
            head = re.split(r"[:：]", head)[-1]
        entry["source"] = head.strip().rstrip(".")
        if entry["source"] == entry["authors"][0]:
            entry["source"] = ""      # 作者就是发布机构时不再重复一遍
    if kind == "D":
        entry["publisher"] = (re.split(r"[,，]", rest)[0] or "").strip().rstrip(".")
    return entry


def cn_authors(names):
    return "、".join(names)


def en_author(name):
    """「Wang H」→「Wang，H.」；已经是「Weiss，D.」的原样返回。"""
    name = name.strip()
    m = re.match(r"^(?P<sur>[A-Z][A-Za-z'\-]+)\s+(?P<ini>[A-Z](?:\s*[A-Z])*)$", name)
    if m:
        initials = " ".join(x + "." for x in m.group("ini").split())
        return f"{m.group('sur')}，{initials}"
    m = re.match(r"^(?P<sur>[A-Z][A-Za-z'\-]+)[,，]\s*(?P<ini>.*)$", name)
    if m:
        ini = " ".join(x + "." for x in re.split(r"[\s.]+", m.group("ini")) if x)
        return f"{m.group('sur')}，{ini}"
    return name


def en_authors(names):
    """两作者用 and 连接，三作者以上前若干用全角逗号、最后一个用 and。"""
    parts = [en_author(n) for n in names]
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} and {parts[1]}"
    return "，".join(parts[:-1]) + " and " + parts[-1]


def in_text_style(rules=None):
    """作者-年份正文引用的版式：((左括号, 右括号), 多条引用的分隔符)。

    只消费作者-年份制刊物的声明——citation.in_text.bracket / multiple_separator
    （英文 Harvard 体例声明半角「()」与「; 」，mwm 等声明全角「；」，此前
    两者都无消费者，英文刊被硬编码成全角中文括号，PB1）。编号制刊物的同名
    声明（cjyj 等的 bracket「[]」）描述的是编号标注的写法，作者-年份转换
    用它会产出刊物从未声明过的形态，一律不取。缺省回退全角括号与全角
    分号——中文刊行为与硬编码时期逐字节一致。
    """
    cite = (rules or {}).get("citation") or {}
    cfg = cite.get("in_text") or {}
    if cite.get("system") != "author-year":
        cfg = {}
    b = cfg.get("bracket") or "（）"
    return (b[0], b[-1]), (cfg.get("multiple_separator") or "；")


def label_body(entry, rules=None):
    """只要「作者，年份」这半截，供多引用合并成一个括号用。"""
    (open_b, close_b), _sep = in_text_style(rules)
    full = in_text_label(entry, rules)
    if full.startswith(open_b) and full.endswith(close_b):
        return full[1:-1]
    return full


def in_text_label(entry, rules=None):
    """正文里的作者-年份写法：（周志华，2026）/（Kang and Kim，2008）。"""
    cfg = ((rules or {}).get("citation") or {}).get("in_text") or {}
    names = entry["authors"]
    if entry["lang"] == "cn":
        if len(names) == 1:
            who = names[0]
        elif len(names) == 2:
            who = "、".join(names)
        else:
            who = names[0] + "等"
    else:
        if len(names) == 1:
            who = en_author(names[0]).split("，")[0]
        elif len(names) == 2:
            who = f"{names[0].split()[0]} and {names[1].split()[0]}"
        else:
            who = names[0].split()[0] + (" et al." if cfg.get("en_three_plus") else "等")
    (open_b, close_b), _sep = in_text_style(rules)
    return f"{open_b}{who}，{entry['year']}{close_b}"


def volume_part(entry):
    if not entry["volume"]:
        return ""
    if entry["issue"]:
        return f"，{entry['volume']}({entry['issue']})"
    return f"，vol.{entry['volume']}"


def page_part(entry):
    if not entry["pages"]:
        return ""
    pages = entry["pages"].replace("-", "~").replace("–", "~").replace("—", "~")
    return f"，pp.{pages}"


def _entry_fields(entry):
    """模板可用的占位符值。format_entry 与带斜体区间的渲染共用这一份。"""
    if entry["lang"] == "cn":
        who = cn_authors(entry["authors"])
    else:
        who = en_authors(entry["authors"])
    return {
        "authors": who,
        "title": entry["title"],
        "year": entry["year"],
        "journal": entry["journal"],
        "publisher": entry["publisher"],
        "source": entry["source"],
        "issue_part": f"第{entry['issue']}期" if entry["issue"] else "",
        "volume_part": volume_part(entry),
        "page_part": page_part(entry),
        "issue": entry["issue"],
    }


def _tpl_for(entry, rules):
    rl = ((rules or {}).get("citation") or {}).get("reference_list") or {}
    tpls = rl.get("templates") or {}
    key = f"{entry['lang']}_{entry['slot']}"
    tpl = (tpls.get(key) or tpls.get(f"{entry['lang']}_book")
           or tpls.get(f"{entry['lang']}_other"))
    if not tpl or "{" not in tpl:
        return None, key if key in tpls else None, rl
    return tpl, key, rl


def italic_fields(rules, key):
    """该模板声明要排斜体的占位符名。

    规则集写法（journal-schema §4.3）二选一：
      "italic_fields": ["journal"]                          # 所有模板共用
      "italic_fields": {"en_journal": ["journal"],
                        "en_book": ["title"]}               # 按模板键区分
    中文刊的《刊名》本来就带书名号，多数不需要斜体；不声明就一律不加，
    绝不替期刊猜——斜体也是版式事实，得有出处。
    """
    conf = (((rules or {}).get("citation") or {}).get("reference_list")
            or {}).get("italic_fields")
    if isinstance(conf, list):
        return {str(x) for x in conf}
    if isinstance(conf, dict):
        vals = conf.get(key) or []
        if isinstance(vals, str):
            vals = [vals]
        return {str(x) for x in vals if isinstance(x, str)}
    return set()


def format_entry_with_spans(entry, rules):
    """按模板渲染一条文献，并给出声明为斜体的文字区间。

    返回 (text, [(start, end)])；没有可用模板时返回 (None, [])。区间是在
    最终文本上的偏移，供渲染器只给刊名/书名加斜体，条目其余字符与 run
    格式不动。未声明 italic_fields 时区间为空，行为与旧版逐字节一致。
    """
    tpl, key, _rl = _tpl_for(entry, rules)
    if tpl is None:
        return None, []
    fields = _entry_fields(entry)
    want = italic_fields(rules, key or "")
    # 用控制字符标出斜体字段的位置，渲染完再换算成偏移并剥掉标记。
    # 标记字符不会出现在规则集模板或稿件文字里；即便出现，format 只读
    # fields 的值，模板本身不参与转义。
    marked = dict(fields)
    for name in want:
        if name in marked and marked[name]:
            marked[name] = "\x01" + marked[name] + "\x02"
    try:
        rendered = tpl.format(**marked)
    except KeyError:
        return None, []
    out, spans, pos, depth = [], [], 0, []
    for ch in rendered:
        if ch == "\x01":
            depth.append(len(out))
        elif ch == "\x02":
            if depth:
                spans.append((depth.pop(), len(out)))
        else:
            out.append(ch)
    text = "".join(out)
    # 模板里空字段会留下「，，」「。。」——旧版用字符串 replace 收掉。这里
    # 同时维护旧位置→新位置的映射，斜体区间不会因为前面少了一个逗号而错位。
    collapsed, mapping = [], []
    i = 0
    while i < len(text):
        mapping.append(len(collapsed))
        if text[i] in "，。" and i + 1 < len(text) and text[i + 1] == text[i]:
            i += 1                    # 跳过重复的一个标点
            continue
        collapsed.append(text[i])
        i += 1
    mapping.append(len(collapsed))
    text = "".join(collapsed)
    spans = [(mapping[s], mapping[e]) for s, e in spans]
    # 合并相邻/重叠区间，防止两个声明字段背靠背时区间交叉
    spans.sort()
    merged = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return text, merged


def format_entry(entry, rules):
    """按规则集的 templates 渲染一条文献。没有可用模板时返回 None。

    模板必须带 {字段} 占位符才是可渲染的体例模板。不带占位符的键值（如
    GB/T 7714 的著录写法说明「[序号]主要责任者.文献题名[M]…」）是给人看的
    格式描述——拿去 format 只会原样返回，把真条目替换成一句格式说明，
    属于语义级破坏，一律视为不可渲染。
    """
    text, _spans = format_entry_with_spans(entry, rules)
    return text


def expand_numbers(spec):
    """「1,3-5」→ [1,3,4,5]；保持出现顺序。"""
    out = []
    for piece in re.split(r"[,，]", spec):
        piece = piece.strip()
        if not piece:
            continue
        m = re.match(r"^(\d{1,3})\s*[\-–—~]\s*(\d{1,3})$", piece)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(min(a, b), max(a, b) + 1))
        elif piece.isdigit():
            out.append(int(piece))
    return out


def build_index(entries):
    """条目序号 → 解析结果。序号取自条目开头的 [n]。"""
    index = {}
    for raw in entries:
        m = NUMBERED_PREFIX.match(raw)
        if not m:
            continue
        parsed = parse_entry(raw)
        if parsed:
            parsed["raw"] = raw
            index[int(m.group(1))] = parsed
    return index


NUMBERED_PREFIX = re.compile(r"^\s*\[\s*(\d{1,3})\s*\]")


def unnumbered_entries(texts):
    """文献表里「像条目但没有 [n] 编号」的段落原文（F23）。

    build_index 只收带编号的条目：非编号条目进不了双向校验，条目重排时
    也会被留在原地，与搬移后的条目拼成错序的混合表。返回非空时调用方必须
    整批放弃转换，并把这些段落点名进审计。
    """
    return [t for t in (x.strip() for x in texts)
            if t and not NUMBERED_PREFIX.match(t)]


def footnote_citation_hits(texts):
    """脚注文本里的编号引用（F23），与正文同一套 classify_spans 口径。

    脚注部件不在正文体检与改写范围内：条目一搬家，脚注里的编号就指向
    错位、已被改写的条目。返回非空时调用方整批放弃并点名脚注。数学语境
    的判定与正文一致——「区间为[90, 100]」这类样文本不算引用，不误杀。
    """
    hits = []
    for text in texts:
        hits.extend(classify_spans(text or "")[0])
    return hits


def gb_sort_key(text):
    """中文按 GB2312 序取排序键：一级汉字库本身就是拼音序，用它近似拼音。

    不引入第三方拼音依赖——requirements 里没有 pypinyin，靠 import 探测
    的「真拼音」永远走不到，码位序会冒充拼音序把「周一」排到「张三」前面
    （F22）。二级汉字在 GB2312 里按部首排，对文献条目的常用字够用；编码
    取不到的字统一排到最后。调用方的审计 note 必须如实写「GB 序（近似
    拼音）」，不得写成真拼音序。
    """
    buf = bytearray()
    for ch in text or "":
        try:
            buf += ch.encode("gb2312")
        except UnicodeEncodeError:
            buf += b"\xff\xff"
    return bytes(buf)


def check(paragraph_texts, index, rules):
    """转换前体检：正文引用编号与文献表必须双向对得上。

    返回 (mapping, problems)。problems 非空时调用方必须整批放弃——
    缺一条映射就会把引用指向错误的文献，那是语义错误，不能凑合。
    """
    problems = []
    cited = set()
    for text in paragraph_texts:
        for m in citation_spans(text or ""):
            for num in expand_numbers(m.group(1)):
                cited.add(num)
                if num not in index:
                    problems.append(f"正文引用了 [{num}]，但文献表里没有这一条")

    if ((rules or {}).get("citation") or {}).get("reference_list", {}) \
            .get("must_require_bidirectional_match"):
        for num in index:
            if num not in cited:
                problems.append(f"文献 [{num}] 在正文里没有被引用")

    for num, e in index.items():
        if not e["year"]:
            problems.append(f"文献 [{num}] 解析不出年份，无法生成作者-年份写法")
        if e["kind"] == "J" and not e["journal"]:
            problems.append(f"文献 [{num}] 解析不出刊名，条目重排会缺字段")
    return cited, sorted(set(problems))
