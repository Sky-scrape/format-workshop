"""两个渲染器共用的 docx 版式写入原语。

样板渲染器（render_sample.py）与交付稿渲染器（render_from_ir.py）各写过一份
set_run／缩进写入。复制品迟早漂移：同一规则集下样板写 11.5pt、交付稿写 12pt，
用户拿样板核对时会以为自己排错了。所以纯函数与缺省值只留这一份，两个渲染器
都从这里 import。

这里只放不依赖 IR、规则集语义的「往 XML 里写字」的动作；缺省值 DEFAULTS 也
放这里，因为样板与交付稿必须用同一张缺省表（F30：渲染器的缺省值只有一源）。
"""

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# ---- F30：渲染器的缺省值只有这一张表 --------------------------------------
# 键 → (缺省值, 缺省说明)。规则集没写时从这里取，代码里不再散落第二真源；
# 用到缺省值的审计条目 rule 走 rules_ref 的 '~' 可缺省约定、note 带「缺省值」
# 说明——用户必须分得清哪些值是期刊明文规定、哪些是代码约定（DESIGN §6）。
# 「缺省值」与「推定值」是两回事：推定值专指规则集标了 measured:false、登记
# 进 assumptions 的取值（见 assumed()）；刊里整条没写时走的是代码缺省，不得
# 冒称推定——否则审计会谎报该刊登记过这条推定。
#
# 样板渲染器与交付稿渲染器共用这张表：以前样板里有 20、1.2、1.35 这类裸数字，
# 改了一边另一边不动，样板就不再是交付稿的版式基准。数值型行距倍数的键名统一
# 以 _line_mult 结尾，单位是「字号×倍数」。
DEFAULTS = {
    "body_line_pt":          (20,     "正文行距未声明固定磅值，兜底取"),
    "body_line_mult":        (1.2,    "正文行距未声明数值，按字号×1.2 折算"),
    "title_size_pt":         (22,     "标题字号未声明，取"),
    "title_font_cn":         ("黑体", "标题中文字体未声明，取"),
    "title_line_mult":       (1.5,    "标题行距按字号×1.5 折算"),
    "subtitle_font_cn":      ("黑体", "副标题中文字体未声明，取"),
    "subtitle_size_pt":      (16,     "副标题字号未声明，取"),
    "subtitle_line_mult":    (1.5,    "副标题行距按字号×1.5 折算"),
    "authors_line_mult":     (1.5,    "作者行行距按自身字号×1.5 折算"),
    "affiliations_line_mult": (1.5,   "作者单位行距按自身字号×1.5 折算"),
    "classification_line_mult": (1.45, "号段等小字号行按字号×1.45 折算行距"),
    "table_font_size_pt":    (9,      "表内文字字号未声明，取"),
    "caption_font_cn":       ("黑体", "题注中文字体未声明，取"),
    "caption_gap":           ("一字空", "题注编号与标题的间隔未声明，取"),
    "note_size_pt":          (9,      "表注字号未声明，取"),
    "author_note_line_mult": (1.6,    "作者简介行距按字号×1.6 折算"),
    "author_note_label_font_cn": ("黑体", "作者简介标签中文字体未声明，取"),
}


def set_run(run, cn=None, latin=None, size=None, bold=None, letter_spacing=None):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    if cn:
        rf.set(qn("w:eastAsia"), cn)
    if latin:
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)
    if size:
        # Word 的字号以半磅为刻度（w:sz 是 half-point 整数）。直接传 11.8 会被截断成 11.5，
        # 规范页写 11.8、交付稿却是 11.5，三方对不上。这里显式取最近的半磅。
        run.font.size = Pt(round(size * 2) / 2)
    if bold is not None:
        run.font.bold = bold
    if letter_spacing is not None:
        # 字间距按 twips 写入（1pt = 20twips）；《经济研究》标题每字加宽约 1.3pt。
        # 只有 None 才是「规则集未声明」——此时不动 run 上已有的字间距；数值
        # （含 0）都生效：声明 0 时移除已有的 w:spacing，否则从带字间距模板
        # 切到 0pt 的刊，成品残留旧字间距（SETRUN0）
        old = rpr.find(qn("w:spacing"))
        if old is not None:
            rpr.remove(old)
        if letter_spacing:
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:val"), str(int(round(letter_spacing * 20))))
            rpr.append(sp)


def set_chars_indent(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(int(chars * 210)))


def set_hanging_indent(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:leftChars"), str(int(chars * 100)))
    ind.set(qn("w:left"), str(int(chars * 210)))
    ind.set(qn("w:hangingChars"), str(int(chars * 100)))
    ind.set(qn("w:hanging"), str(int(chars * 210)))


# render_from_ir 的公开名历来叫 set_hanging（tests/test_fix_audit_and_indent.py
# 与 preflight 从这里取）。两个渲染器共用一份实现后，旧名作为别名留着，
# 免得改一次名把调用点全掀一遍。
set_hanging = set_hanging_indent


def set_line_unit_space(p, before_lines=0, after_lines=0):
    ppr = p._p.get_or_add_pPr()
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        ppr.append(sp)
    sp.set(qn("w:beforeLines"), str(int(before_lines * 100)))
    sp.set(qn("w:afterLines"), str(int(after_lines * 100)))
