"""中文本科论文（设计）排版助手库。

严格实现《本科生论文（设计）撰写要求》的页面、字号、字体、行距、
页码与标题级次规定。只依赖 python-docx。

用法::

    from thesisdocx import *
    doc = new_thesis_document()
    setup_page(doc)
    add_heading(doc, 1, "引言")
    add_body(doc, "正文内容……")
    setup_front_matter_page_numbers(doc)
    doc.save("thesis.docx")
"""

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import re

# --------------------------------------------------------------------------
# 常量：字号、字体
# --------------------------------------------------------------------------

#: 中文习惯字号名 -> 磅值
CN_PT = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}

_SONG, _HEI = "SimSun", "SimHei"
TNR = "Times New Roman"

# 规范中直接引用的字号（磅值）
PT_TITLE = CN_PT["小三"]        # 论文题目 小3号黑体
PT_SUBTITLE = CN_PT["四号"]     # 副标题 4号黑体
PT_BODY = CN_PT["小四"]         # 正文/标题 小4号
PT_TOC_HEAD = CN_PT["三号"]     # 目录题头 3号加粗宋体
PT_ABSTRACT_HEAD = CN_PT["小二"]  # 摘要题头 小2号加粗宋体
PT_KEYWORDS_LABEL = CN_PT["四号"]  # “关键词” 4号黑体
PT_HEADER = CN_PT["五号"]       # 页眉 5号宋体
PT_CAPTION = CN_PT["五号"]      # 表序表名/图号图名 5号宋体
PT_SMALL_TABLE = CN_PT["小五"]  # 表内文字 小5号宋体

LINE_SPACING = 1.5              # 全文统一 1.5 倍行距

# 页面设置（Word 默认值，规范指定）
PAGE_W, PAGE_H = Cm(21.0), Cm(29.7)
MARGIN_TB, MARGIN_LR = Cm(2.54), Cm(3.17)
HEADER_DIST, FOOTER_DIST = Cm(1.5), Cm(1.75)

# 三线表线宽（八分之一磅为单位的 OOXML sz 值）
BORDER_THICK = 12   # 1.5 磅
BORDER_THIN = 4     # 0.5 磅


# --------------------------------------------------------------------------
# 底层工具
# --------------------------------------------------------------------------

def set_run_font(run, size_pt=None, ascii_font=TNR, east_asia_font=_SONG, bold=None):
    """同时设置西文字体与中文字体（必须写 w:eastAsia，否则中文不变）。"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east_asia_font)
    if east_asia_font is not None:
        rFonts.set(qn("w:cs"), ascii_font)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    return run


def set_paragraph_spacing(paragraph, line_spacing=LINE_SPACING,
                          before_pt=0, after_pt=0, first_line_indent_chars=None):
    """设置行距与段前段后。first_line_indent_chars 用“字符”为单位（w:firstLineChars）。"""
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line_spacing
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if first_line_indent_chars is not None:
        pPr = paragraph._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(first_line_indent_chars * 100)))
    return paragraph


def _no_first_line_indent(paragraph):
    """题头、题注、参考文献条目等不应有首行缩进（覆盖样式继承）。"""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")
    return paragraph


def blank_line(document):
    """插入一个真正的空行（用于标题“上下均空一行”）。"""
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING)
    r = p.add_run("")
    set_run_font(r, PT_BODY)
    return p


#: 第 3 级的规范编号是“1”（原文不加点号）；“1.”“1．”“1、”只作兼容识别
_HEADING_STRIP = {
    1: re.compile(r"^[一二三四五六七八九十百]+[、.．]\s*"),
    2: re.compile(r"^[（(][一二三四五六七八九十百]+[）)]\s*"),
    3: re.compile(r"^(?:\d+[.．、]\s*|\d{1,2}(?![\d.．、%])\s*)"),
    4: re.compile(r"^[（(]\d+[）)]\s*"),
}

#: 带分隔符的 3 级编号（“1. 标题”“1、标题”），用于规范化为“1标题”
_HEADING3_SEPARATED = re.compile(r"^(\d+)\s*[.．、](?!\d)\s*(\S.*)$")


def _strip_heading_prefix(text, level):
    """剥除调用者已写入的编号前缀，避免“引言、”“（研究背景）”这类重复。"""
    text = (text or "").strip()
    for lvl in (level, 4, 3, 2, 1):
        pat = _HEADING_STRIP.get(lvl)
        if pat:
            stripped = pat.sub("", text, count=1)
            if stripped != text and stripped.strip():
                return stripped.strip()
    return text


_CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
           "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八",
           "十九", "二十"]

#: 每个标题级次的下一序号，供 add_heading 自动编号
_HEADING_COUNTER = {}


def reset_heading_numbers():
    """重置自动编号计数（新建文档或重新开始时调用）。"""
    _HEADING_COUNTER.clear()


def _already_prefixed(text, level):
    """文本是否已是本级的完整编号，如一级的“一、引言”。"""
    pat = _HEADING_STRIP[level]
    stripped = pat.sub("", text, count=1)
    return stripped != text and bool(stripped.strip())


def _number_heading(text, level, counter=None):
    """补全标题编号。

    - 已带完整编号（“一、引言”“1数据来源”）→ 原样返回，不产生“引言、”。
    - 只带编号骨架（“、引言”“. 数据”“（）背景”）→ 补上自动序号。
    - 纯标题文字（“引言”）→ 自动编号，按同级别调用先后递增。
    - 第 3 级统一写成规范要求的“1标题”；写成“1. 标题”时会去掉分隔符。
    """
    counter = _HEADING_COUNTER if counter is None else counter
    if level not in (1, 2, 3, 4):
        raise ValueError("level 必须为 1—4")
    text = (text or "").strip()
    if level == 3:
        m = _HEADING3_SEPARATED.match(text)
        if m:
            text = f"{m.group(1)}{m.group(2).strip()}"
    if _already_prefixed(text, level):
        return text

    pat = _HEADING_STRIP[level]
    leftover = pat.sub("", text, count=1).strip()
    if not leftover:
        leftover, text = text, ""
    elif leftover != text:
        text = leftover  # 骨架形式，保留剩余标题文字

    counter[level] = counter.get(level, 0) + 1
    seq = counter[level]
    if level == 1:
        cn = _CN_NUM[seq - 1] if seq <= len(_CN_NUM) else str(seq)
        return f"{cn}、{text}"
    if level == 2:
        cn = _CN_NUM[seq - 1] if seq <= len(_CN_NUM) else str(seq)
        return f"（{cn}）{text}"
    if level == 3:
        # 规范编号“1”后不加点号；标题自身以数字开头时补一个空格以免粘连
        return f"{seq}{' ' if text[:1].isdigit() else ''}{text}"
    return f"（{seq}）{text}"


# --------------------------------------------------------------------------
# 页面与页码
# --------------------------------------------------------------------------

def setup_page(section):
    """A4 + Word 默认页面设置 + 1.5 倍行距。"""
    section.page_width, section.page_height = PAGE_W, PAGE_H
    section.top_margin = section.bottom_margin = MARGIN_TB
    section.left_margin = section.right_margin = MARGIN_LR
    section.header_distance = HEADER_DIST
    section.footer_distance = FOOTER_DIST
    # 版心宽度，供表格列宽使用
    return section


def set_page_numbering(section, fmt="decimal", start=None):
    """设置本节页码格式。

    fmt: "decimal"（阿拉伯数字）| "upperRoman"（大写罗马数字）
         也接受 "chineseCounting" 等 OOXML 合法值。
    start: 起始页码；None 表示连续编号。
    """
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        # pgNumType 必须排在 cols 之前才符合 schema 顺序
        cols = sectPr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pgNumType)
        else:
            sectPr.append(pgNumType)
    pgNumType.set(qn("w:fmt"), fmt)
    if start is not None:
        pgNumType.set(qn("w:start"), str(start))
    else:
        if pgNumType.get(qn("w:start")) is not None:
            del pgNumType.attrib[qn("w:start")]
    return section


def add_centered_page_number(section, size_pt=PT_HEADER, ascii_font=TNR,
                             east_asia_font=_SONG, different_first_page=False):
    """页脚居中插入 PAGE 域（默认五号，与页眉一致）。"""
    section.different_first_page_header_footer = different_first_page
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, 1.0)
    _insert_field(p, "PAGE", size_pt, ascii_font, east_asia_font)
    return p


def _insert_field(paragraph, instr, size_pt, ascii_font, east_asia_font):
    """插入简单域，如 PAGE / NUMPAGES。"""
    run = paragraph.add_run()
    set_run_font(run, size_pt, ascii_font, east_asia_font)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr_el, fld_end):
        run._element.append(el)
    return run


def setup_front_matter_page_numbers(doc):
    """把第一节设为大写罗马数字（声明/开题报告表/目录/摘要）。

    摘要页本身也要有页码，故不启用“首页不同”。
    """
    for section in doc.sections:
        set_page_numbering(section, "upperRoman", start=1)
        add_centered_page_number(section, different_first_page=False)
    return doc


def start_body_section(doc):
    """新起一节（正文起），页码改为阿拉伯数字并从 1 开始。"""
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_page(section)
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    set_page_numbering(section, "decimal", start=1)
    add_centered_page_number(section)
    return section


# --------------------------------------------------------------------------
# 页眉
# --------------------------------------------------------------------------

def set_header_text(section, lines, size_pt=PT_HEADER, ascii_font=TNR,
                    east_asia_font=_SONG, link_to_previous=False):
    """写入居中页眉。lines 为字符串或字符串列表（长题名分两行）。"""
    section.header.is_linked_to_previous = link_to_previous
    header = section.header
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, 1.0)
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        if i:
            p.add_run().add_break(WD_BREAK.LINE)
        run = p.add_run(line)
        set_run_font(run, size_pt, ascii_font, east_asia_font)
    return p


def add_page_break(document):
    p = document.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


def start_front_matter_section(doc):
    """新起“前置部分”节（原创性声明/开题报告表/目录/摘要）。

    这些页同样要有页眉和页码（只有封面、扉页例外），页码为大写罗马数字。
    封面/扉页应放在本函数之前的独立节中，且不设页眉页脚。
    """
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_page(section)
    set_page_numbering(section, "upperRoman", start=1)
    add_centered_page_number(section, different_first_page=False)
    return section


def add_section_break(document):
    return document.add_section(WD_SECTION.NEW_PAGE)


# --------------------------------------------------------------------------
# 段落内容
# --------------------------------------------------------------------------

def add_title(document, title, subtitle=None):
    """论文题目（小3号黑体）与副标题（4号黑体），居中。"""
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(title), PT_TITLE, TNR, _HEI, bold=False)
    if subtitle:
        q = document.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(q, LINE_SPACING)
        set_run_font(q.add_run(subtitle), PT_SUBTITLE, TNR, _HEI, bold=False)
    return p


def add_front_matter_head(document, text, size_pt=PT_ABSTRACT_HEAD, *, full_width_space=True):
    """题头（“摘　要”“目　录”），居中加粗宋体，其后隔一行。

    text 为不含空格的词，如 "摘要" -> "摘　要"。
    """
    if full_width_space and len(text) == 2 and "\u3000" not in text:
        text = text[0] + "\u3000" + text[1]
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(text), size_pt, TNR, _SONG, bold=True)
    blank_line(document)  # “隔一行”
    return p


def add_heading(document, level, text, *, blank_around=None):
    """按规范添加 1—4 级标题：小4号加粗宋体，编号依次为“一、”“（一）”“1”“（1）”。

    编号自动补全：
      add_heading(doc, 1, "引言")     -> 一、引言
      add_heading(doc, 2, "研究背景") -> （一）研究背景
      add_heading(doc, 3, "数据来源") -> 1数据来源
      add_heading(doc, 4, "样本筛选") -> （1）样本筛选
    已带完整编号的文本（“一、引言”“1数据来源”）原样保留，不会写成“引言、”。
    第 3 级按原文不加点号；传入“1. 数据来源”会被规范化为“1数据来源”。
    同一级次按调用顺序递增；跨文档复用请先调用 reset_heading_numbers()。
    """
    numbered = _number_heading(text, level)

    if blank_around is None:
        blank_around = level in (1, 2)
    if blank_around:
        blank_line(document)

    style_name = f"ThesisH{level}"
    if style_name not in [s.name for s in document.styles]:
        _ensure_heading_styles(document)
    if style_name in [s.name for s in document.styles]:
        p = document.add_paragraph(style=style_name)
    else:
        p = document.add_paragraph()  # 外来文档缺少自定义样式时降级
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    set_run_font(p.add_run(numbered), PT_BODY, TNR, _SONG, bold=True)

    if blank_around:
        blank_line(document)
    return p


def add_body(document, text, *, indent=True, bold=False, align=None, size_pt=PT_BODY):
    """正文段落：中文小4号宋体，西文 Times New Roman 小4号，1.5 倍行距，首行缩进 2 字符。"""
    p = document.add_paragraph()
    if align is not None:
        p.alignment = align
    set_paragraph_spacing(p, LINE_SPACING, first_line_indent_chars=2 if indent else None)
    set_run_font(p.add_run(text), size_pt, TNR, _SONG, bold=bold)
    return p


def add_body_mixed(document, parts, *, indent=True, align=None):
    """混排段落：parts 为 (文本, 是否西文) 列表，用于中英混排的字体切换。"""
    p = document.add_paragraph()
    if align is not None:
        p.alignment = align
    set_paragraph_spacing(p, LINE_SPACING, first_line_indent_chars=2 if indent else None)
    for text, is_latin in parts:
        set_run_font(p.add_run(text), PT_BODY, TNR, _SONG)
    return p


def add_keywords(document, keywords, label="关键词"):
    """“关键词”4号黑体 + 内容小4号宋体，之间空一个全角空格（不缩进，1.5 倍行距）。"""
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(label + "\u3000"), PT_KEYWORDS_LABEL, TNR, _HEI, bold=False)
    body = keywords if isinstance(keywords, str) else "；".join(keywords)
    set_run_font(p.add_run(body), PT_BODY, TNR, _SONG)
    return p


def add_english_abstract_head(document, text="Abstract"):
    """外文摘要题头：Times New Roman，加粗，小4号。"""
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(text), PT_BODY, TNR, TNR, bold=True)
    blank_line(document)
    return p


def add_english_body(document, text, *, indent=True, bold_lead=None):
    """外文摘要正文：全部 Times New Roman 小4号。"""
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING, first_line_indent_chars=2 if indent else None)
    if not indent:
        _no_first_line_indent(p)
    if bold_lead:
        set_run_font(p.add_run(bold_lead), PT_BODY, TNR, TNR, bold=True)
    set_run_font(p.add_run(text), PT_BODY, TNR, TNR)
    return p


def add_formula(document, formula, number=None):
    """公式：居中，末尾不加标点；编号放行末右对齐。"""
    if number:
        p = document.add_paragraph()
        set_paragraph_spacing(p, LINE_SPACING)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(7.0), 1)   # 居中制表位
        p.paragraph_format.tab_stops.add_tab_stop(Cm(14.0), 2)  # 右对齐制表位
        set_run_font(p.add_run("\t" + formula + "\t" + f"（{number}）"), PT_BODY, TNR, _SONG)
        return p
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    set_run_font(p.add_run(formula.rstrip("。，,.;；")), PT_BODY, TNR, _SONG)
    return p


# --------------------------------------------------------------------------
# 表格（三线表）
# --------------------------------------------------------------------------

def _set_cell_border(cell, edge, sz, val="single", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    el = borders.find(qn(f"w:{edge}"))
    if el is None:
        el = OxmlElement(f"w:{edge}")
        borders.append(el)
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)


def apply_three_line_table(table, header_rows=1):
    """把表格改为三线表：上下 1.5 磅，表头下 0.5 磅，无左右边线。"""
    rows = table.rows
    for ri, row in enumerate(rows):
        # 每行 row 的单元格（含合并）
        for cell in row.cells:
            _set_cell_border(cell, "left", 0, "none")
            _set_cell_border(cell, "right", 0, "none")
            _set_cell_border(cell, "top", 0, "none")
            _set_cell_border(cell, "bottom", 0, "none")
            if ri == 0:
                _set_cell_border(cell, "top", BORDER_THICK)
            if ri == header_rows - 1:
                _set_cell_border(cell, "bottom", BORDER_THIN)
            if ri == len(rows) - 1:
                _set_cell_border(cell, "bottom", BORDER_THICK)
    return table


def add_table_caption(document, index, name):
    """表序表名置于表上、居中、5号宋体；表序与表名之间空一格，名后不加标点。"""
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(f"表{index} {name.rstrip('。，,. ')}"), PT_CAPTION, TNR, _SONG)
    return p


def add_table_note(document, text):
    """表的说明文字：小5号宋体。"""
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(text), PT_SMALL_TABLE, TNR, _SONG)
    return p


def style_table_text(table):
    """表内文字统一小5号宋体。"""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                set_paragraph_spacing(p, 1.0)
                for r in p.runs:
                    set_run_font(r, PT_SMALL_TABLE, TNR, _SONG)
    return table


def empty_cell_dash(cell):
    """数字空缺的格内填“——”（两个字符宽）。"""
    cell.text = ""
    p = cell.paragraphs[0]
    set_paragraph_spacing(p, 1.0)
    set_run_font(p.add_run("——"), PT_SMALL_TABLE, TNR, _SONG)
    return cell


# --------------------------------------------------------------------------
# 图
# --------------------------------------------------------------------------

def add_figure_caption(document, index, name, cite=None):
    """图号图名置于图下、居中、5号宋体；图名与图号空一格；引用号放图名右上角。"""
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(f"图{index} {name.rstrip('。，,. ')}"), PT_CAPTION, TNR, _SONG)
    if cite:
        run = set_run_font(p.add_run(f"[{cite}]"), PT_CAPTION, TNR, _SONG)
        run.font.superscript = True
    return p


# --------------------------------------------------------------------------
# 参考文献与注释
# --------------------------------------------------------------------------

REF_TYPES = ("M", "J", "C", "R", "D", "P", "S", "N", "EB/OL")


def add_reference(document, index, entry):
    """参考文献条目：方括号序号 + 小4号宋体，悬挂缩进。"""
    p = document.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = LINE_SPACING
    pf.left_indent = Cm(0.74)
    pf.first_line_indent = Cm(-0.74)
    set_run_font(p.add_run(f"[{index}] {entry}"), PT_BODY, TNR, _SONG)
    return p


CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def add_note(document, text, per_page_index=1):
    """注释：数字加圆圈，按页分别编号（每页从①重新开始）。"""
    mark = CIRCLED[per_page_index - 1]
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    set_run_font(p.add_run(f"{mark} {text}"), PT_SMALL_TABLE, TNR, _SONG)
    return p


# --------------------------------------------------------------------------
# 新文档骨架
# --------------------------------------------------------------------------

def new_thesis_document():
    """新建符合规范的论文文档（A4、Word 默认边距、1.5 倍行距、标题样式）。"""
    reset_heading_numbers()
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = TNR
    normal.font.size = Pt(PT_BODY)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), TNR)
    rFonts.set(qn("w:hAnsi"), TNR)
    rFonts.set(qn("w:eastAsia"), _SONG)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = LINE_SPACING
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    for section in doc.sections:
        setup_page(section)
    _ensure_heading_styles(doc)
    return doc


def _ensure_heading_styles(doc):
    """创建 ThesisH1—H4：小4号加粗宋体，并继承 Word 内置 Heading 1—4。

    继承内置 Heading 样式的意义：Word 的自动目录（TOC 域）按大纲级别收录
    段落，基于 Heading 1—4 派生才能生成目录；同时必须覆盖掉内置样式带来的
    大字号、主题字体和蓝色，否则会违反“小4号加粗宋体”的规定。
    """
    from docx.enum.style import WD_STYLE_TYPE
    for level in (1, 2, 3, 4):
        name = f"ThesisH{level}"
        if name in [s.name for s in doc.styles]:
            style = doc.styles[name]
        else:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        builtin_name = f"Heading {level}"
        try:
            style.base_style = doc.styles[builtin_name]
        except KeyError:
            style.base_style = doc.styles["Normal"]

        style.font.name = TNR
        style.font.size = Pt(PT_BODY)
        style.font.bold = True
        style.font.italic = False
        style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # 覆盖内置标题的蓝色

        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        # 清除主题字体引用，否则 Word 仍用主题字体渲染
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rFonts.get(qn(f"w:{attr}")) is not None:
                del rFonts.attrib[qn(f"w:{attr}")]
        rFonts.set(qn("w:ascii"), TNR)
        rFonts.set(qn("w:hAnsi"), TNR)
        rFonts.set(qn("w:eastAsia"), _SONG)  # 正文中标题：小4号加粗宋体
        rFonts.set(qn("w:cs"), TNR)

        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        style.paragraph_format.line_spacing = LINE_SPACING
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(0)
        style.paragraph_format.first_line_indent = Cm(0)

        # 大纲级别，供 Word 自动生成目录
        pPr = style.element.get_or_add_pPr()
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            pPr.append(outline)
        outline.set(qn("w:val"), str(level - 1))
    return doc


def add_toc_field(document, title="目　录"):
    """插入“目　录”题头与 TOC 域（需在 Word 中按 F9 更新）。"""
    add_front_matter_head(document, "目录", PT_TOC_HEAD)
    p = document.add_paragraph()
    set_paragraph_spacing(p, LINE_SPACING)
    _no_first_line_indent(p)
    run = p.add_run()
    set_run_font(run, PT_BODY, TNR, _SONG)
    # 关闭域底纹，避免灰色占位条干扰版面
    rPr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    rPr.append(shd)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-2" \h \z \u '
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "右键此处选择“更新域”生成目录"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr, fld_sep, placeholder, fld_end):
        run._element.append(el)
    return p
