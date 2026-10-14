#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""中文本科论文（设计）格式规范校验器。

按《本科生论文（设计）撰写要求》机械校验 .docx 的字号、字体、边距、
行距、页码格式、标题级次、三线表、图表题注、参考文献与注释。

用法::

    python audit_thesis.py thesis.docx
    python audit_thesis.py thesis.docx --json report.json
    python audit_thesis.py thesis.docx --only page,fonts,headings
    python audit_thesis.py thesis.docx --max-issues 10

退出码：0 无问题；1 存在需修正项；2 文档无法读取。
"""

import argparse
import json
import re
import sys

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Pt
except ImportError:
    sys.stderr.write("需要 python-docx：请使用捆绑 Python 运行本脚本。\n")
    raise SystemExit(2)

# Windows 控制台默认不是 UTF-8；强制中文输出正确显示
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# --------------------------------------------------------------------------
# 规范常量
# --------------------------------------------------------------------------
EMU_CM = 360000.0
SPEC_MARGIN_TB_CM = 2.54
SPEC_MARGIN_LR_CM = 3.17
SPEC_HEADER_CM = 1.5
SPEC_FOOTER_CM = 1.75
SPEC_PAGE_W_CM = 21.0
SPEC_PAGE_H_CM = 29.7
SPEC_LINE_SPACING = 1.5
TOL_CM = 0.03

PT_BODY = 12.0        # 小四
PT_TITLE = 15.0       # 小三
PT_SUBTITLE = 14.0    # 四号
PT_TOC_HEAD = 16.0    # 三号
PT_ABSTRACT_HEAD = 18.0  # 小二号
PT_HEADER = 10.5      # 五号
PT_CAPTION = 10.5     # 五号
PT_SMALL = 9.0        # 小五

SONG, HEI = "SimSun", "SimHei"
TNR = "Times New Roman"

ALL_CHECKS = ["page", "fonts", "headings", "header", "abstract", "toc",
              "tables", "figures", "refs", "formula", "notes"]

CJK = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
LATIN = re.compile(r"[A-Za-z]")

#: 第 3 级规范写法是“1数据来源”（原文不加点号）；“1.”“1．”“1、”一并兼容
H3_SEPARATED = re.compile(r"^\d+[.．、]")
H3_CANONICAL = re.compile(r"^\d{1,2}(?![\d.．、%])\s*\S")

HEADING_PATTERNS = [
    (1, re.compile(r"^[一二三四五六七八九十百]+、")),
    (2, re.compile(r"^（[一二三四五六七八九十百]+）")),
    (3, re.compile(r"^(?:\d+[.．、]|\d{1,2}(?![\d.．、%])\s*\S)")),
    (4, re.compile(r"^（\d+）")),
]
CAPTION_TABLE = re.compile(r"^表\s*\d+")
CAPTION_FIGURE = re.compile(r"^图\s*\d+")
REF_ENTRY = re.compile(r"^\[\d+\]")
REF_TYPE_ID = re.compile(r"\[(M|J|C|R|D|P|S|N|EB/OL)\]")
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
CIRCLED_RE = re.compile("[" + CIRCLED + "]")
FORMULA_RE = re.compile(r"[=＝]")
NUM_HALF_WIDTH = re.compile(r"\d")


# --------------------------------------------------------------------------
# 收集器
# --------------------------------------------------------------------------
class Issue:
    __slots__ = ("check", "severity", "where", "message", "fix")

    def __init__(self, check, severity, where, message, fix=""):
        self.check, self.severity, self.where, self.message, self.fix = (
            check, severity, where, message, fix)

    def as_dict(self):
        return {"check": self.check, "severity": self.severity, "where": self.where,
                "message": self.message, "fix": self.fix}


class Auditor:
    def __init__(self, path):
        self.path = path
        self.doc = Document(path)
        self.issues = []
        self.info = {}

    # ---------------- 工具 ----------------
    def add(self, check, severity, where, message, fix=""):
        self.issues.append(Issue(check, severity, where, message, fix))

    def paragraphs(self):
        return list(self.doc.paragraphs)

    @staticmethod
    def _rpr(run, style):
        """返回 (size_pt, latin_font, ea_font, bold)，从 run 直接格式回退到样式。"""
        size = ascii_f = ea_f = bold = None
        if style is not None:
            try:
                if style.font.size:
                    size = style.font.size.pt
                if style.font.name:
                    ascii_f = style.font.name
                if style.font.bold is not None:
                    bold = style.font.bold
                srPr = style.element.find(qn("w:rPr"))
                if srPr is not None:
                    rf = srPr.find(qn("w:rFonts"))
                    if rf is not None:
                        ascii_f = rf.get(qn("w:ascii")) or ascii_f
                        ea_f = rf.get(qn("w:eastAsia")) or ea_f
            except Exception:
                pass
        rPr = run._element.find(qn("w:rPr"))
        if rPr is not None:
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None:
                ascii_f = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi")) or ascii_f
                ea_f = rf.get(qn("w:eastAsia")) or ea_f
            sz = rPr.find(qn("w:sz"))
            if sz is not None and sz.get(qn("w:val")):
                size = int(sz.get(qn("w:val"))) / 2.0
            b = rPr.find(qn("w:b"))
            if b is not None:
                bold = b.get(qn("w:val")) not in ("0", "false", "none")
        return size, ascii_f, ea_f, bold

    @classmethod
    def eff_props(cls, paragraph):
        """段落首个非空 run 的有效字体属性。"""
        style = paragraph.style
        for run in paragraph.runs:
            if run.text and run.text.strip():
                return cls._rpr(run, style)
        return (None, None, None, None)

    @staticmethod
    def eff_line_spacing(paragraph):
        """有效行距（倍数），从段落回退到样式。"""
        pf = paragraph.paragraph_format
        if pf.line_spacing is not None:
            return pf.line_spacing
        style = paragraph.style
        seen = 0
        while style is not None and seen < 6:
            try:
                if style.paragraph_format.line_spacing is not None:
                    return style.paragraph_format.line_spacing
            except Exception:
                pass
            style = style.base_style
            seen += 1
        return None

    @staticmethod
    def is_body_paragraph(p):
        t = p.text.strip()
        if not t:
            return False
        if CAPTION_TABLE.match(t) or CAPTION_FIGURE.match(t):
            return False
        if REF_ENTRY.match(t):
            return False
        if CIRCLED_RE.match(t):
            return False
        return True

    @classmethod
    def heading_level(cls, p):
        t = p.text.strip()
        if not t or len(t) > 60:
            return None
        for level, pat in HEADING_PATTERNS:
            if not pat.match(t):
                continue
            if level == 3 and not H3_SEPARATED.match(t):
                # “1数据来源”是规范写法，但正文也可能以数字开头（“3个样本……”）；
                # 规范要求正文中标题为小4号加粗宋体，故以“加粗”作为判据避免误判。
                _, _, _, bold = cls.eff_props(p)
                if bold is not True:
                    continue
            return level
        return None


# --------------------------------------------------------------------------
# 各检查项
# --------------------------------------------------------------------------
def check_page(a):
    for i, s in enumerate(a.doc.sections):
        where = f"第{i + 1}节"
        w, h = s.page_width.cm, s.page_height.cm
        if abs(w - SPEC_PAGE_W_CM) > TOL_CM or abs(h - SPEC_PAGE_H_CM) > TOL_CM:
            a.add("page", "error", where, f"纸张为 {w:.2f}×{h:.2f}cm，规范要求 A4（21×29.7cm）",
                  "section.page_width=21cm, page_height=29.7cm")
        for label, got, want in (
            ("上边距", s.top_margin.cm, SPEC_MARGIN_TB_CM),
            ("下边距", s.bottom_margin.cm, SPEC_MARGIN_TB_CM),
            ("左边距", s.left_margin.cm, SPEC_MARGIN_LR_CM),
            ("右边距", s.right_margin.cm, SPEC_MARGIN_LR_CM),
            ("页眉距边界", s.header_distance.cm, SPEC_HEADER_CM),
            ("页脚距边界", s.footer_distance.cm, SPEC_FOOTER_CM),
        ):
            if abs(got - want) > TOL_CM:
                a.add("page", "error", where, f"{label}为 {got:.2f}cm，规范要求 {want}cm",
                      "用 thesisdocx.setup_page(section) 统一设置")

    # 行距
    bad = 0
    for p in a.paragraphs():
        if not a.is_body_paragraph(p):
            continue
        ls = a.eff_line_spacing(p)
        if ls is None or abs(ls - SPEC_LINE_SPACING) > 0.02:
            bad += 1
    if bad:
        a.add("page", "error", "全文", f"{bad} 个正文段落行距不是 1.5 倍",
              "thesisdocx.set_paragraph_spacing(p, 1.5)")

    # 页码格式：应有 大写罗马 + 阿拉伯 两节
    fmts = []
    for i, s in enumerate(a.doc.sections):
        pg = s._sectPr.find(qn("w:pgNumType"))
        fmts.append(pg.get(qn("w:fmt")) if pg is not None else "decimal")
    a.info["page_number_formats"] = fmts
    if "upperRoman" not in fmts:
        a.add("page", "warn", "页码", "未发现大写罗马数字页码节（声明/开题/目录/摘要应为Ⅰ、Ⅱ、Ⅲ）",
              "thesisdocx.setup_front_matter_page_numbers(doc)")
    if len(fmts) < 2:
        a.add("page", "warn", "页码", "只有一个节，无法区分罗马数字与阿拉伯数字页码；正文应另起一节",
              "thesisdocx.start_body_section(doc)")

    # 页码居中
    for i, s in enumerate(a.doc.sections):
        p = s.footer.paragraphs[0] if s.footer.paragraphs else None
        has_page = p is not None and "PAGE" in p._p.xml
        if not has_page:
            if is_cover_like(a, i):
                a.add("page", "info", f"第{i + 1}节页脚",
                      "该节无页码，判断为封面/扉页节（规范允许封面和扉页不编页码）",
                      "确认封面、扉页后另起一节再开始编页码")
            else:
                a.add("page", "error", f"第{i + 1}节页脚", "页脚中没有页码域（PAGE）",
                      "thesisdocx.add_centered_page_number(section)")
        elif str(p.alignment) not in ("CENTER (1)",) and "w:jc" not in p._p.xml:
            a.add("page", "warn", f"第{i + 1}节页脚", "页码可能未居中",
                  "p.alignment = WD_ALIGN_PARAGRAPH.CENTER")


def is_cover_like(a, section_index):
    """判断某节是否为封面/扉页节：本节无页码、无页眉，且含大字号居中标（题名）。

    非末节的 sectPr 挂在上一段落（或最后一段）的 pPr 里，因此先回到段落再向前扫。
    """
    s = a.doc.sections[section_index]
    footer_p = s.footer.paragraphs[0] if s.footer.paragraphs else None
    if footer_p is not None and "PAGE" in footer_p._p.xml:
        return False
    header_xml = s.header._element.xml if s.header is not None else ""
    if "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", header_xml, re.S)).strip():
        return False

    el = s._sectPr.getparent()          # 非末节为 w:pPr，末节为 w:body
    if el is not None and el.tag.endswith("}pPr"):
        el = el.getparent()             # 回到该 w:p
    scanned = 0
    while el is not None and scanned < 40:
        if el.tag.endswith("}p"):
            text = "".join(el.itertext()).strip()
            szs = [int(m) / 2.0 for m in re.findall(r'<w:sz w:val="(\d+)"', el.xml)]
            if text and CJK.search(text) and szs:
                # 封面含论文题目（小3号黑体=15pt）；题名/题头类大字号或
                # 正文以外的零散文字都视为封面特征
                if max(szs) >= 15 or szs.count(12.0) <= 2:
                    return True
        el = el.getprevious()
        if el is not None:
            scanned += 1
    return False


def check_fonts(a):
    # 题头/题名/题注/参考文献条目另有规定字号，不按正文字号提示
    oversized_ok = re.compile(
        r"^(摘\s*要|目\s*录|Abstract|Keywords|关键词|声\s*明|致\s*谢|参考文献|附录|"
        r"表\s*\d+|图\s*\d+|\[\d+\])")
    for p in a.paragraphs():
        t = p.text.strip()
        if not t:
            continue
        size, ascii_f, ea_f, bold = a.eff_props(p)
        where = f"“{t[:24]}…”"
        has_cjk = bool(CJK.search(t))
        if has_cjk and ea_f is None:
            a.add("fonts", "warn", where, "中文段落未设置中文字体（w:eastAsia），字形不会真正变为宋体/黑体",
                  "thesisdocx.set_run_font(run, 12, 'Times New Roman', 'SimSun')")
            continue
        if has_cjk and ea_f not in (SONG, HEI):
            a.add("fonts", "warn", where, f"中文字体为 {ea_f}，规范中正文/标题应为宋体（SimSun）",
                  "改用 SimSun；黑体仅用于论文题目")
        if LATIN.search(t) and ascii_f and ascii_f not in (TNR,):
            a.add("fonts", "warn", where, f"西文字体为 {ascii_f}，规范要求 Times New Roman", "改用 TNR")
        if size is None:
            continue
        if (size > PT_BODY + 0.01 and a.heading_level(p) is None
                and len(t) < 60 and not oversized_ok.match(t)):
            a.add("fonts", "info", where,
                  f"字号 {size}pt 大于正文小四（12pt）；确认这是题名/题头而非正文",
                  "正文用 12pt")


def check_headings(a):
    levels = []
    for p in a.paragraphs():
        lvl = a.heading_level(p)
        if lvl is None:
            continue
        t = p.text.strip()
        levels.append((lvl, t, p))
        size, ascii_f, ea_f, bold = a.eff_props(p)
        where = f"“{t[:24]}”"
        if size is not None and abs(size - PT_BODY) > 0.01:
            a.add("headings", "error", where, f"字号 {size}pt，规范要求正文中标题为小4号（12pt）",
                  "thesisdocx.add_heading(doc, level, text)")
        if ea_f not in (None, SONG):
            a.add("headings", "warn", where, f"中文字体 {ea_f}，规范要求宋体", "改用 SimSun")
        if bold is not True:
            a.add("headings", "error", where, "标题未加粗，规范要求小4号加粗宋体", "run.font.bold = True")

    # 级次顺序：不得跳级
    prev = 0
    for lvl, t, _ in levels:
        if prev and lvl > prev + 1:
            a.add("headings", "error", f"“{t[:24]}”",
                  f"标题从 {prev} 级跳到 {lvl} 级", "补齐中间级次或调整编号")
        prev = lvl

    # 一、二级标题上下空一行
    body = a.paragraphs()
    for idx, p in enumerate(body):
        lvl = a.heading_level(p)
        if lvl not in (1, 2):
            continue
        before = body[idx - 1].text.strip() if idx > 0 else None
        after = body[idx + 1].text.strip() if idx + 1 < len(body) else None
        if before != "":
            a.add("headings", "error", f"“{p.text.strip()[:24]}”",
                  "一、二级标题上方未空一行", "thesisdocx.add_heading(...) 会自动插入空行")
        if after != "":
            a.add("headings", "error", f"“{p.text.strip()[:24]}”",
                  "一、二级标题下方未空一行", "thesisdocx.add_heading(...) 会自动插入空行")
    a.info["heading_count"] = len(levels)


def check_header(a):
    for i, s in enumerate(a.doc.sections):
        xml = s.header._element.xml if s.header is not None else ""
        texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, re.S)
        joined = "".join(texts).strip()
        if not joined:
            if s.different_first_page_header_footer:
                continue
            if is_cover_like(a, i):
                a.add("header", "info", f"第{i + 1}节",
                      "该节无页眉，判断为封面/扉页节（规范允许封面和扉页不设页眉）",
                      "确认前置部分（声明/目录/摘要）所在节已设页眉")
            else:
                a.add("header", "error", f"第{i + 1}节",
                      "该节没有页眉文字（除封面、扉页外各页均应有页眉）",
                      "thesisdocx.set_header_text(section, '一、引言')")
            continue
        if "PAGE" in joined or not joined:
            a.add("header", "warn", f"第{i + 1}节", f"页眉内容可疑：“{joined[:30]}”", "页眉应以本页一级标题为内容")
        # 页眉字体 5 号宋体
        rpr = re.search(r"<w:rPr>.*?</w:rPr>", xml, re.S)
        if rpr:
            sz = re.search(r'<w:sz w:val="(\d+)"', rpr.group(0))
            if sz and abs(int(sz.group(1)) / 2.0 - PT_HEADER) > 0.01:
                a.add("header", "error", f"第{i + 1}节",
                      f"页眉字号 {int(sz.group(1)) / 2.0}pt，规范要求 5号（10.5pt）",
                      "set_header_text(..., size_pt=10.5)")
            ea = re.search(r'w:eastAsia="([^"]+)"', rpr.group(0))
            if ea and ea.group(1) != SONG:
                a.add("header", "error", f"第{i + 1}节",
                      f"页眉中文字体 {ea.group(1)}，规范要求宋体", "改用 SimSun")
        if "w:jc" not in (s.header.paragraphs[0]._p.xml if s.header.paragraphs else ""):
            a.add("header", "warn", f"第{i + 1}节", "页眉可能未居中",
                  "p.alignment = WD_ALIGN_PARAGRAPH.CENTER")


def check_abstract(a):
    paras = a.paragraphs()
    texts = [p.text.strip() for p in paras]
    joined = "\n".join(texts)

    # 中文摘要题头
    idx = next((i for i, t in enumerate(texts) if re.match(r"^摘\s*要", t)), None)
    if idx is None:
        a.add("abstract", "warn", "全文", "未找到中文摘要题头“摘　要”", "thesisdocx.add_front_matter_head(doc, '摘要')")
    else:
        p = paras[idx]
        size, _, ea, bold = a.eff_props(p)
        w = f"摘要题头“{texts[idx]}”"
        if "\u3000" not in texts[idx]:
            a.add("abstract", "error", w, "“摘要”两字之间未空一个全角空格", "写作“摘　要”")
        if size is not None and abs(size - PT_ABSTRACT_HEAD) > 0.01:
            a.add("abstract", "error", w, f"字号 {size}pt，规范要求小2号（18pt）", "size_pt=18")
        if bold is not True:
            a.add("abstract", "error", w, "题头未加粗", "run.font.bold = True")
        if texts[idx + 1:idx + 2] and texts[idx + 1] != "":
            a.add("abstract", "error", w, "题头后未隔一行再写摘要正文", "题头后插入空段落")

    # 关键词
    kw = next((i for i, t in enumerate(texts) if t.startswith("关键词")), None)
    if kw is None:
        a.add("abstract", "warn", "全文", "未找到“关键词”行", "thesisdocx.add_keywords(doc, '…')")
    else:
        p = paras[kw]
        label = p.runs[0] if p.runs else None
        if label is not None:
            size, _, ea, _ = a._rpr(label, p.style)
            if size is not None and abs(size - PT_SUBTITLE) > 0.01:
                a.add("abstract", "error", "关键词", f"“关键词”字号 {size}pt，规范要求4号（14pt）", "size_pt=14")
            if ea not in (None, HEI):
                a.add("abstract", "error", "关键词", f"“关键词”字体 {ea}，规范要求黑体（SimHei）", "改用 SimHei")
        if "\u3000" not in p.text:
            a.add("abstract", "error", "关键词", "“关键词”与其后内容之间未空一个全角空格", "写作“关键词　…”")

    # 英文摘要
    if "Abstract" in joined:
        p = next(pp for pp in paras if "Abstract" in pp.text)
        size, ascii_f, _, bold = a.eff_props(p)
        if bold is not True:
            a.add("abstract", "error", "Abstract", "“Abstract”未加粗", "run.font.bold = True")
        if ascii_f and ascii_f != TNR:
            a.add("abstract", "error", "Abstract", f"西文字体 {ascii_f}，规范要求 Times New Roman", "改用 TNR")
        if size is not None and abs(size - PT_BODY) > 0.01:
            a.add("abstract", "error", "Abstract", f"字号 {size}pt，规范要求小4号（12pt）", "size_pt=12")
    if "Keywords" in joined:
        p = next(pp for pp in paras if "Keywords" in pp.text)
        _, _, _, bold = a.eff_props(p)
        if bold is not True:
            a.add("abstract", "error", "Keywords", "“Keywords”未加粗", "run.font.bold = True")


def check_toc(a):
    texts = [p.text.strip() for p in a.paragraphs()]
    idx = next((i for i, t in enumerate(texts) if re.match(r"^目\s*录", t)), None)
    if idx is None:
        a.add("toc", "warn", "全文", "未找到目录题头“目　录”", "thesisdocx.add_toc_field(doc)")
        return
    if "\u3000" not in texts[idx]:
        a.add("toc", "error", f"目录题头“{texts[idx]}”", "“目录”两字之间未空一个全角空格", "写作“目　录”")
    p = a.paragraphs()[idx]
    size, _, _, bold = a.eff_props(p)
    if size is not None and abs(size - PT_TOC_HEAD) > 0.01:
        a.add("toc", "error", "目录题头", f"字号 {size}pt，规范要求3号（16pt）", "size_pt=16")
    if bold is not True:
        a.add("toc", "error", "目录题头", "题头未加粗", "run.font.bold = True")
    # 目录条目：Word 自动目录用 TOC 1/TOC 2 样式；手工目录用直接格式。
    # 规范要求一级题序及标题用小4号黑体（SimHei），其余小4号宋体。
    # 只在目录区域内扫描，遇到正文段落（有首行缩进）即停止，避免把正文标题当目录条目。
    auto_toc = 0
    for i in range(idx + 1, min(len(texts), idx + 60)):
        t = texts[i]
        if not t:
            continue
        p = a.paragraphs()[i]
        style_name = p.style.name if p.style is not None else ""
        # 只认 Word 自动目录的 TOC n 段落样式；不要用“xml 含 TOC”，
        # 否则会把自己那条 TOC 域占位段落误判为目录条目
        is_auto = bool(re.match(r"^TOC\s*\d", style_name))
        if not is_auto and "w:fldChar" in p._p.xml:
            continue      # 跳过 TOC 域本身的占位段落
        if not is_auto:
            if style_name in ("ThesisH1", "ThesisH2", "ThesisH3", "ThesisH4") or \
                    re.match(r"^(Heading|标题)\s*\d", style_name):
                continue  # 正文标题本身不是目录条目
            ind = p._p.find(qn("w:pPr"))
            indented = ind is not None and ind.find(qn("w:ind")) is not None and \
                (ind.find(qn("w:ind")).get(qn("w:firstLineChars")) or
                 ind.find(qn("w:ind")).get(qn("w:firstLine")))
            if indented:
                break       # 进入正文，目录区域结束
        is_l1 = bool(re.match(r"^[一二三四五六七八九十百]+、", t))
        if is_auto:
            auto_toc += 1
            size = p.style.font.size.pt if (p.style is not None and p.style.font.size) else None
            if size is not None and abs(size - PT_BODY) > 0.01:
                a.add("toc", "warn", f"目录条目“{t[:20]}”",
                      f"目录样式字号 {size}pt，规范要求小4号（12pt）",
                      "修改“TOC 1/TOC 2”样式的字号为 12pt")
            if is_l1 and p.style.font.name not in (None, HEI):
                a.add("toc", "info", f"目录条目“{t[:20]}”",
                      "一级标题条目宜用小4号黑体（Word 自动目录可用“TOC 1”样式设置）",
                      "把“TOC 1”样式的中文字体改为 SimHei")
            continue
        if is_l1:
            size, _, ea, _ = a.eff_props(p)
            if ea not in (None, HEI):
                a.add("toc", "warn", f"目录条目“{t[:20]}”",
                      f"一级标题条目字体 {ea}，规范要求小4号黑体",
                      "改用 SimHei（thesisdocx.set_run_font(run, 12, 'Times New Roman', 'SimHei')）")
        elif re.match(r"^(（[一二三四五六七八九十百]+）|\d+[.．、]|（\d+）)", t):
            _, _, ea, _ = a.eff_props(p)
            if ea not in (None, SONG):
                a.add("toc", "warn", f"目录条目“{t[:20]}”",
                      f"目录条目字体 {ea}，规范要求小4号宋体", "改用 SimSun")
    if auto_toc:
        a.add("toc", "info", "目录",
              f"检测到 {auto_toc} 条 Word 自动目录条目；目录需在 Word 中按 F9 更新后才反映最新标题",
              "在 Word 中右键目录 → 更新域 → 更新整个目录")


def check_tables(a):
    for ti, table in enumerate(a.doc.tables, 1):
        where = f"第{ti}个表"
        rows = table.rows
        if not rows:
            continue
        # 左右边线必须为 none
        for ri, row in enumerate(rows):
            for cell in row.cells:
                xml = cell._tc.xml
                for edge in ("left", "right"):
                    m = re.search(r'<w:' + edge + r'[^>]*w:val="([^"]+)"', xml)
                    if m and m.group(1) not in ("none", "nil"):
                        a.add("tables", "error", where,
                              f"表格存在{('左' if edge == 'left' else '右')}边线，规范要求三线表不加左、右边线",
                              "thesisdocx.apply_three_line_table(table)")
                        break
        # 上下粗线 1.5 磅
        first, last = rows[0], rows[-1]
        for edge, row, label in (("top", first, "上"), ("bottom", last, "下")):
            xml = row.cells[0]._tc.xml
            m = re.search(r'<w:' + edge + r'[^>]*w:val="([^"]+)"[^>]*w:sz="(\d+)"', xml)
            if not m:
                m2 = re.search(r'<w:' + edge + r'[^>]*w:sz="(\d+)"', xml)
                if m2:
                    sz = int(m2.group(1))
                else:
                    a.add("tables", "error", where, f"表格缺{label}边线（应为1.5磅粗实线）",
                          "thesisdocx.apply_three_line_table(table)")
                    continue
            else:
                sz = int(m.group(2))
            if sz != 12:
                a.add("tables", "error", where,
                      f"表格{label}边线 {sz / 8:.2f}磅，规范要求 1.5磅", "BORDER_THICK = 12")
        # 表内文字小5号
        for row in rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if not p.text.strip():
                        continue
                    size, _, _, _ = a.eff_props(p)
                    if size is not None and size > PT_SMALL + 0.01:
                        a.add("tables", "warn", where,
                              f"表内文字 {size}pt，规范要求小5号（9pt）", "thesisdocx.style_table_text(table)")
                        break

    # 表序表名在表上
    body = a.doc.element.body
    paras = a.paragraphs()
    for ti, table in enumerate(a.doc.tables, 1):
        t_el = table._tbl
        # 找表格前一个段落
        prev = t_el.getprevious()
        prev_text = ""
        while prev is not None and not prev.tag.endswith("}p"):
            prev = prev.getprevious()
        if prev is not None:
            prev_text = "".join(prev.itertext()).strip()
        if not CAPTION_TABLE.match(prev_text):
            a.add("tables", "error", f"第{ti}个表",
                  f"表格上方未找到表序表名（规范要求表序表名置于表上、居中）",
                  "thesisdocx.add_table_caption(doc, i, '名称') 必须放在 add_table 之前")
        else:
            p = next((pp for pp in paras if pp.text.strip() == prev_text), None)
            if p is not None:
                size, _, ea, _ = a.eff_props(p)
                if size is not None and abs(size - PT_CAPTION) > 0.01:
                    a.add("tables", "error", f"表序表名“{prev_text[:20]}”",
                          f"字号 {size}pt，规范要求5号（10.5pt）", "size_pt=10.5")
                if "\u3000" in prev_text or prev_text.endswith("。") or prev_text.endswith("，") or prev_text.endswith("."):
                    a.add("tables", "error", f"表序表名“{prev_text[:20]}”",
                          "表名后不加标点，或表序与表名之间应空一个半角空格而非全角空格", "写作“表1 名称”")


def check_figures(a):
    doc_xml = a.doc.element.body.xml
    has_image = "<w:drawing" in doc_xml or "<w:pict" in doc_xml
    texts = [p.text.strip() for p in a.paragraphs()]
    fig_captions = [t for t in texts if CAPTION_FIGURE.match(t)]
    if has_image and not fig_captions:
        a.add("figures", "error", "全文", "文档含插图但未找到图号图名",
              "thesisdocx.add_figure_caption(doc, i, '名称')")
    for t in fig_captions:
        p = next(pp for pp in a.paragraphs() if pp.text.strip() == t)
        size, _, ea, _ = a.eff_props(p)
        if size is not None and abs(size - PT_CAPTION) > 0.01:
            a.add("figures", "error", f"图号图名“{t[:20]}”", f"字号 {size}pt，规范要求5号（10.5pt）", "size_pt=10.5")
        if "\u3000" in t:
            a.add("figures", "error", f"图号图名“{t[:20]}”", "图名与图号之间应空一个半角空格，不要用全角空格", "写作“图1 名称”")
        if str(p.alignment) != "CENTER (1)" and "w:jc" not in p._p.xml:
            a.add("figures", "error", f"图号图名“{t[:20]}”", "图号图名未居中", "p.alignment = CENTER")
        # 图名应在图下：检查前一个元素是否为图片
        prev = p._p.getprevious()
        before = ""
        while prev is not None and not prev.tag.endswith("}p"):
            prev = prev.getprevious()
        if prev is not None:
            before = prev.xml
        if "<w:drawing" not in before and "<w:pict" not in before:
            a.add("figures", "warn", f"图号图名“{t[:20]}”", "该题注上方未发现插图，确认图名置于图下", "先插入图片再写图名")
        if "[1]" not in t and "[" not in t:
            a.add("figures", "info", f"图号图名“{t[:20]}”",
                  "引用图必须在图名右上角加引用文献号", "add_figure_caption(..., cite=1)")
    if has_image:
        a.add("figures", "info", "插图", "有数字标注的坐标图必须标明坐标单位；图内文字应为小5号宋体（需在制图时设定）",
              "人工核对")


def check_refs(a):
    texts = [p.text.strip() for p in a.paragraphs()]
    entries = [t for t in texts if REF_ENTRY.match(t)]
    if not entries:
        if any(t in ("参考文献", "参考文献：") for t in texts):
            a.add("refs", "error", "参考文献", "参考文献节存在但没有 [1] 形式的条目",
                  "thesisdocx.add_reference(doc, 1, '著者.书名[M].出版者，出版年.页码.')")
        return
    a.info["reference_count"] = len(entries)
    prev_n = 0
    for t in entries:
        n = int(re.match(r"^\[(\d+)\]", t).group(1))
        if n != prev_n + 1:
            a.add("refs", "error", f"“{t[:28]}”", f"序号不连续：{prev_n} 之后为 {n}", "按 1、2、3… 顺序编号")
        prev_n = n
        if not REF_TYPE_ID.search(t):
            a.add("refs", "error", f"“{t[:28]}”",
                  "缺少文献类型标识（[M]/[J]/[C]/[R]/[D]/[P]/[S]/[N]/[EB/OL]）",
                  "在书名或篇名之后加类型标识")
        if not t.rstrip().endswith("."):
            a.add("refs", "warn", f"“{t[:28]}”", "条目末尾应以英文句点结束", "补 “.”")
        if "\uff1a" in t and "[D]" in t:
            a.add("refs", "info", f"“{t[:28]}”", "学位论文格式为“……学位论文，授予单位，授予年.”，注意逗号", "核对著录格式")


def check_formula(a):
    formulas = 0
    for p in a.paragraphs():
        t = p.text.strip()
        if not t or len(t) > 80:
            continue
        if "=" not in t and "＝" not in t:
            continue
        if not re.search(r"[A-Za-z0-9\u4e00-\u9fff]\s*[=＝]", t):
            continue
        formulas += 1
        where = f"“{t[:28]}”"
        if t.endswith(("。", "，", ",", ";", "；", ".", "、")) and not t.endswith("."):
            a.add("formula", "error", where, "公式末不应加标点", "去掉公式末尾标点")
        if str(p.alignment) != "CENTER (1)" and "w:jc" not in p._p.xml and "\t" not in t:
            a.add("formula", "warn", where, "公式原则上应居中书写", "p.alignment = WD_ALIGN_PARAGRAPH.CENTER")
    a.info["formula_like_paragraphs"] = formulas
    if formulas == 0:
        a.add("formula", "info", "全文", "未识别到公式段落；如有公式请人工核对是否居中、末尾无标点、长公式在“＝”处转行", "人工核对")


def check_notes(a):
    marks = [t for t in (p.text.strip() for p in a.paragraphs()) if CIRCLED_RE.match(t)]
    if marks:
        a.add("notes", "info", "注释",
              f"发现 {len(marks)} 处圆圈注释序号；规范要求按页分别编号（每页从①重新开始），本工具无法判定分页",
              "人工核对每页是否从①重新编号，且注释文字在当页页脚")
    # 全篇连续编号的迹象
    seq = []
    for t in marks:
        seq.append(CIRCLED.index(CIRCLED_RE.search(t).group(0)) + 1)
    if len(seq) > 3 and seq == sorted(seq) and seq[-1] > 5:
        a.add("notes", "warn", "注释", f"注释序号疑似全篇连续编号（{seq[:8]}…），规范要求按页分别编号",
              "每页从①重新开始")


CHECKS = {
    "page": check_page, "fonts": check_fonts, "headings": check_headings,
    "header": check_header, "abstract": check_abstract, "toc": check_toc,
    "tables": check_tables, "figures": check_figures, "refs": check_refs,
    "formula": check_formula, "notes": check_notes,
}


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
SEV_ORDER = {"error": 0, "warn": 1, "info": 2}
SEV_LABEL = {"error": "错误", "warn": "警告", "info": "提示"}


def main(argv=None):
    ap = argparse.ArgumentParser(description="按《本科生论文（设计）撰写要求》校验 docx")
    ap.add_argument("document", help="待校验的 .docx 路径")
    ap.add_argument("--only", default="", help=f"只运行指定检查项，逗号分隔：{','.join(ALL_CHECKS)}")
    ap.add_argument("--json", dest="json_path", default=None, help="输出 JSON 报告的路径")
    ap.add_argument("--max-issues", type=int, default=25, help="每项最多列出多少条（默认 25）")
    args = ap.parse_args(argv)

    selected = [c.strip() for c in args.only.split(",") if c.strip()] or ALL_CHECKS
    unknown = [c for c in selected if c not in CHECKS]
    if unknown:
        ap.error(f"未知检查项：{','.join(unknown)}；可用：{','.join(ALL_CHECKS)}")

    try:
        a = Auditor(args.document)
    except Exception as exc:
        sys.stderr.write(f"无法读取文档：{exc}\n")
        return 2

    for name in selected:
        try:
            CHECKS[name](a)
        except Exception as exc:  # 单项失败不影响其它项
            a.add(name, "warn", name, f"检查项执行失败：{exc}", "请人工核对此项")

    for name in selected:
        items = sorted([i for i in a.issues if i.check == name],
                       key=lambda i: SEV_ORDER[i.severity])
        print(f"\n== {name} ==")
        if not items:
            print("  OK  未发现问题")
            continue
        for i in items[:args.max_issues]:
            print(f"  [{SEV_LABEL[i.severity]}] {i.where}：{i.message}")
            if i.fix:
                print(f"          → {i.fix}")
        if len(items) > args.max_issues:
            print(f"  …… 另有 {len(items) - args.max_issues} 条同类问题（用 --max-issues 调整）")

    errors = sum(1 for i in a.issues if i.severity == "error")
    warns = sum(1 for i in a.issues if i.severity == "warn")
    infos = sum(1 for i in a.issues if i.severity == "info")
    print(f"\n=== 汇总 ===")
    print(f"错误 {errors} 项，警告 {warns} 项，提示 {infos} 项")
    print("提示：内容准确性、页眉是否等于本页一级标题、图内文字字号、坐标单位、注释按页重编需人工核对。")

    if args.json_path:
        payload = {
            "document": args.document,
            "summary": {"error": errors, "warn": warns, "info": infos},
            "info": a.info,
            "issues": [i.as_dict() for i in a.issues],
        }
        with open(args.json_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        print(f"JSON 报告已写入 {args.json_path}")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
