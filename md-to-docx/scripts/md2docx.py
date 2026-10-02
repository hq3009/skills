#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx.py —— 将 Markdown(.md) 转换为符合中文规范的 Word(.docx) 文档。

支持两种版式:
    official : 党政机关公文格式 (GB/T 9704-2012)
    zju      : 浙江大学研究生学位论文编写规则

标题层级映射 (按 len(docx 内置样式) 与 markdown 的对应关系):
    md #      -> docx Title      （中文 Word 显示为「标题」）
    md ##     -> docx Heading 1  （中文 Word 显示为「标题 1」）
    md ###    -> docx Heading 2  （「标题 2」）
    md ####   -> docx Heading 3  （「标题 3」）
    md #####  -> docx Heading 4  （「标题 4」）
    md ###### -> docx Heading 5  （「标题 5」）

样式名兼容：文件里内置样式的 w:name 仍是英文，中文 Word/WPS 界面显示为中文名，
部分国产办公软件保存的文档会直接写成中文名，脚本两种都能识别（见 resolve_style）。

分节符（独占一行的 HTML 注释指令，md 渲染时不可见）:
    <!-- 分节符 -->                  -> 下一页分节符（默认）
    <!-- 分节符: 连续 -->            -> 连续分节符
    <!-- 分节符: 偶数页 -->          -> 偶数页分节符
    <!-- 分节符: 奇数页 -->          -> 奇数页分节符
    <!-- section-break: continuous -->  英文写法，类型同上
    其余 <!-- ... --> 注释（含行内）一律不写入正文。

用法示例:
    python md2docx.py input.md -o output.docx --mode official
    python md2docx.py input.md -o output.docx --mode zju --check-fonts
"""

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # pragma: no cover
        pass

try:
    from docx import Document
    from docx.enum.section import WD_SECTION_START
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "缺少依赖 python-docx，请先安装:\n"
        "    <python> -m pip install python-docx\n"
    )
    raise SystemExit(2)


# --------------------------------------------------------------------------- #
# 0. 常量
# --------------------------------------------------------------------------- #

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

HEADER_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)
FOOTER_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
)

CN_FONT_SIZES: Dict[str, float] = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

CIRCLED_DIGITS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

BULLET_STYLES = {
    "dash": ["—", "○", "·", "·"],
    "dot": ["●", "○", "■", "□"],
    "disc": ["•", "◦", "▪", "▫"],
}

ORDINAL_STYLES = {
    # 层级 -> 编号模板
    "arabic": ["{}.", "（{}）", "{})"],
    "chinese": ["{}、", "（{}）", "{}.", "（{}）"],
    "circle": [CIRCLED_DIGITS, "（{}）", "{})"],
}


def _size(value: Any) -> float:
    """接受 16 / '三号' / '16pt' 形式的字号，统一返回磅值。"""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in CN_FONT_SIZES:
        return CN_FONT_SIZES[text]
    return float(text.lower().replace("pt", "").strip())


def escape_xml(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# --------------------------------------------------------------------------- #
# 1. XML 辅助: 字体 / 缩进 / 行距 / 样式
# --------------------------------------------------------------------------- #

def _rpr_xml(cn: str, en: str, size_pt: float, bold: bool, italic: bool,
             strike: bool = False) -> str:
    half = int(round(size_pt * 2))
    parts = [
        f'<w:rFonts w:ascii="{escape_xml(en)}" w:hAnsi="{escape_xml(en)}" '
        f'w:eastAsia="{escape_xml(cn)}" w:cs="{escape_xml(en)}"/>'
    ]
    if bold:
        parts.append("<w:b/><w:bCs/>")
    if italic:
        parts.append("<w:i/><w:iCs/>")
    if strike:
        parts.append("<w:strike/>")
    parts.append('<w:color w:val="000000"/>')
    parts.append(f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def set_run_font(run, cn: str, en: str, size_pt: float, bold: bool = False,
                 italic: bool = False, underline: bool = False,
                 strike: bool = False) -> None:
    """给 run 设置中英文字体、字号、字重，颜色恒为黑色。"""
    font = run.font
    if en:
        font.name = en
    if size_pt:
        font.size = Pt(size_pt)
    font.bold = bool(bold)
    font.italic = bool(italic)
    font.underline = bool(underline)
    font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    if cn:
        rfonts.set(qn("w:eastAsia"), cn)
    if en:
        rfonts.set(qn("w:ascii"), en)
        rfonts.set(qn("w:hAnsi"), en)
        rfonts.set(qn("w:cs"), en)
    if strike:
        node = rpr.find(qn("w:strike"))
        if node is None:
            rpr.append(OxmlElement("w:strike"))


def set_para_indent(p, first_line_chars: Optional[float] = None,
                    hanging_chars: Optional[float] = None,
                    left_chars: Optional[float] = None,
                    right_chars: Optional[float] = None) -> None:
    """按「字符」设置缩进(写入 firstLineChars / hangingChars / leftChars / rightChars)。"""
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
    if hanging_chars is not None:
        ind.set(qn("w:hangingChars"), str(int(round(hanging_chars * 100))))
        ind.set(qn("w:hanging"), str(int(round(hanging_chars * size_to_twips_ref(p)))))
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
        ind.set(qn("w:left"), str(int(round(left_chars * size_to_twips_ref(p)))))
    if right_chars is not None:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
        ind.set(qn("w:right"), str(int(round(right_chars * size_to_twips_ref(p)))))


def size_to_twips_ref(p) -> int:
    """以段落现有字号估算 1 字符宽度(磅->twips)，用于给 Chars 缩进兜底。"""
    size_pt = 12.0
    try:
        rpr = p._p.pPr.rPr if (p._p.pPr is not None and p._p.pPr.rPr is not None) else None
        if rpr is not None and rpr.sz is not None and rpr.sz.val is not None:
            size_pt = rpr.sz.val / 2.0
        elif p.style is not None and p.style.font.size is not None:
            size_pt = p.style.font.size.pt
    except Exception:
        pass
    return int(round(size_pt * 20))


def set_line_spacing(p, multiples: Optional[float] = None,
                     exact_pt: Optional[float] = None,
                     rule: str = "exact") -> None:
    ppr = p._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if multiples is not None:
        spacing.set(qn("w:line"), str(int(round(multiples * 240))))
        spacing.set(qn("w:lineRule"), "auto")
    elif exact_pt is not None:
        spacing.set(qn("w:line"), str(int(round(exact_pt * 20))))
        spacing.set(qn("w:lineRule"), "exact" if rule == "exact" else "atLeast")


def set_spacing_around(p, before_pt: Optional[float] = None,
                       after_pt: Optional[float] = None) -> None:
    ppr = p._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if before_pt is not None:
        spacing.set(qn("w:before"), str(int(round(before_pt * 20))))
    if after_pt is not None:
        spacing.set(qn("w:after"), str(int(round(after_pt * 20))))


def set_keep(p, keep_next: bool = False, page_break_before: bool = False) -> None:
    ppr = p._p.get_or_add_pPr()
    if keep_next:
        node = ppr.find(qn("w:keepNext"))
        if node is None:
            ppr.append(OxmlElement("w:keepNext"))
    if page_break_before:
        node = ppr.find(qn("w:pageBreakBefore"))
        if node is None:
            ppr.append(OxmlElement("w:pageBreakBefore"))


# --------------------------------------------------------------------------- #
# 2. 样式配置
# --------------------------------------------------------------------------- #

@dataclass
class PStyle:
    """一个段落样式(含字体)的完整描述。"""

    cn: str = "仿宋"
    en: str = "Times New Roman"
    size_pt: float = 16.0
    bold: bool = False
    italic: bool = False
    align: Optional[str] = "justify"
    first_line_chars: float = 0.0
    left_chars: float = 0.0
    hanging_chars: float = 0.0
    before_pt: float = 0.0
    after_pt: float = 0.0
    line_pt: Optional[float] = None
    line_multiple: Optional[float] = None
    line_rule: str = "exact"
    keep_next: bool = False
    page_break_before: bool = False


@dataclass
class PageSpec:
    width_cm: float = 21.0
    height_cm: float = 29.7
    top_cm: float = 3.7
    bottom_cm: float = 3.5
    left_cm: float = 2.8
    right_cm: float = 2.6
    header_cm: float = 2.5
    footer_cm: float = 2.5


@dataclass
class Config:
    mode: str = "official"
    page: PageSpec = field(default_factory=PageSpec)

    # 正文与各级标题
    body: PStyle = field(default_factory=PStyle)
    title: PStyle = field(default_factory=PStyle)
    h1: PStyle = field(default_factory=PStyle)
    h2: PStyle = field(default_factory=PStyle)
    h3: PStyle = field(default_factory=PStyle)
    h4: PStyle = field(default_factory=PStyle)
    h5: PStyle = field(default_factory=PStyle)

    # 其他块级
    quote: PStyle = field(default_factory=PStyle)
    listitem: PStyle = field(default_factory=PStyle)
    code: PStyle = field(default_factory=PStyle)
    table_title: PStyle = field(default_factory=PStyle)
    figure_title: PStyle = field(default_factory=PStyle)

    # 表格
    table_cell_cn: str = "宋体"
    table_cell_en: str = "Times New Roman"
    table_cell_size: float = 10.5
    table_cell_align: str = "center"
    table_header_bold: bool = True
    table_align: str = "center"
    table_style: Optional[str] = "Table Grid"

    # 页码
    page_number: bool = True
    page_number_odd_even: bool = False
    page_number_cn: str = "宋体"
    page_number_en: str = "Times New Roman"
    page_number_size: float = 14.0
    page_number_dash: bool = True

    # 页眉
    header_enabled: bool = False
    header_left: str = ""
    header_right: str = ""
    header_style_ref: Optional[str] = None   # 例如 "Heading 1"
    header_cn: str = "宋体"
    header_en: str = "Times New Roman"
    header_size: float = 9.0
    header_rule: bool = True

    # 杂项
    chapter_page_break: bool = False
    bullet_markers: List[str] = field(default_factory=lambda: list(BULLET_STYLES["dash"]))
    ordinal_patterns: List[str] = field(
        default_factory=lambda: list(ORDINAL_STYLES["arabic"])
    )
    auto_number: bool = False
    auto_number_patterns: List[str] = field(
        default_factory=lambda: ["一、", "（一）", "{}.", "（{}）"]
    )

    def heading_style(self, level: int) -> PStyle:
        """level: docx 标题级别 1..5"""
        return {1: self.h1, 2: self.h2, 3: self.h3, 4: self.h4, 5: self.h5}[level]


FS = "仿宋"
FS_GB = "仿宋_GB2312"
HEI = "黑体"
KAI = "楷体"
SONG = "宋体"
XBS = "方正小标宋简体"
TNR = "Times New Roman"
MONO = "Consolas"


def official_config(opts: argparse.Namespace) -> Config:
    """党政机关公文格式 —— GB/T 9704-2012"""
    fs = FS_GB if getattr(opts, "prefer_fangsong_gb2312", False) else FS
    line_pt = _size(opts.line_spacing) if getattr(opts, "line_spacing", None) else 28.8
    title_font = getattr(opts, "title_font", None) or XBS
    title_style_kind = getattr(opts, "official_title_style", "standard")

    if title_style_kind == "heading":
        # 将 md # 视作正文一级标题处理（黑体三号，首行缩进 2 字）
        title = PStyle(cn=HEI, en=TNR, size_pt=16, bold=False, align="justify",
                       first_line_chars=2, line_pt=line_pt)
    else:
        # 标准：文件标题用 2 号小标宋体字，居中排布
        title = PStyle(cn=title_font, en=TNR, size_pt=22, bold=False,
                       align="center", first_line_chars=0,
                       line_multiple=1.0, after_pt=12)

    cfg = Config(
        mode="official",
        page=PageSpec(21.0, 29.7, 3.7, 3.5, 2.8, 2.6, 2.5, 2.5),
        # 正文：3 号仿宋，每面 22 行 ⇒ 行距固定值 ≈ 28.8 磅
        body=PStyle(cn=fs, en=TNR, size_pt=16, bold=False, italic=False,
                    align="justify", first_line_chars=2, line_pt=line_pt),
        title=title,
        # 一级标题：黑体 3 号
        h1=PStyle(cn=HEI, en=TNR, size_pt=16, bold=False, align="justify",
                  first_line_chars=2, line_pt=line_pt),
        # 二级标题：楷体 3 号加粗
        h2=PStyle(cn=KAI, en=TNR, size_pt=16, bold=True, align="justify",
                  first_line_chars=2, line_pt=line_pt),
        # 三级标题：仿宋 3 号加粗
        h3=PStyle(cn=fs, en=TNR, size_pt=16, bold=True, align="justify",
                  first_line_chars=2, line_pt=line_pt),
        h4=PStyle(cn=fs, en=TNR, size_pt=16, bold=True, align="justify",
                  first_line_chars=2, line_pt=line_pt),
        h5=PStyle(cn=fs, en=TNR, size_pt=16, bold=False, align="justify",
                  first_line_chars=2, line_pt=line_pt),
        # 注释/附注：3 号仿宋，左空二字
        quote=PStyle(cn=fs, en=TNR, size_pt=16, bold=False, align="justify",
                     left_chars=2, first_line_chars=0, line_pt=line_pt),
        listitem=PStyle(cn=fs, en=TNR, size_pt=16, bold=False, align="justify",
                        first_line_chars=0, line_pt=line_pt),
        # 代码/原文引用
        code=PStyle(cn=SONG, en=MONO, size_pt=12, bold=False, align="left",
                    left_chars=2, line_pt=18),
        # 表题：黑体 4 号，居中，置于表上方
        table_title=PStyle(cn=HEI, en=TNR, size_pt=14, bold=False, align="center",
                           first_line_chars=0, line_pt=22, before_pt=6, after_pt=6),
        # 图题：黑体 4 号，居中，置于图下方
        figure_title=PStyle(cn=HEI, en=TNR, size_pt=14, bold=False, align="center",
                            first_line_chars=0, line_pt=22, before_pt=6, after_pt=6),
        table_cell_cn=SONG,
        table_cell_en=TNR,
        table_cell_size=10.5,
        table_cell_align="center",
        table_header_bold=True,
        table_align="center",
        page_number=True,
        page_number_odd_even=True,
        page_number_cn=SONG,
        page_number_en=TNR,
        page_number_size=14.0,   # 页码用 4 号半角宋体
        page_number_dash=True,   # 数字左右各一条一字线
        header_enabled=False,
    )
    return cfg


def zju_config(opts: argparse.Namespace) -> Config:
    """浙江大学研究生学位论文编写规则"""
    line_opt = getattr(opts, "line_spacing", None)
    if line_opt:
        if str(line_opt).startswith(("1.", "1", "multi")):
            body_line = None
            body_multiple = _size(line_opt.replace("multi", ""))
        else:
            body_line = _size(line_opt)
            body_multiple = None
    else:
        body_line = 20.0        # 材料学院等细则：行距 20 磅
        body_multiple = None
    if getattr(opts, "zju_line_spacing", "fixed20") == "multi15" and not line_opt:
        body_line, body_multiple = None, 1.5

    header_left = getattr(opts, "header_left", None)
    if header_left is None:
        header_left = "浙江大学学位论文"

    cfg = Config(
        mode="zju",
        # A4；上、下 2.54cm，左、右 3.17cm；页眉 1.5cm、页脚 1.75cm
        page=PageSpec(21.0, 29.7, 2.54, 2.54, 3.17, 3.17, 1.5, 1.75),
        # 正文：仿宋小四号，1.5 倍/20 磅行距，首行缩进 2 字符
        body=PStyle(cn=FS, en=TNR, size_pt=12, bold=False, align="justify",
                    first_line_chars=2, line_pt=body_line,
                    line_multiple=body_multiple),
        # md # ：论文题目，仿宋小二号加粗居中
        title=PStyle(cn=FS, en=TNR, size_pt=18, bold=True, align="center",
                     first_line_chars=0, line_multiple=1.0, after_pt=12),
        # 章标题：仿宋三号加粗，居中，单倍行距，段前 24 磅、段后 18 磅
        h1=PStyle(cn=FS, en=TNR, size_pt=16, bold=True, align="center",
                  first_line_chars=0, line_multiple=1.0,
                  before_pt=24, after_pt=18),
        # 一级节标题：仿宋四号加粗，顶左，段前 24 磅、段后 6 磅
        h2=PStyle(cn=FS, en=TNR, size_pt=14, bold=True, align="left",
                  first_line_chars=0, line_multiple=1.0,
                  before_pt=24, after_pt=6),
        # 二级节标题：仿宋小四号加粗，顶左，段前 12 磅、段后 6 磅
        h3=PStyle(cn=FS, en=TNR, size_pt=12, bold=True, align="left",
                  first_line_chars=0, line_multiple=1.0,
                  before_pt=12, after_pt=6),
        # 三级节标题：仿宋小四号（不加粗），顶左，段前 12 磅、段后 6 磅
        h4=PStyle(cn=FS, en=TNR, size_pt=12, bold=False, align="left",
                  first_line_chars=0, line_multiple=1.0,
                  before_pt=12, after_pt=6),
        h5=PStyle(cn=FS, en=TNR, size_pt=12, bold=False, align="left",
                  first_line_chars=0, line_multiple=1.0,
                  before_pt=12, after_pt=6),
        quote=PStyle(cn=FS, en=TNR, size_pt=12, bold=False, align="justify",
                     left_chars=2, first_line_chars=0,
                     line_pt=body_line, line_multiple=body_multiple),
        listitem=PStyle(cn=FS, en=TNR, size_pt=12, bold=False, align="justify",
                        first_line_chars=0,
                        line_pt=body_line, line_multiple=body_multiple),
        code=PStyle(cn=SONG, en=MONO, size_pt=10.5, bold=False, align="left",
                    left_chars=2, line_pt=15),
        # 表题：置于表上方，仿宋五号，单倍行距，段前 6 磅、段后 6 磅，表序加粗
        table_title=PStyle(cn=FS, en=TNR, size_pt=10.5, bold=False, align="center",
                           first_line_chars=0, line_multiple=1.0,
                           before_pt=6, after_pt=6),
        # 图题：置于图下方，仿宋五号，段前 6 磅、段后 12 磅
        figure_title=PStyle(cn=FS, en=TNR, size_pt=10.5, bold=False, align="center",
                            first_line_chars=0, line_multiple=1.0,
                            before_pt=6, after_pt=12),
        table_cell_cn=SONG,
        table_cell_en=TNR,
        table_cell_size=10.5,
        table_cell_align="center",
        table_header_bold=True,
        table_align="center",
        page_number=True,
        page_number_odd_even=False,
        page_number_cn=SONG,
        page_number_en=TNR,
        page_number_size=9.0,     # 页码宋体小五号，页脚居中
        page_number_dash=False,
        # 页眉：小五号宋体居中；左上「浙江大学学位论文」，右上当前一级标题
        header_enabled=not getattr(opts, "no_header", False),
        header_left=header_left,
        header_right=getattr(opts, "header_right", "") or "",
        header_style_ref="Heading 1",
        header_cn=SONG,
        header_en=TNR,
        header_size=9.0,
        chapter_page_break=True,   # 每一章另起页
    )
    return cfg


# --------------------------------------------------------------------------- #
# 3. Markdown 解析
# --------------------------------------------------------------------------- #

HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
HR_RE = re.compile(r"^\s{0,3}([-*_])(?:\s*\1){2,}\s*$")
FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})(.*)$")
ULI_RE = re.compile(r"^(\s*)([-*+])(\s+)(.*)$")
OLI_RE = re.compile(r"^(\s*)(\d+)([.)])(\s+)(.*)$")
BQ_RE = re.compile(r"^\s{0,3}>\s?(.*)$")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
# 表题/图题识别：以「表 1-1」「表 A.1」「图2.3」「Table 1」等开头
CAPTION_RE = re.compile(
    r"^\s*[:：]?\s*(表|图|附表|附图|插图|表名|图名|Table|Figure|Fig\.?)"
    r"[\s　]*([0-9A-Za-z]+(?:[-.、\s][0-9A-Za-z]+)*)[\s　]*[:：、.]?"
)
# 单独的表题：只认「表/Table」开头，避免把图题误判成表题
TABLE_CAPTION_RE = re.compile(
    r"^\s*[:：]?\s*(表|附表|Table)[\s　]*[0-9A-Za-z]"
)
# 分节符指令（写成 HTML 注释，md 渲染时不可见）：
#   <!-- 分节符 -->                 下一页（默认）
#   <!-- 分节符: 连续 -->           连续
#   <!-- 分节符: 偶数页 / 奇数页 -->
#   <!-- section-break: continuous -->
SECTION_BREAK_RE = re.compile(
    r"^<!--\s*(?:分节符|section-break|sectionbreak)"
    r"(?:\s*[:：]\s*(?P<kind>[^>]*?))?\s*-->$"
)
SECTION_BREAK_KINDS = {
    "": WD_SECTION_START.NEW_PAGE,
    "下一页": WD_SECTION_START.NEW_PAGE,
    "新页": WD_SECTION_START.NEW_PAGE,
    "next-page": WD_SECTION_START.NEW_PAGE,
    "nextpage": WD_SECTION_START.NEW_PAGE,
    "new-page": WD_SECTION_START.NEW_PAGE,
    "page": WD_SECTION_START.NEW_PAGE,
    "连续": WD_SECTION_START.CONTINUOUS,
    "continuous": WD_SECTION_START.CONTINUOUS,
    "偶数页": WD_SECTION_START.EVEN_PAGE,
    "偶数": WD_SECTION_START.EVEN_PAGE,
    "even-page": WD_SECTION_START.EVEN_PAGE,
    "even": WD_SECTION_START.EVEN_PAGE,
    "奇数页": WD_SECTION_START.ODD_PAGE,
    "奇数": WD_SECTION_START.ODD_PAGE,
    "odd-page": WD_SECTION_START.ODD_PAGE,
    "odd": WD_SECTION_START.ODD_PAGE,
}


def section_start_of(kind: Optional[str]) -> WD_SECTION_START:
    """分节符类型名 -> WD_SECTION_START，未识别的一律按「下一页」。"""
    if not kind:
        return WD_SECTION_START.NEW_PAGE
    key = kind.strip().lower().replace("_", "-").replace(" ", "")
    return SECTION_BREAK_KINDS.get(key, WD_SECTION_START.NEW_PAGE)

INLINE_RE = re.compile(
    r"""
    (?P<code>`+[^`]*?`+)
  | (?P<image>!\[(?P<ialt>[^\]]*)\]\(\s*(?P<isrc>[^)\s]+)(?:\s+
        (?:"(?P<ititle1>[^"]*)"|'(?P<ititle2>[^']*)'))?\s*\))
  | (?P<link>\[(?P<ltext>[^\]]*)\]\(\s*(?P<lsrc>[^)\s]+)(?:\s+
        (?:"(?P<ltitle1>[^"]*)"|'(?P<ltitle2>[^']*)'))?\s*\))
  | (?P<htmlcomment><!--.*?-->)
  | (?P<autolink><(?P<asrc>[a-zA-Z][a-zA-Z0-9+.-]*://[^>\s]+)>)
  | (?P<strong>\*\*(?P<b1>.+?)\*\*|__(?P<b2>.+?)__)
  | (?P<em>\*(?P<i1>[^*\n]+?)\*|_(?P<i2>[^_\n]+?)_)
  | (?P<del>~~(?P<d1>.+?)~~)
  | (?P<prefix_esc>\\(?P<esc>[\\`*_{}\[\]()#+\-.!~>|]))
  | (?P<brtag><br\s*/?>)
  | (?P<htmltag></?[a-zA-Z][a-zA-Z0-9-]*(?:\s[^<>]*)?/?>)
  | (?P<hardbreak>  \n)
    """,
    re.VERBOSE | re.DOTALL,
)


@dataclass
class InlineToken:
    text: str = ""
    bold: bool = False
    italic: bool = False
    code: bool = False
    strike: bool = False
    url: Optional[str] = None
    image: Optional[str] = None
    break_line: bool = False


def parse_inline(text: str) -> List[InlineToken]:
    """把 markdown 行内语法解析成 token 序列。"""
    tokens: List[InlineToken] = []
    pos = 0
    buffer = ""

    def flush():
        nonlocal buffer
        if buffer:
            tokens.append(InlineToken(text=buffer))
            buffer = ""

    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            buffer += text[pos:m.start()]
        pos = m.end()

        if m.group("prefix_esc"):
            buffer += m.group("esc")
            continue
        if m.group("code"):
            flush()
            raw = m.group("code")
            n = len(raw) - len(raw.lstrip("`"))
            inner = raw[n:-n].strip() if raw[-n:] == "`" * n else raw[n:]
            tokens.append(InlineToken(text=inner, code=True))
            continue
        if m.group("image"):
            flush()
            tokens.append(InlineToken(image=m.group("isrc"),
                                      text=m.group("ialt") or ""))
            continue
        if m.group("link"):
            flush()
            start = len(tokens)
            if m.group("ltext"):
                tokens.extend(parse_inline(m.group("ltext")))
            src = m.group("lsrc")
            if start == len(tokens):
                tokens.append(InlineToken(text=src, url=src))
            else:
                for tk in tokens[start:]:
                    tk.url = src
            continue
        if m.group("autolink"):
            flush()
            tokens.append(InlineToken(text=m.group("asrc"), url=m.group("asrc")))
            continue
        if m.group("strong"):
            flush()
            inner = m.group("b1") or m.group("b2") or ""
            for tk in parse_inline(inner):
                tk.bold = True
                tokens.append(tk)
            continue
        if m.group("em"):
            flush()
            inner = m.group("i1") or m.group("i2") or ""
            for tk in parse_inline(inner):
                tk.italic = True
                tokens.append(tk)
            continue
        if m.group("del"):
            flush()
            inner = m.group("d1") or ""
            for tk in parse_inline(inner):
                tk.strike = True
                tokens.append(tk)
            continue
        if m.group("htmlcomment"):
            continue
        if m.group("brtag") or m.group("hardbreak"):
            flush()
            tokens.append(InlineToken(break_line=True))
            continue
        if m.group("htmltag"):
            continue
        # 未识别一律当作文本
        buffer += m.group(0)

    buffer += text[pos:]
    flush()
    return tokens


def strip_inline_text(text: str) -> str:
    return "".join(tk.text for tk in parse_inline(text))


@dataclass
class Block:
    kind: str                      # heading/para/list/table/code/quote/hr/image
    lines: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    align: List[str] = field(default_factory=list)
    level: int = 1
    lang: str = ""
    items: List[dict] = field(default_factory=list)
    meta: Dict[str, Any] = field(default_factory=dict)


def _is_table_start(lines: Sequence[str], i: int) -> Optional[int]:
    """若 lines[i] 是表格分隔行且 lines[i-1] 含 |，返回表头行下标。"""
    if i <= 0 or i >= len(lines):
        return None
    if not lines[i].strip().startswith(("|", "-", ":",)):
        return None
    if not TABLE_SEP_RE.match(lines[i]):
        return None
    header = lines[i - 1]
    if "|" not in header and lines[i].count("|") >= 2:
        return None
    if "|" not in header:
        return None
    return i - 1


def _split_row(line: str, expected: Optional[int] = None) -> Tuple[List[str], List[str]]:
    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|") and not text.endswith("\\|"):
        text = text[:-1]
    cells: List[str] = []
    aligns: List[str] = []
    buf = ""
    esc = False
    for ch in text:
        if esc:
            buf += ch
            esc = False
        elif ch == "\\":
            buf += ch
            esc = True
        elif ch == "|":
            cells.append(buf.strip())
            buf = ""
        else:
            buf += ch
    if esc:
        buf = buf[:-1]
    cells.append(buf.strip())
    cells = [c.replace("\\|", "|") for c in cells]
    return cells, aligns


def _col_aligns(sep_line: str, n: int) -> List[str]:
    text = sep_line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|"):
        text = text[:-1]
    out = []
    for cell in text.split("|"):
        cell = cell.strip()
        if cell.startswith(":") and cell.endswith(":"):
            out.append("center")
        elif cell.endswith(":"):
            out.append("right")
        elif cell.startswith(":"):
            out.append("left")
        else:
            out.append("")
    while len(out) < n:
        out.append("")
    return out[:n]


def parse_markdown(md_text: str) -> List[Block]:
    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    # 跳过 YAML front matter
    if lines and lines[0].strip() == "---":
        for j in range(1, len(lines)):
            if lines[j].strip() in ("---", "..."):
                lines = lines[j + 1:]
                break

    blocks: List[Block] = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # 代码块
        fence = FENCE_RE.match(line)
        if fence:
            marker = fence.group(1)[0] * 3
            lang = fence.group(2).strip()
            body: List[str] = []
            i += 1
            while i < n and not lines[i].strip().startswith(marker):
                body.append(lines[i])
                i += 1
            i += 1
            blocks.append(Block(kind="code", lines=body, lang=lang))
            continue

        # 分节符指令 / HTML 注释（注释一律不输出）
        if stripped.startswith("<!--") and stripped.endswith("-->"):
            m = SECTION_BREAK_RE.match(stripped)
            if m:
                blocks.append(Block(kind="section_break",
                                    meta={"start": section_start_of(m.group("kind"))}))
            i += 1
            continue

        # 分隔线
        if HR_RE.match(line) and not TABLE_SEP_RE.match(line):
            blocks.append(Block(kind="hr"))
            i += 1
            continue

        # 表格（分隔符行可能落在当前 i 或 i+1）
        if _is_table_start(lines, i) is not None:
            sep_idx = i
        elif i + 1 < n and _is_table_start(lines, i + 1) is not None:
            sep_idx = i + 1
        else:
            sep_idx = None

        if sep_idx is not None:
            header_idx = sep_idx - 1
            # 表题写在表格上方时，前面那个普通段落改成表题
            if blocks and blocks[-1].kind == "para" \
                    and TABLE_CAPTION_RE.match(_plain_para_text(blocks[-1])):
                blocks[-1].kind = "table_title"
            cells, _ = _split_row(lines[header_idx])
            ncol = max(len(cells), 1)
            aligns = _col_aligns(lines[sep_idx], ncol)
            rows = [cells]
            i = sep_idx + 1
            while i < n and lines[i].strip() and "|" in lines[i]:
                row_cells, _ = _split_row(lines[i])
                while len(row_cells) < ncol:
                    row_cells.append("")
                rows.append(row_cells[:ncol])
                i += 1
            blocks.append(Block(kind="table", rows=rows, align=aligns))
            # 表题写在表格下方的情况
            if i < n and lines[i].strip() and TABLE_CAPTION_RE.match(lines[i].strip()) \
                    and "|" not in lines[i]:
                blocks.append(Block(kind="table_title", lines=[lines[i].strip()],
                                    meta={"below": True}))
                i += 1
            continue

        # 标题
        hm = HEADING_RE.match(line)
        if hm:
            blocks.append(Block(kind="heading", level=len(hm.group(1)),
                                lines=[hm.group(2).strip()]))
            i += 1
            continue

        # 引用
        if BQ_RE.match(line):
            body = []
            while i < n:
                bm = BQ_RE.match(lines[i])
                if bm:
                    body.append(bm.group(1))
                    i += 1
                elif i < n and lines[i].strip() and not _starts_block(lines[i]):
                    body.append(lines[i].strip())
                    i += 1
                else:
                    break
            blocks.append(Block(kind="quote", lines=body))
            continue

        # 列表
        if ULI_RE.match(line) or OLI_RE.match(line):
            items, i = _collect_list(lines, i)
            blocks.append(Block(kind="list", items=items))
            continue

        # 段落
        para: List[str] = []
        while i < n:
            cur = lines[i].rstrip()
            if not cur.strip() or _starts_block(cur):
                break
            # 段落里遇到表格起始行则中断，交给表格分支处理
            if _is_table_start(lines, i) is not None or (
                    i + 1 < n and _is_table_start(lines, i + 1) is not None):
                break
            para.append(cur.strip())
            i += 1
        if para:
            blocks.append(Block(kind="para", lines=para))
        else:
            i += 1

    return blocks


def _plain_para_text(blk: Block) -> str:
    return " ".join(blk.lines).strip()


def _starts_block(line: str) -> bool:
    s = line.strip()
    return bool(
        HEADING_RE.match(line)
        or FENCE_RE.match(line)
        or ULI_RE.match(line)
        or OLI_RE.match(line)
        or BQ_RE.match(line)
        or HR_RE.match(line)
    )


def _indent_level(ws: str) -> int:
    expanded = ws.replace("\t", "    ")
    return len(expanded) // 2


def _collect_list(lines: Sequence[str], i: int) -> Tuple[List[dict], int]:
    items: List[dict] = []
    n = len(lines)
    while i < n:
        cur = lines[i].rstrip()
        if not cur.strip():
            # 空行后若仍是缩进列表项则继续，否则结束
            if i + 1 < n and (ULI_RE.match(lines[i + 1]) or OLI_RE.match(lines[i + 1])):
                i += 1
                continue
            break
        m = OLI_RE.match(cur)
        if m:
            level = _indent_level(m.group(1))
            items.append({"type": "ol", "level": level, "start": int(m.group(2)),
                          "text": m.group(5).strip()})
            i += 1
            continue
        m = ULI_RE.match(cur)
        if m:
            level = _indent_level(m.group(1))
            items.append({"type": "ul", "level": level,
                          "text": m.group(4).strip()})
            i += 1
            continue
        # 续行：并入上一个 item
        if items and cur.startswith((" ", "\t")) and cur.strip():
            items[-1]["text"] += " " + cur.strip()
            i += 1
            continue
        break
    return items, i


# --------------------------------------------------------------------------- #
# 4. 文档渲染
# --------------------------------------------------------------------------- #

CHINESE_NUM = "一二三四五六七八九十"


def _cn_number(n: int) -> str:
    if n <= 10:
        return CHINESE_NUM[n - 1]
    if n < 20:
        return "十" + (CHINESE_NUM[n - 11] if n > 10 else "")
    if n < 100:
        tens, ones = divmod(n, 10)
        s = CHINESE_NUM[tens - 1] + "十"
        if ones:
            s += CHINESE_NUM[ones - 1]
        return s
    return str(n)


def apply_pstyle(p, st: PStyle) -> None:
    if st.align and st.align in ALIGN_MAP:
        p.alignment = ALIGN_MAP[st.align]
    set_para_indent(p, first_line_chars=st.first_line_chars,
                    left_chars=st.left_chars,
                    hanging_chars=st.hanging_chars or None,
                    right_chars=None)
    set_line_spacing(p, multiples=st.line_multiple, exact_pt=st.line_pt,
                     rule=st.line_rule)
    set_spacing_around(p, st.before_pt, st.after_pt)
    set_keep(p, keep_next=st.keep_next, page_break_before=st.page_break_before)
    try:
        p.style.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# 内置样式名解析（兼容中文 Word / WPS 的中文样式名）
#
# .docx 文件里内置样式的 w:name 仍是英文（heading 1 / title / caption），
# 中文 Word 界面上显示为「标题 1」「标题」「题注」；WPS 等软件保存的文档
# 也可能直接把样式名写成中文。下面按「英文名 → 样式 id → 中文名」兜底解析。
# --------------------------------------------------------------------------- #

STYLE_CN_ALIAS = {
    "Title": "标题",
    "Subtitle": "副标题",
    "Caption": "题注",
    "Header": "页眉",
    "Footer": "页脚",
    "Quote": "引用",
    "Intense Quote": "明显引用",
}


def style_candidates(name: str) -> List[str]:
    """内置样式英文名 → 该文档里可能出现的样式名列表（含中文界面名）。"""
    m = re.match(r"^Heading (\d+)$", name)
    if m:
        n = m.group(1)
        cands = [name, f"heading {n}", f"标题 {n}", f"标题{n}"]
    else:
        cands = [name, name.lower(), STYLE_CN_ALIAS.get(name, "")]
    seen, out = set(), []
    for c in cands:
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def resolve_style(doc, name: str):
    """解析样式；找不到时返回 None（不要抛 KeyError 中断转换）。"""
    for cand in style_candidates(name):
        try:
            return doc.styles[cand]
        except KeyError:
            continue
    return None


def add_styled_paragraph(doc, name: str):
    """按样式名新建段落；样式不存在时退回默认正文段落。"""
    st = resolve_style(doc, name)
    return doc.add_paragraph(style=st) if st is not None else doc.add_paragraph()


def header_style_ref_name(doc, name: str) -> str:
    """页眉 STYLEREF 引用的样式名：优先用英文内置名，缺失时用文档里的实际名。"""
    try:
        doc.styles[name]
        return name
    except KeyError:
        st = resolve_style(doc, name)
        return st.name if st is not None else name


def add_tokens_to_paragraph(p, tokens: List[InlineToken], base: PStyle,
                            doc=None, image_dir: str = "",
                            max_width_cm: float = 14.9) -> bool:
    """把行内 token 写进段落。返回是否包含图片。"""
    has_image = False
    for tk in tokens:
        if tk.break_line:
            run = p.add_run()
            run.add_break(WD_BREAK.LINE)
            continue
        if tk.image:
            has_image = True
            _add_image(p, tk.image, image_dir, max_width_cm)
            continue
        if not tk.text:
            continue
        cn = base.cn
        en = base.en
        size = base.size_pt
        if tk.code:
            en = MONO
            cn = SONG
            size = max(size - 1.5, 8.0)
        run = p.add_run(tk.text)
        set_run_font(run, cn, en, size,
                     bold=base.bold or tk.bold,
                     italic=base.italic or tk.italic,
                     strike=tk.strike)
        if tk.url:
            _apply_hyperlink(doc, p, run, tk.url)
    return has_image


def _apply_hyperlink(doc, p, run, url: str) -> None:
    """把已有 run 包装成超链接（颜色强制黑色）。"""
    if doc is None:
        return
    try:
        rpr = run._element.get_or_add_rPr()
        color = rpr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            rpr.append(color)
        color.set(qn("w:val"), "000000")
        u = rpr.find(qn("w:u"))
        if u is None:
            u = OxmlElement("w:u")
            rpr.append(u)
        u.set(qn("w:val"), "single")
        r_id = p.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        p._p.insert(list(p._p).index(run._element), hyperlink)
        hyperlink.append(run._element)
    except Exception:
        # 链接建立失败不影响正文
        pass


def _add_image(p, src: str, image_dir: str, max_width_cm: float) -> Optional[object]:
    path = src
    if not os.path.isabs(path):
        candidate = os.path.join(image_dir, src)
        if os.path.exists(candidate):
            path = candidate
    if not os.path.exists(path):
        run = p.add_run(f"[图片缺失: {src}]")
        set_run_font(run, SONG, TNR, 10.5)
        return None
    try:
        run = p.add_run()
        inline = run.add_picture(path)
        max_w = Cm(max_width_cm)
        if inline.width and inline.width > max_w:
            ratio = inline.height / inline.width
            inline.width = int(max_w)
            inline.height = int(max_w * ratio)
        return inline
    except Exception as exc:  # pragma: no cover
        run = p.add_run(f"[图片无法插入: {src} ({exc})]")
        set_run_font(run, SONG, TNR, 10.5)
        return None


def setup_section(doc, cfg: Config) -> None:
    section = doc.sections[0]
    pg = cfg.page
    section.page_width = Cm(pg.width_cm)
    section.page_height = Cm(pg.height_cm)
    section.top_margin = Cm(pg.top_cm)
    section.bottom_margin = Cm(pg.bottom_cm)
    section.left_margin = Cm(pg.left_cm)
    section.right_margin = Cm(pg.right_cm)
    section.header_distance = Cm(pg.header_cm)
    section.footer_distance = Cm(pg.footer_cm)
    # 不启用装订线、不改变版式则为默认；此处显式归零以符合通行做法
    try:
        section.gutter = Cm(0)
    except Exception:
        pass


def build_footer_xml(align: str, cfg: Config) -> str:
    """GB/T 9704：页码 4 号半角宋体，数字左右各一条一字线，单页码居右空一字、
    双页码居左空一字。"""
    size = cfg.page_number_size
    rpr = _rpr_xml(cfg.page_number_cn, cfg.page_number_en, size, False, False)
    gap = "　"      # 一个汉字的空
    body = ""
    if cfg.page_number_dash:
        prefix = f"{gap}－ " if align == "left" else "－ "
        suffix = " －" if align == "left" else f" －{gap}"
    else:
        prefix = suffix = ""
    if prefix:
        body += f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(prefix)}</w:t></w:r>'
    body += f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
    body += f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>'
    body += f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
    body += f'<w:r>{rpr}<w:t>1</w:t></w:r>'
    body += f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    if suffix:
        body += f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(suffix)}</w:t></w:r>'
    return (
        f'<w:p xmlns:w="{W_NS}">'
        f'<w:pPr><w:jc w:val="{align}"/></w:pPr>'
        f"{body}</w:p>"
    )


def build_header_xml(cfg: Config, twips_width: int) -> str:
    half = twips_width // 2
    rpr = _rpr_xml(cfg.header_cn, cfg.header_en, cfg.header_size, False, False)
    tabs = (
        '<w:tabs>'
        f'<w:tab w:val="center" w:pos="{half}"/>'
        f'<w:tab w:val="right" w:pos="{twips_width}"/>'
        "</w:tabs>"
    )
    ppr_style = (
        '<w:pStyle w:val="Header"/>'
        '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" '
        'w:color="auto"/></w:pBdr>'
        if cfg.header_rule else '<w:pStyle w:val="Header"/>'
    )
    body = ""
    if cfg.header_left:
        body += f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(cfg.header_left)}</w:t></w:r>'
    body += f'<w:r>{rpr}<w:tab/></w:r>'
    if cfg.header_style_ref:
        body += f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        body += (
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> '
            f'STYLEREF "{escape_xml(cfg.header_style_ref)}" \\* MERGEFORMAT '
            "</w:instrText></w:r>"
        )
        body += f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        body += f'<w:r>{rpr}<w:t>章节标题</w:t></w:r>'
        body += f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    else:
        body += (
            f'<w:r>{rpr}<w:t xml:space="preserve">'
            f"{escape_xml(cfg.header_right)}</w:t></w:r>"
        )
    return (
        f'<w:p xmlns:w="{W_NS}"><w:pPr>{ppr_style}{tabs}</w:pPr>{body}</w:p>'
    )


def _add_story_part(document_part, package, kind: str, name: str,
                    inner_xml: str):
    from docx.opc.part import Part

    content_type = HEADER_CT if kind == "header" else FOOTER_CT
    partname = PackURI(f"/word/{name}.xml")
    part = Part(partname, content_type, XML_DECL.encode("utf-8") +
                inner_xml.encode("utf-8"), package)
    return document_part.relate_to(part, RT.HEADER if kind == "header" else RT.FOOTER)


def install_headers_footers(doc, cfg: Config) -> None:
    section = doc.sections[0]
    sect_pr = section._sectPr
    document_part = section.part
    package = document_part.package

    # 文本可用宽度 (twips)
    usable_cm = cfg.page.width_cm - cfg.page.left_cm - cfg.page.right_cm
    twips_width = int(round(usable_cm / 2.54 * 1440))

    # ---- 页眉 ----
    for node in list(sect_pr.findall(qn("w:headerReference"))):
        sect_pr.remove(node)
    if cfg.header_enabled:
        if cfg.header_style_ref:
            cfg.header_style_ref = header_style_ref_name(doc, cfg.header_style_ref)
        inner = build_header_xml(cfg, twips_width)
        xml = (
            f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}">{inner}</w:hdr>'
        )
        r_id = _add_story_part(document_part, package, "header", "header1", xml)
        ref = OxmlElement("w:headerReference")
        ref.set(qn("w:type"), "default")
        ref.set(qn("r:id"), r_id)
        sect_pr.append(ref)

    # ---- 页脚 ----
    for node in list(sect_pr.findall(qn("w:footerReference"))):
        sect_pr.remove(node)
    if not cfg.page_number:
        return
    if cfg.page_number_odd_even:
        odd_xml = (
            f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
            f"{build_footer_xml('right', cfg)}</w:ftr>"
        )
        even_xml = (
            f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
            f"{build_footer_xml('left', cfg)}</w:ftr>"
        )
        r_id = _add_story_part(document_part, package, "footer", "footer1", odd_xml)
        ref = OxmlElement("w:footerReference")
        ref.set(qn("w:type"), "default")
        ref.set(qn("r:id"), r_id)
        sect_pr.append(ref)
        r_id2 = _add_story_part(document_part, package, "footer", "footer2", even_xml)
        ref2 = OxmlElement("w:footerReference")
        ref2.set(qn("w:type"), "even")
        ref2.set(qn("r:id"), r_id2)
        sect_pr.append(ref2)
        enable_even_and_odd(doc)
    else:
        xml = (
            f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
            f"{build_footer_xml('center', cfg)}</w:ftr>"
        )
        r_id = _add_story_part(document_part, package, "footer", "footer1", xml)
        ref = OxmlElement("w:footerReference")
        ref.set(qn("w:type"), "default")
        ref.set(qn("r:id"), r_id)
        sect_pr.append(ref)


def enable_even_and_odd(doc) -> None:
    """在 settings.xml 中写入 w:evenAndOddHeaders（按 schema 顺序插入）。"""
    settings = doc.settings.element
    node = settings.find(qn("w:evenAndOddHeaders"))
    if node is None:
        node = OxmlElement("w:evenAndOddHeaders")
        anchor_tags = [qn("w:compat"), qn("w:footnotePr"), qn("w:endnotePr"),
                       qn("w:rsids"), qn("w:mathPr")]
        placed = False
        for tag in anchor_tags:
            anchor = settings.find(tag)
            if anchor is not None:
                anchor.addprevious(node)
                placed = True
                break
        if not placed:
            settings.append(node)
    node.set(qn("w:val"), "true")


def add_section_break(doc, start: WD_SECTION_START = WD_SECTION_START.NEW_PAGE) -> None:
    """插入分节符。

    做法与 python-docx 的 add_section 一致：把 body 级 sectPr 克隆一份挂到段落的 pPr 上。
    段落级 sectPr 描述「分节符之前那一节」的属性（页面设置、页眉页脚引用全部保留，
    否则会丢失奇偶页页眉/页码），分节符之后的内容继续沿用 body 级 sectPr。
    """
    body = doc.element.body
    sentinel = body.get_or_add_sectPr()
    cloned = sentinel.clone()
    cloned.start_type = start

    p = doc.add_paragraph()
    p._p.set_sectPr(cloned)

    # 让承载分节符的段落尽量不占高度
    set_spacing_around(p, 0, 0)
    set_line_spacing(p, exact_pt=1, rule="atLeast")
    rpr = OxmlElement("w:rPr")
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "2")
        rpr.append(el)
    cloned.addprevious(rpr)   # pPr 中 rPr 必须排在 sectPr 之前


# --------------------------- 表格 --------------------------- #

def add_caption_paragraph(doc, text: str, st: PStyle) -> None:
    p = add_styled_paragraph(doc, "Caption")
    apply_pstyle(p, st)
    tokens = parse_inline(text)
    add_tokens_to_paragraph(p, tokens, st, doc=doc, image_dir="")
    # 表题/图题整体不参与自动编号时对「序号」加粗
    m = CAPTION_RE.match(text)
    if m:
        end = m.end(2) + (1 if m.group(0).rstrip() else 0)
        for run in p.runs:
            run.font.bold = True
            if len(run.text) >= end:
                tail = run.text[end:]
                run.text = run.text[:end]
                if tail:
                    run2 = p.add_run(tail)
                    set_run_font(run2, st.cn, st.en, st.size_pt, bold=False)
                break
            end -= len(run.text)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)


def add_table(doc, block: Block, cfg: Config) -> None:
    rows = block.rows
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncol)
    if cfg.table_style:
        try:
            table.style = cfg.table_style
        except KeyError:
            table.style = "Table Grid"
    table.autofit = True
    if cfg.table_align == "center":
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_align = ALIGN_MAP.get(cfg.table_cell_align, WD_ALIGN_PARAGRAPH.CENTER)

    for r_idx, row in enumerate(rows):
        for c_idx in range(ncol):
            cell = table.cell(r_idx, c_idx)
            cell_p = cell.paragraphs[0]
            text = row[c_idx] if c_idx < len(row) else ""
            # 清空单元格原有内容（不要留下空 run）
            for _r in list(cell_p.runs):
                _r._element.getparent().remove(_r._element)
            cell_p.alignment = cell_align
            set_line_spacing(cell_p, multiples=1.0)
            set_spacing_around(cell_p, 0, 0)
            # 单元格上下留白最小化
            tc_pr = cell._tc.get_or_add_tcPr()
            va = tc_pr.find(qn("w:vAlign"))
            if va is None:
                va = OxmlElement("w:vAlign")
                tc_pr.append(va)
            va.set(qn("w:val"), "center")
            mar = tc_pr.find(qn("w:tcMar"))
            if mar is None:
                mar = OxmlElement("w:tcMar")
                tc_pr.append(mar)
            for side in ("top", "bottom"):
                node = mar.find(qn(f"w:{side}"))
                if node is None:
                    node = OxmlElement(f"w:{side}")
                    mar.append(node)
                node.set(qn("w:w"), "28")
                node.set(qn("w:type"), "dxa")
            for side in ("left", "right"):
                node = mar.find(qn(f"w:{side}"))
                if node is None:
                    node = OxmlElement(f"w:{side}")
                    mar.append(node)
                node.set(qn("w:w"), "72")
                node.set(qn("w:type"), "dxa")

            is_header = (r_idx == 0)
            tokens = parse_inline(text)
            base = PStyle(cn=cfg.table_cell_cn, en=cfg.table_cell_en,
                          size_pt=cfg.table_cell_size,
                          bold=is_header and cfg.table_header_bold,
                          align="center")
            if not tokens:
                tokens = [InlineToken(text="")]
            add_tokens_to_paragraph(cell_p, tokens, base, doc=doc)
            for run in cell_p.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)

    # 表头跨页重复
    first_row = table.rows[0]._tr
    tr_pr = first_row.get_or_add_trPr()
    tbl_header = tr_pr.find(qn("w:tblHeader"))
    if tbl_header is None:
        tr_pr.append(OxmlElement("w:tblHeader"))

    # 表后补一个空行，避免与后续段落粘连
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    set_line_spacing(spacer, exact_pt=6, rule="atLeast")


# --------------------------- 主渲染 --------------------------- #

def render(doc, blocks: List[Block], cfg: Config, image_dir: str) -> None:
    usable_cm = cfg.page.width_cm - cfg.page.left_cm - cfg.page.right_cm
    max_width_cm = usable_cm

    heading_counters = [0] * 6
    seen_heading_1 = False

    idx = 0
    while idx < len(blocks):
        block = blocks[idx]

        if block.kind == "table_title":
            add_caption_paragraph(doc, _plain_para_text(block), cfg.table_title)
            idx += 1
            continue

        if block.kind == "heading":
            md_level = block.level
            text = _plain_para_text(block)

            need_page_break = (
                cfg.chapter_page_break and md_level == 2 and seen_heading_1
            )
            if need_page_break:
                # 独占一段的「段的开始处分页」，避免重复分页产生空白页
                br = doc.add_paragraph()
                set_line_spacing(br, exact_pt=6, rule="atLeast")
                br._p.get_or_add_pPr().append(OxmlElement("w:pageBreakBefore"))

            if md_level == 1:
                p = add_styled_paragraph(doc, "Title")
                style_obj = cfg.title
            else:
                docx_level = min(md_level - 1, 5)
                p = add_styled_paragraph(doc, f"Heading {docx_level}")
                style_obj = cfg.heading_style(docx_level)
                if docx_level == 1:
                    seen_heading_1 = True

            if cfg.auto_number and md_level >= 2:
                k = md_level - 2
                heading_counters[k] += 1
                for deeper in range(k + 1, 6):
                    heading_counters[deeper] = 0
                pattern = cfg.auto_number_patterns[min(k, len(cfg.auto_number_patterns) - 1)]
                prefix = pattern.replace("{}", _cn_number(heading_counters[k]))
                text = f"{prefix}{text}"

            apply_pstyle(p, style_obj)
            tokens = parse_inline(text)
            add_tokens_to_paragraph(p, tokens, style_obj, doc=doc,
                                    image_dir=image_dir, max_width_cm=max_width_cm)
            for run in p.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)
            idx += 1
            continue

        if block.kind == "table":
            add_table(doc, block, cfg)
            idx += 1
            continue

        if block.kind == "para":
            text = " ".join(block.lines)
            tokens = parse_inline(text)
            only_image = bool(tokens) and all(
                tk.image for tk in tokens) and any(tk.image for tk in tokens)
            if only_image:
                p = doc.add_paragraph()
                p.alignment = ALIGN_MAP["center"]
                set_line_spacing(p, multiples=1.0)
                set_spacing_around(p, 6, 6)
                add_tokens_to_paragraph(p, tokens, cfg.body, doc=doc,
                                        image_dir=image_dir,
                                        max_width_cm=max_width_cm)
            else:
                p = doc.add_paragraph()
                apply_pstyle(p, cfg.body)
                add_tokens_to_paragraph(p, tokens, cfg.body, doc=doc,
                                        image_dir=image_dir,
                                        max_width_cm=max_width_cm)
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)
            # 单独成段的图题/表题
            if CAPTION_RE.match(text.strip()):
                cap = resolve_style(doc, "Caption")
                if cap is not None:
                    p.style = cap
                p.paragraph_format.first_line_indent = Pt(0)
                for run in p.runs:
                    set_run_font(run, cfg.figure_title.cn, cfg.figure_title.en,
                                 cfg.figure_title.size_pt,
                                 bold=cfg.figure_title.bold)
                p.alignment = ALIGN_MAP[cfg.figure_title.align or "center"]
                set_spacing_around(p, cfg.figure_title.before_pt,
                                   cfg.figure_title.after_pt)
                set_line_spacing(p, multiples=cfg.figure_title.line_multiple,
                                 exact_pt=cfg.figure_title.line_pt)
            idx += 1
            continue

        if block.kind == "list":
            counters: Dict[int, int] = {}
            for item in block.items:
                level = max(0, int(item.get("level", 0)))
                ttype = item.get("type", "ul")
                for deeper in list(counters):
                    if deeper > level:
                        counters.pop(deeper, None)
                if level not in counters:
                    # 以该层第一个列表项的起始序号为基准
                    counters[level] = int(item.get("start") or 1) - 1
                counters[level] += 1
                if ttype == "ol":
                    seq = counters[level]
                    pattern = cfg.ordinal_patterns[
                        min(level, len(cfg.ordinal_patterns) - 1)]
                    if CIRCLED_DIGITS and pattern == CIRCLED_DIGITS:
                        marker = (
                            CIRCLED_DIGITS[seq - 1]
                            if 0 < seq <= len(CIRCLED_DIGITS) else f"({seq})"
                        )
                    else:
                        marker = pattern.replace("{}", str(seq))
                else:
                    marker = cfg.bullet_markers[
                        min(level, len(cfg.bullet_markers) - 1)]
                p = doc.add_paragraph()
                st = cfg.listitem
                p.alignment = ALIGN_MAP.get(st.align or "justify",
                                            WD_ALIGN_PARAGRAPH.JUSTIFY)
                set_line_spacing(p, multiples=st.line_multiple, exact_pt=st.line_pt,
                                 rule=st.line_rule)
                set_spacing_around(p, st.before_pt, st.after_pt)
                set_para_indent(p,
                                first_line_chars=None,
                                hanging_chars=2.0,
                                left_chars=2.0 + 2.0 * level)
                r0 = p.add_run(marker + " ")
                set_run_font(r0, st.cn, st.en, st.size_pt, bold=st.bold,
                             italic=st.italic)
                add_tokens_to_paragraph(p, parse_inline(item.get("text", "")), st,
                                        doc=doc, image_dir=image_dir,
                                        max_width_cm=max_width_cm)
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)
            spacer = doc.add_paragraph()
            set_line_spacing(spacer, exact_pt=6, rule="atLeast")
            idx += 1
            continue

        if block.kind == "quote":
            text = "\n".join(block.lines).strip()
            p = doc.add_paragraph()
            apply_pstyle(p, cfg.quote)
            add_tokens_to_paragraph(p, parse_inline(text), cfg.quote, doc=doc,
                                    image_dir=image_dir,
                                    max_width_cm=max_width_cm)
            for run in p.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)
            idx += 1
            continue

        if block.kind == "code":
            st = cfg.code
            for line_txt in block.lines:
                p = doc.add_paragraph()
                p.alignment = ALIGN_MAP.get(st.align or "left",
                                            WD_ALIGN_PARAGRAPH.LEFT)
                set_line_spacing(p, multiples=st.line_multiple, exact_pt=st.line_pt,
                                 rule=st.line_rule)
                set_spacing_around(p, st.before_pt, st.after_pt)
                set_para_indent(p, left_chars=st.left_chars, first_line_chars=0)
                if not line_txt and len(block.lines) == 1:
                    line_txt = " "
                run = p.add_run(line_txt if line_txt else " ")
                set_run_font(run, st.cn, st.en, st.size_pt, bold=st.bold)
                if hasattr(p, "_p"):
                    ppr = p._p.get_or_add_pPr()
                    shd = ppr.find(qn("w:shd"))
                    if shd is None:
                        shd = OxmlElement("w:shd")
                        ppr.append(shd)
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:fill"), "F2F2F2")
            spacer = doc.add_paragraph()
            set_line_spacing(spacer, exact_pt=6, rule="atLeast")
            idx += 1
            continue

        if block.kind == "section_break":
            add_section_break(doc, block.meta.get("start", WD_SECTION_START.NEW_PAGE))
            idx += 1
            continue

        if block.kind == "hr":
            p = doc.add_paragraph()
            ppr = p._p.get_or_add_pPr()
            pbdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "auto")
            pbdr.append(bottom)
            ppr.append(pbdr)
            set_line_spacing(p, exact_pt=6, rule="atLeast")
            idx += 1
            continue

        idx += 1
    return

def cleanup_empty_trailing_paragraphs(doc) -> None:
    """删除文档末尾多余的空段落。"""
    body = doc.element.body
    children = list(body)
    while children:
        last = children[-1]
        if last.tag == qn("w:p"):
            text = "".join(n.text or "" for n in last.iter(qn("w:t")))
            has_drawing = last.find(f".//{{{W_NS}}}drawing") is not None
            has_sectpr = last.find(f".//{{{W_NS}}}sectPr") is not None
            if not text.strip() and not has_drawing and not has_sectpr:
                body.remove(last)
                children.pop()
                continue
        break


# --------------------------------------------------------------------------- #
# 5. 字体可用性检查
# --------------------------------------------------------------------------- #

# 中文常用字体 -> 注册表中的英文面名 / 字体文件名（任一命中即视为已安装）
CN_FONT_ALIASES: Dict[str, set] = {
    "宋体": {"宋体", "SimSun", "NSimSun", "simsun.ttc", "simsun.ttf"},
    "黑体": {"黑体", "SimHei", "simhei.ttf"},
    "楷体": {"楷体", "楷体_GB2312", "KaiTi", "KaiTi_GB2312", "simkai.ttf"},
    "仿宋": {"仿宋", "FangSong", "simfang.ttf"},
    "仿宋_GB2312": {"仿宋_GB2312", "FangSong_GB2312", "FangSong_GB2312.ttf",
                 "仿宋_GB2312.ttf"},
    "方正小标宋简体": {"方正小标宋简体", "FZXiaoBiaoSong-B05S", "FZXiaoBiaoSong-B05",
                  "FZXBSJW.TTF", "Founder Xiaobiaosong", "FZXiaoBiaoSong"},
    "华文中宋": {"华文中宋", "STZhongsong", "STZHONGS.TTF"},
    "微软雅黑": {"微软雅黑", "Microsoft YaHei", "msyh.ttc"},
    "Times New Roman": {"Times New Roman", "times.ttf"},
    "Consolas": {"Consolas", "consola.ttf"},
    "Arial": {"Arial", "arial.ttf"},
}


def installed_fonts() -> Optional[set]:
    """返回本机可用字体的候选集合（注册表条目名 + 字体文件名）。"""
    names: set = set()
    if sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
            )
            idx = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, idx)
                except OSError:
                    break
                if name:
                    names.add(re.sub(r"\s*\([^()]*\)\s*$", "", name).strip())
                if value:
                    names.add(os.path.basename(str(value)).strip())
                idx += 1
            winreg.CloseKey(key)
        except Exception:
            pass
        font_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        try:
            for fn in os.listdir(font_dir):
                names.add(fn)
        except Exception:
            pass
    if not names:
        return None
    return names


def font_available(font_name: str, names: set) -> bool:
    candidates = CN_FONT_ALIASES.get(font_name, {font_name})
    lowered = {n.lower() for n in names}
    return any(c.lower() in lowered for c in candidates if c)


def check_fonts(cfg: Config, extra: Optional[Sequence[str]] = None) -> List[str]:
    fonts = installed_fonts()
    if fonts is None:
        return []
    needed = {
        cfg.body.cn, cfg.title.cn, cfg.h1.cn, cfg.h2.cn, cfg.h3.cn,
        cfg.table_cell_cn, cfg.page_number_cn, cfg.header_cn,
    }
    if extra:
        needed.update(extra)
    needed.discard("")
    return sorted(f for f in needed if f and not font_available(f, fonts))


# --------------------------------------------------------------------------- #
# 6. CLI
# --------------------------------------------------------------------------- #

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Markdown 转 docx（党政公文格式 / 浙江大学学位论文格式）",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", help="输入 .md 文件")
    parser.add_argument("-o", "--output", help="输出 .docx 文件（默认与输入同目录同名）")
    parser.add_argument("-m", "--mode", choices=["official", "zju"],
                        default="official", help="版式：official=党政公文；zju=浙江大学学位论文")

    parser.add_argument("--line-spacing", help="正文行距，可填磅值(28.8)或倍数(1.5)")
    parser.add_argument("--zju-line-spacing", choices=["fixed20", "multi15"],
                        default="fixed20", help="浙大模式行距：固定20磅 或 1.5倍")
    parser.add_argument("--title-font", help="公文标题字体（默认方正小标宋简体）")
    parser.add_argument("--official-title-style", choices=["standard", "heading"],
                        default="standard",
                        help="公文标题：standard=二号小标宋居中；heading=按正文一级标题处理")
    parser.add_argument("--prefer-fangsong-gb2312", action="store_true",
                        help="正文仿宋改用 仿宋_GB2312")

    parser.add_argument("--table-cell-font", help="表格正文字体（默认宋体）")
    parser.add_argument("--table-cell-size", help="表格正文字号（默认五号=10.5）")
    parser.add_argument("--no-table-header-bold", action="store_true",
                        help="表头行不加粗")

    parser.add_argument("--no-page-number", action="store_true", help="不生成页码")
    parser.add_argument("--no-odd-even", action="store_true",
                        help="公文页码不区分奇偶页（统一居中）")

    parser.add_argument("--no-header", action="store_true", help="不生成页眉")
    parser.add_argument("--header-left", help="页眉左側文字")
    parser.add_argument("--header-right", help="页眉右側文字（不使用 STYLEREF 时生效）")

    parser.add_argument("--no-chapter-page-break", action="store_true",
                        help="浙大模式下每一章不另起页")
    parser.add_argument("--bullet-style", choices=list(BULLET_STYLES.keys()),
                        default="dash", help="无序列表项目符号")
    parser.add_argument("--ordinal-style", choices=list(ORDINAL_STYLES.keys()),
                        default="arabic", help="有序列表编号样式")
    parser.add_argument("--auto-number", action="store_true",
                        help="为 md 二级及以下标题自动生成一、/（一）/1. 序号")

    parser.add_argument("--check-fonts", action="store_true",
                        help="检查所需中文字体是否已安装（默认已开启）")
    parser.add_argument("--no-font-check", action="store_true",
                        help="关闭字体检测提示")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    src = os.path.abspath(args.input)
    if not os.path.isfile(src):
        sys.stderr.write(f"找不到输入文件: {src}\n")
        return 1
    out = args.output
    if not out:
        base = os.path.splitext(src)[0]
        out = f"{base}_{args.mode}.docx"
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)

    with open(src, "r", encoding="utf-8") as fh:
        md_text = fh.read()

    cfg = official_config(args) if args.mode == "official" else zju_config(args)

    if args.table_cell_font:
        cfg.table_cell_cn = args.table_cell_font
    if args.table_cell_size:
        cfg.table_cell_size = _size(args.table_cell_size)
    if args.no_table_header_bold:
        cfg.table_header_bold = False
    if args.no_page_number:
        cfg.page_number = False
    if args.no_odd_even:
        cfg.page_number_odd_even = False
    if args.no_header:
        cfg.header_enabled = False
    if args.no_chapter_page_break:
        cfg.chapter_page_break = False
    cfg.bullet_markers = list(BULLET_STYLES[args.bullet_style])
    cfg.ordinal_patterns = list(ORDINAL_STYLES[args.ordinal_style])
    cfg.auto_number = bool(args.auto_number)

    missing = [] if args.no_font_check else check_fonts(cfg)
    if missing:
        sys.stderr.write("提示：以下字体在本机未安装，Word/WPS 会自动替换，"
                         "建议安装或改用已装字体："
                         f"{'、'.join(missing)}\n")

    doc = Document()
    setup_section(doc, cfg)
    install_headers_footers(doc, cfg)

    try:
        blocks = parse_markdown(md_text)
        render(doc, blocks, cfg, image_dir=os.path.dirname(src))
        cleanup_empty_trailing_paragraphs(doc)
        doc.save(out)
    except Exception as exc:
        sys.stderr.write(f"生成失败: {exc}\n")
        raise

    sys.stdout.write(f"已生成: {out}  (模式: {args.mode})\n")
    if missing:
        sys.stdout.write("注意缺失字体：" + "、".join(missing) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
