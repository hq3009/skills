# -*- coding: utf-8 -*-
"""
由定稿的评审意见 Markdown 生成"评审意见概要" word 文档（填表/简版）。

用法：
    python scripts/fill_docx_summary.py --md <定稿.md 或 项目目录> [--template <底本.docx>] [--out <输出.docx>]

--md 传目录时，自动选用该目录下日期最新的 评审意见_YYYY-MM-DD.md（没有时回退到 评审意见.md）。
输出默认命名为 评审意见概要_<当天日期>.docx（与 md 同目录）；同一天再次生成会覆盖当天文件，
如需保留可用 --out 另行命名。

生成的文档内容：
- 标题段：<实施单位>-<项目名称>-评审意见（套用模板 heading 1 样式）
- 正文：md "## 评审意见概要" 的全部内容——总体意见段（不编号）与编号条目，一条一段，条目保留 1. 2. 3. 编号

样式约束：只写文本，**不新建、不修改任何样式定义，不改页面设置**，
标题与正文段落均复制模板原有空段落的段落属性，格式完全由模板决定。

本脚本仅依赖 Python 标准库。
"""
import argparse
import copy
import os
import re
import sys
import zipfile
from datetime import date
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fill_docx_review import (  # noqa: E402
    W,
    XML_SPACE,
    MISSING_VALUES,
    build_docx,
    clean,
    eprint,
    parse_md,
    pick_latest_md,
)

DEFAULT_TEMPLATE = "评审意见概要模板.docx"
# 底本模板 styles.xml 中的标题样式名（找不到时回退到 styleId "3"）
HEADING_STYLE_NAME = "heading 1"
HEADING_STYLE_FALLBACK = "3"

TITLE_FIELD_UNIT = "实施单位"
TITLE_FIELD_NAME = "项目名称"
TITLE_SUFFIX = "评审意见"


def find_style_id(tpl_path, style_name):
    """在底本 styles.xml 中按样式名查找 styleId，找不到返回 None。"""
    target = clean(style_name)
    try:
        with zipfile.ZipFile(tpl_path) as z:
            styles = z.read("word/styles.xml").decode("utf-8")
    except (KeyError, zipfile.BadZipFile):
        return None
    for m in re.finditer(r"<w:style\b([^>]*)>(.*?)</w:style>", styles, re.S):
        attrs, body = m.group(1), m.group(2)
        if "w:styleId=" not in attrs:
            continue
        nm = re.search(r'w:name w:val="([^"]+)"', body)
        if nm and clean(nm.group(1)) == target:
            sid = re.search(r'w:styleId="([^"]+)"', attrs)
            if sid:
                return sid.group(1)
    return None


def clear_runs(p):
    """清空段落中的全部内容子元素（保留 pPr）。"""
    for child in list(p):
        if child.tag != W + "pPr":
            p.remove(child)


def set_pstyle(p, val):
    """给段落套用指定段落样式；pStyle 置于 pPr 首位以符合 schema 顺序。"""
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        p.insert(0, ppr)
    st = ppr.find(W + "pStyle")
    if st is None:
        st = ET.Element(W + "pStyle")
        ppr.insert(0, st)
    st.set(W + "val", val)


def add_run(p, text):
    r = ET.SubElement(p, W + "r")
    t = ET.SubElement(r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = text


def build_paragraphs(body, title, items, heading_style):
    """沿用模板首个空段落的格式生成标题段与条目段，并替换该空段落。

    返回写入的条目数。段落属性（间距、缩进、对齐等）全部来自模板，不做任何改写。
    """
    tpl = body.find(W + "p")
    if tpl is None:
        raise RuntimeError("模板中未找到可沿用的段落，无法生成文档")

    new = []
    p_title = copy.deepcopy(tpl)
    clear_runs(p_title)
    set_pstyle(p_title, heading_style)
    add_run(p_title, title)
    new.append(p_title)

    for it in items:
        p = copy.deepcopy(tpl)
        clear_runs(p)
        add_run(p, it)
        new.append(p)

    idx = list(body).index(tpl)
    body.remove(tpl)
    for i, p in enumerate(new):
        body.insert(idx + i, p)
    return len(items)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    default_template = os.path.join(here, "..", "assets", DEFAULT_TEMPLATE)

    ap = argparse.ArgumentParser(description="由定稿评审意见 md 生成评审意见概要 word 文档")
    ap.add_argument("--md", required=True, help="评审意见定稿 Markdown 文件路径；也可传目录，自动选用其中日期最新的评审意见")
    ap.add_argument("--out", default=None, help="输出 docx 路径（默认：与 md 同目录的 评审意见概要_<当天日期>.docx）")
    ap.add_argument("--template", default=default_template, help="底本 docx：默认使用 assets/%s（只读，不修改）" % DEFAULT_TEMPLATE)
    args = ap.parse_args()

    md_path = os.path.abspath(args.md)
    tpl_path = os.path.abspath(args.template)

    # --md 传目录时，自动选用日期最新的评审意见
    if os.path.isdir(md_path):
        picked = pick_latest_md(md_path)
        if not picked:
            eprint("目录中未找到评审意见 Markdown（评审意见_YYYY-MM-DD.md 或 评审意见.md）：%s" % md_path)
            sys.exit(1)
        print("自动选用日期最新的定稿：%s" % picked)
        md_path = picked

    # 输出默认命名：与 md 同目录的 评审意见概要_<当天日期>.docx
    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        out_path = os.path.join(os.path.dirname(md_path), "评审意见概要_%s.docx" % date.today().isoformat())
    if os.path.exists(out_path):
        print("注意：同名文件已存在，将覆盖：%s" % out_path)

    if not os.path.isfile(md_path):
        eprint("找不到 Markdown 文件：%s" % md_path)
        sys.exit(1)
    if not os.path.isfile(tpl_path):
        eprint("找不到模板文件：%s" % tpl_path)
        sys.exit(1)
    if os.path.abspath(tpl_path) == out_path:
        eprint("输出路径不能与模板路径相同（模板只读，不得覆盖）。")
        sys.exit(1)

    with open(md_path, encoding="utf-8") as f:
        md_text = f.read()

    summary_items, _body, fields = parse_md(md_text)
    if not summary_items:
        eprint("md 中未解析到“评审意见概要”条目，未生成文档。")
        sys.exit(1)

    notes = []

    def field(name):
        v = (fields.get(name) or "").strip()
        if v in MISSING_VALUES:
            notes.append("%s：md 未提供，标题中记为“未提供”" % name)
            return "未提供"
        return v

    title = "%s-%s-%s" % (field(TITLE_FIELD_UNIT), field(TITLE_FIELD_NAME), TITLE_SUFFIX)
    heading_style = find_style_id(tpl_path, HEADING_STYLE_NAME) or HEADING_STYLE_FALLBACK

    with zipfile.ZipFile(tpl_path) as zin:
        xml_path = "word/document.xml"
        document_xml = zin.read(xml_path).decode("utf-8")
        root = build_docx(document_xml)
        body = root.find(W + "body")
        n_items = build_paragraphs(body, title, summary_items, heading_style)

        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == xml_path:
                    data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
                    # 规范为 OOXML 要求的声明形式（双引号 + standalone）
                    data = data.replace(
                        b"<?xml version='1.0' encoding='UTF-8'?>",
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    )
                zout.writestr(item, data)

    print("已生成：%s" % out_path)
    print("  标题：%s" % title)
    print("  概要条目：%d 条" % n_items)
    if notes:
        print("  以下项目需手工确认：")
        for n in notes:
            print("    - %s" % n)


if __name__ == "__main__":
    main()
