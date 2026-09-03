# -*- coding: utf-8 -*-
"""
将评审意见最终稿（Markdown）的"评审意见概要"与"三问三看三查三核"
回填到官方 word 评审表（assets/三问三看模板.docx 的结构）中。

用法：
    python scripts/fill_docx_review.py --md <最终稿.md 或 报告工作目录> [--template <底本.docx>] [--out <输出.docx>]

--md 传目录时，自动选用该目录下版本号最大的 评审意见_NNN.md（没有版本文件时回退到 评审意见.md）。
输出默认命名为 评审意见表_<当天日期>.docx（与 md 同目录），表示"更新日期"；同一天再次生成会覆盖当天文件，
如需保留可用 --out 另行命名。Markdown 侧的版本流水号（_001/_002）不用于 word 文件名。

说明：
- 底本默认取本 skill 的 assets/三问三看模板.docx（只读，不修改）。若要在上一版 word 基础上更新，
  用 --template 指定旧版文件（如 评审意见表_2026-09-01.docx，改名前的 评审意见_2026-09-01.docx 亦可）：此时结论勾选会先复位再按新 md 勾选，
  md 中未提供的字段则沿用旧版原值。
- "评审意见"大栏 <- "## 评审意见概要"的全部条目（保留 1. 2. 3. 编号，每条一段）。
- "评审要点"表"评审意见"列 <- 三问三看三查三核 12 个子项的判断文字；尊重模板原有合并单元格：
  三问(3 行)、三核(3 行)整组一格；三看/三查前两行一格、第三行单独一格。各子项文字按行顺序分段写入。
- 项目基本信息（项目名称/项目编码/预算部门/实施单位/资金总需求/事前评估周期）、
  "评审结论"勾选框、"核定入库金额"：md 中信息可得时自动填入；不可得（未填、填"未提供/待核实"、
  或结论未唯一确定）则保持原样，并在终端列出待手工填写的项目。
- 只向空单元格写入文本，不改变表格结构；段落沿用单元格内空段的字体、字号、缩进与对齐设置。
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

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# word 模板"评审要点"表中 12 个二级指标（顺序即表格行顺序，与 md 的四级标题一致）
SUBTITLES = [
    "政府和市场的边界是否清晰",
    "是否属于市级财政事权和支出责任",
    "是否属于本部门职能",
    "立项依据是否充分",
    "是否有明显的预期效益",
    "预算是否经济节约",
    "与其他政策项目是否重叠交叉",
    "是否存在实施风险",
    "资金来源是否符合相关规定",
    "绩效目标和政策项目初衷是否匹配",
    "绩效指标是否全面细化",
    "指标值和预算是否匹配",
]

SECTION_SUMMARY = "评审意见概要"
SECTION_TABLE = "三问三看三查三核"
SECTION_BASIC = "项目基本信息"

# 基本信息区可自动填写的字段（字段名 -> 值写入该字段标签后的空单元格）
BASIC_FIELDS = [
    "项目名称",
    "项目编码",
    "预算部门",
    "实施单位",
    "资金总需求",
    "事前评估周期",
]
CONCLUSION_FIELD = "评审结论"
AMOUNT_FIELD = "核定入库金额"
# 结论选项：长选项先匹配，避免"予以入库"被"不予入库"误命中
CONCLUSION_OPTIONS = ["不予入库", "调整后再审", "予以入库"]
# Wingdings 2：00A3 = ⬜ 空方框（模板原符号），0052 = ☑ 带勾方框
CHECK_BOX_EMPTY = "00A3"
CHECK_BOX_CHECKED = "0052"
# 视为"信息不可得"的取值
MISSING_VALUES = {"", "未提供", "待核实", "未提供/待核实", "无", "/", "-", "—", "－"}


def eprint(*args):
    print(*args, file=sys.stderr)


def register_xmlns(xml_text):
    """收集 document.xml 中声明的全部 xmlns 前缀，序列化时保持前缀不变。"""
    for m in re.finditer(r'\sxmlns:([A-Za-z_][\w.-]*)="([^"]+)"', xml_text):
        ET.register_namespace(m.group(1), m.group(2))
    # 无前缀的默认命名空间由 ElementTree 在根元素序列化时自行处理


def clean(text):
    return re.sub(r"\s+", "", text or "")


# ---------------- Markdown 解析 ----------------

def _join_lines(lines):
    """把若干非空行按 markdown 语义合并为一段（软换行不加空格，中文语境更自然）。"""
    out = []
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        # 去掉引用块前缀、无序列表符号
        s = re.sub(r"^>\s?", "", s)
        s = re.sub(r"^[-*]\s+", "", s)
        out.append(s)
    return "".join(out)


def parse_md(md_text):
    """返回 (summary_items, body_by_title, fields)。

    summary_items: 评审意见概要的条目文本列表（保留原编号前缀，如 "1. …"）。
    body_by_title: {二级指标标题: 该子项判断文字}，文字保留组内编号（如 "① …"）。
    fields: {项目基本信息字段名: 取值}，含评审结论与核定入库金额。
    """
    lines = md_text.splitlines()

    # 找到各小节位置
    def find_section(prefix):
        for i, ln in enumerate(lines):
            m = re.match(r"^#{2,3}\s+(.*?)\s*$", ln.strip())
            if m and clean(m.group(1)) == clean(prefix):
                return i
        return -1

    sum_start = find_section(SECTION_SUMMARY)
    table_start = find_section(SECTION_TABLE)

    # ---- 概要条目 ----
    summary_items = []
    if sum_start >= 0:
        end = table_start if table_start > sum_start else len(lines)
        cur = None
        for ln in lines[sum_start + 1:end]:
            if not ln.strip():
                continue
            m = re.match(r"^\s*(\d+)\s*[.．、]\s*(.*)$", ln)
            if m:
                cur = m.group(2).strip()
                summary_items.append("%s. %s" % (m.group(1), cur))
            elif cur is not None:
                cur += re.sub(r"^>\s?", "", ln.strip())

    # ---- 三问三看三查三核 12 个子项 ----
    body_by_title = {}
    if table_start >= 0:
        cur_title = None
        buf = []
        for ln in lines[table_start + 1:]:
            m = re.match(r"^#{2,5}\s+(.*?)\s*$", ln.strip())
            if m:
                if cur_title is not None:
                    body_by_title[cur_title] = _join_lines(buf)
                cur_title = clean(m.group(1))
                buf = []
            elif cur_title is not None:
                buf.append(ln)
        if cur_title is not None:
            body_by_title[cur_title] = _join_lines(buf)

    # ---- 项目基本信息（含评审结论、核定入库金额）----
    fields = {}
    basic_start = -1
    for i, ln in enumerate(lines):
        m = re.match(r"^#{2,3}\s+(.*?)\s*$", ln.strip())
        if m and clean(m.group(1)) == clean(SECTION_BASIC):
            basic_start = i
            break
    if basic_start >= 0:
        known = set(BASIC_FIELDS) | {CONCLUSION_FIELD, AMOUNT_FIELD}
        for ln in lines[basic_start + 1:]:
            if ln.lstrip().startswith("#"):
                break
            m = re.match(r"^\s*[-*>]?\s*(\S+)\s*[：:]\s*(.*)$", ln)
            if m and clean(m.group(1)) in known:
                fields[clean(m.group(1))] = m.group(2).strip()

    return summary_items, body_by_title, fields


# ---------------- docx XML 操作 ----------------

def get_cell_text(tc):
    return "".join(t.text or "" for t in tc.iter(W + "t"))


def get_vmerge(tc):
    tcPr = tc.find(W + "tcPr")
    if tcPr is None:
        return None
    vm = tcPr.find(W + "vMerge")
    if vm is None:
        return None
    return vm.get(W + "val") or "continue"


def _remove_local_attrs(el, local):
    """按本地名删除属性（如 w14:paraId）。"""
    for k in list(el.attrib):
        local_name = k.split("}")[-1]
        if local_name == local:
            del el.attrib[k]


def insert_text_into_p(p, text):
    """在段内写入文本：复制段落标记(rPr)里的字体/字号设置到新 run，保持原格式。"""
    # 清除段内已有 run（模板空段本无 run）
    for r in list(p):
        if r.tag == W + "r":
            p.remove(r)
    r = ET.Element(W + "r")
    pPr = p.find(W + "pPr")
    if pPr is not None:
        rpr = pPr.find(W + "rPr")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    t = ET.SubElement(r, W + "t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    p.append(r)


def fill_tc(tc, paras):
    """把 paras（段落文本列表）按顺序写入单元格 tc。

    第一个段落复用单元格内已有的段落（保证 pPr/run 格式与模板一致），
    后续段落克隆首段的 pPr 生成，不改变表格结构。
    """
    if not paras:
        return
    ps = [child for child in tc if child.tag == W + "p"]
    if not ps:
        return
    first = ps[0]
    for extra in ps[1:]:
        tc.remove(extra)
    insert_text_into_p(first, paras[0])
    for text in paras[1:]:
        newp = copy.deepcopy(first)
        _remove_local_attrs(newp, "paraId")
        for r in list(newp):
            if r.tag != W + "pPr":
                newp.remove(r)
        insert_text_into_p(newp, text)
        tc.append(newp)


def is_missing(value):
    """判断取值是否属于“信息不可得”（应留空手工填写）。"""
    v = (value or "").strip()
    if not v:
        return True
    if v in MISSING_VALUES:
        return True
    # 未替换的占位符，如 XXXX.XX、xxx
    if re.search(r"[Xx]{3,}", v) or re.fullmatch(r"[*_\-\s]+", v):
        return True
    return False


def format_amount(value):
    """核定入库金额：纯数字补“万元”单位；已带单位或为“待调整完善后核定”等结论文字则原样写入。"""
    v = (value or "").strip()
    if not v:
        return ""
    if "万元" in v:
        return v
    if re.fullmatch(r"[\d,]+(\.\d+)?\s*", v):
        return "%s 万元" % v.strip()
    return v


def check_conclusion(tc, option):
    """勾选评审结论：先把格内所有方框复位为空框，再把指定结论前的方框换成带勾方框。

    复位保证以旧版 word 为底本重新生成时，上一版的勾选会被清除。
    """
    for sym in tc.iter(W + "sym"):
        sym.set(W + "char", CHECK_BOX_EMPTY)
    for p in tc.findall(W + "p"):
        last_sym = None
        for r in p.findall(W + "r"):
            syms = r.findall(W + "sym")
            if syms:
                last_sym = syms[0]
                continue
            txt = "".join(t.text or "" for t in r.findall(W + "t"))
            if txt.strip() == option and last_sym is not None:
                last_sym.set(W + "char", CHECK_BOX_CHECKED)
                return True
    return False


def fill_basic_info(rows, fields, notes):
    """项目基本信息、评审结论勾选、核定入库金额：信息可得则填，不可得则留空并记录待办。"""
    filled = 0
    for cells in rows:
        for i, tc in enumerate(cells):
            label = clean(get_cell_text(tc))
            if label not in BASIC_FIELDS and label != CONCLUSION_FIELD and label != AMOUNT_FIELD:
                continue
            if i + 1 >= len(cells):
                continue
            target = cells[i + 1]
            value = fields.get(label, "")

            if label == CONCLUSION_FIELD:
                if is_missing(value):
                    notes.append("评审结论：md 未确定，请在 word 中手工勾选")
                    continue
                hits = [opt for opt in CONCLUSION_OPTIONS if opt in value]
                if len(hits) != 1:
                    notes.append("评审结论：md 中未唯一确定（%s），请在 word 中手工勾选" % value)
                    continue
                if check_conclusion(target, hits[0]):
                    filled += 1
                else:
                    notes.append("评审结论：模板未找到可勾选方框，请在 word 中手工勾选")
                continue

            if is_missing(value):
                # 信息不可得：底本若是旧版 word 则沿用其原值，否则留空手工填写
                old = clean(get_cell_text(target))
                if old:
                    notes.append("%s：md 未提供，已沿用底本原值「%s」" % (label, old))
                else:
                    notes.append("%s：md 未提供，请在 word 中手工填写" % label)
                continue

            text = format_amount(value) if label == AMOUNT_FIELD else value
            fill_tc(target, [text])
            filled += 1
    return filled


def locate_cells(body, head_text):
    """返回 body 下第一个 w:tbl 的所有行（每行为 tc 列表）。"""
    tbl = body.find(W + "tbl")
    if tbl is None:
        return None
    rows = []
    for tr in tbl.findall(W + "tr"):
        rows.append([c for c in tr if c.tag == W + "tc"])
    return rows


def build_docx(document_xml_text):
    """解析 document.xml 文本为 Element 树。"""
    register_xmlns(document_xml_text)
    return ET.fromstring(document_xml_text)


def fill_document(root, summary_items, body_by_title, fields, notes):
    """按模板结构填充。返回写入的 (概要条数, 子项段数, 表头字段数)。"""
    body = root.find(W + "body")
    rows = locate_cells(body, "")
    if rows is None:
        raise RuntimeError("文档中未找到表格")

    n_sum = 0
    # ---- 0) 项目基本信息、评审结论勾选、核定入库金额 ----
    n_head = fill_basic_info(rows, fields, notes)

    # ---- 1) "评审意见"大栏 <- 评审意见概要 ----
    # 定位行首 label 为"评审意见"的行（该行第 2 个单元格为跨列内容格）
    for cells in rows:
        if len(cells) >= 2 and clean(get_cell_text(cells[0])) == "评审意见":
            if summary_items:
                fill_tc(cells[1], summary_items)
                n_sum = len(summary_items)
            break

    # ---- 2) "评审要点"表"评审意见"列 <- 三问三看三查三核 ----
    # 12 个子项行：二级指标在 cells[2]，评审意见在 cells[4]（三问/三核整组纵向合并，
    # 三看/三查为前两行合并 + 第三行独立）。按合并区把若干子项分段写入 restart 单元格。
    groups = []  # (target_tc, [段落...])
    cur = None
    n_body = 0
    for cells in rows:
        if len(cells) < 5:
            continue
        title = clean(get_cell_text(cells[2]))
        if title not in SUBTITLES:
            continue
        text = body_by_title.get(title, "")
        vm = get_vmerge(cells[4])
        if vm == "restart" or vm is None:
            if cur is not None:
                groups.append(cur)
            cur = [cells[4], [text]]
        else:  # continue
            if cur is not None:
                cur[1].append(text)
    if cur is not None:
        groups.append(cur)

    for target_tc, paras in groups:
        paras = [p for p in paras if p.strip()]
        if paras:
            fill_tc(target_tc, paras)
            n_body += len(paras)

    return n_sum, n_body, n_head


# ---------------- 入口 ----------------

def pick_latest_md(directory):
    """在目录中选取版本号最大的 评审意见_NNN.md；没有版本文件时回退到 评审意见.md。"""
    best, best_n, fallback = None, -1, None
    for root, _dirs, files in os.walk(directory):
        for fn in files:
            m = re.match(r"^评审意见_(\d+)\.md$", fn)
            if m:
                n = int(m.group(1))
                if n > best_n:
                    best_n, best = n, os.path.join(root, fn)
            elif fn == "评审意见.md":
                fallback = os.path.join(root, fn)
    return best or fallback


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    default_template = os.path.join(here, "..", "assets", "三问三看模板.docx")

    ap = argparse.ArgumentParser(description="将评审意见最终稿 md 回填到官方 word 评审表")
    ap.add_argument("--md", required=True, help="评审意见最终稿 Markdown 文件路径；也可传目录，自动选用其中版本号最大的评审意见")
    ap.add_argument("--out", default=None, help="输出 docx 路径（默认：与 md 同目录的 评审意见表_<当天日期>.docx）")
    ap.add_argument("--template", default=default_template, help="底本 docx：默认使用空白模板 assets/三问三看模板.docx；"
                                                                 "也可指定上一版 word（如 评审意见表_2026-09-01.docx）作为底本在其基础上更新")
    args = ap.parse_args()

    md_path = os.path.abspath(args.md)
    tpl_path = os.path.abspath(args.template)

    # --md 传目录时，自动选用版本号最大的评审意见
    if os.path.isdir(md_path):
        picked = pick_latest_md(md_path)
        if not picked:
            eprint("目录中未找到评审意见 Markdown（评审意见_NNN.md 或 评审意见.md）：%s" % md_path)
            sys.exit(1)
        print("自动选用最新版本：%s" % picked)
        md_path = picked

    # 输出默认命名：与 md 同目录的 评审意见表_<当天日期>.docx
    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        out_path = os.path.join(os.path.dirname(md_path), "评审意见表_%s.docx" % date.today().isoformat())
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

    summary_items, body_by_title, fields = parse_md(md_text)
    missing = [t for t in SUBTITLES if not body_by_title.get(t)]
    if missing:
        eprint("提示：以下二级指标在 md 中未找到对应正文（这些行将保持空白）：")
        for t in missing:
            eprint("  - %s" % t)

    with zipfile.ZipFile(tpl_path) as zin:
        xml_path = "word/document.xml"
        document_xml = zin.read(xml_path).decode("utf-8")
        root = build_docx(document_xml)
        notes = []
        n_sum, n_body, n_head = fill_document(root, summary_items, body_by_title, fields, notes)

        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == xml_path:
                    data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
                zout.writestr(item, data)

    print("已生成：%s" % out_path)
    print("  评审意见栏写入概要条目：%d 条" % n_sum)
    print("  评审要点表写入子项判断：%d 段" % n_body)
    print("  表头区自动填入：%d 项" % n_head)
    if notes:
        print("  以下项目需手工补全（模板中已留空）：")
        for n in notes:
            print("    - %s" % n)


if __name__ == "__main__":
    main()
