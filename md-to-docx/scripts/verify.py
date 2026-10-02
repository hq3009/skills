"""回归验证：把 examples/ 下的样例 md 重新转换，与 tests/baseline/ 下的基线 docx 逐项比对。

用法：
    python scripts/verify.py              # 比对，全部一致则退出码 0
    python scripts/verify.py --update     # 重新生成基线（版式 preset 有意调整后使用）
    python scripts/verify.py --keep       # 保留本次生成的临时 docx，便于人工打开核对

比对项：段落样式名/文本/对齐/缩进/行距/段前后间距、每个 run 的字体 XML（字体、字号、
加粗）、表格结构与单元格文本、页面与页边距、页眉页脚 XML（含 STYLEREF）、图片部件，
以及 r:id 引用完整性。
"""
import argparse
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

from docx import Document
from docx.shared import Length

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXAMPLES = os.path.join(ROOT, "examples")
BASELINE = os.path.join(ROOT, "tests", "baseline")
CONVERTER = os.path.join(ROOT, "scripts", "md2docx.py")

CASES = [
    ("sample-official.md", "official"),
    ("sample-official.md", "zju"),
    ("sample-figure.md", "official"),
    ("sample-figure.md", "zju"),
    ("sample-section.md", "official"),   # 分节符：下一页/连续/偶数页/奇数页
    ("sample-section.md", "zju"),
]

MAX_DIFFS = 20


def baseline_name(md: str, mode: str) -> str:
    return f"{os.path.splitext(md)[0]}-{mode}.docx"


def _pt(v):
    """统一换算成磅并保留 3 位小数，避免 EMU 与倍数的精度差异造成误报。"""
    if v is None:
        return None
    if isinstance(v, Length):
        return round(v.pt, 3)
    if isinstance(v, (int, float)):
        return round(float(v), 3)
    return str(v)


def _runs(p):
    out = []
    for r in p.runs:
        rpr = r._element.rPr
        out.append((r.text, rpr.xml if rpr is not None else ""))
    return out


def rid_problems(path: str) -> list:
    """r:id 引用完整性：引用的关系 id 是否存在、目标部件是否打包在内。"""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    ids = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    doc = z.read("word/document.xml").decode("utf-8")
    used = set(re.findall(r'r:id="([^"]+)"', doc))
    bad = [i for i in sorted(used) if i not in ids]
    for _i, t in ids.items():
        if t.startswith(("http", "/")):
            continue
        # 相对路径以 word/_rels/ 为基准解析，再判断是否打包在内
        if posixpath.normpath(posixpath.join("word", t)) not in names:
            bad.append(f"target:{t}")
    return bad


def summarize(path: str) -> dict:
    doc = Document(path)

    paras = []
    for p in doc.paragraphs:
        pf = p.paragraph_format
        paras.append({
            "style": p.style.name,
            "text": p.text,
            "align": str(pf.alignment),
            "first_line": _pt(pf.first_line_indent),
            "left": _pt(pf.left_indent),
            "line": _pt(pf.line_spacing),
            "before": _pt(pf.space_before),
            "after": _pt(pf.space_after),
            "runs": _runs(p),
        })

    tables = []
    for t in doc.tables:
        rows = [[c.text for c in row.cells] for row in t.rows]
        header_rpr = ""
        if t.rows and t.rows[0].cells:
            runs = t.rows[0].cells[0].paragraphs[0].runs
            if runs:
                rpr = runs[0]._element.rPr
                header_rpr = rpr.xml if rpr is not None else ""
        tables.append({"rows": rows, "header_rpr": header_rpr})

    s = doc.sections[0]
    section = {
        "page": (_pt(s.page_width), _pt(s.page_height)),
        "margins": (_pt(s.top_margin), _pt(s.bottom_margin),
                    _pt(s.left_margin), _pt(s.right_margin)),
        "header_dist": _pt(s.header_distance),
        "footer_dist": _pt(s.footer_distance),
    }

    z = zipfile.ZipFile(path)
    names = z.namelist()
    stories = {}
    for n in sorted(names):
        if re.match(r"word/(header|footer)\d+\.xml$", n):
            stories[n] = re.sub(r"\s+", " ", z.read(n).decode("utf-8")).strip()

    return {
        "paragraphs": paras,
        "tables": tables,
        "section": section,
        "stories": stories,
        "media": sorted(n for n in names if n.startswith("word/media/")),
        "rid": rid_problems(path),
    }


def compare(a, b, path="") -> list:
    diffs = []
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            if k not in a:
                diffs.append(f"{path}.{k}: 基线缺失")
            elif k not in b:
                diffs.append(f"{path}.{k}: 新产物缺失")
            else:
                diffs += compare(a[k], b[k], f"{path}.{k}")
    elif isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            diffs.append(f"{path}: 数量 {len(a)} != {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            diffs += compare(x, y, f"{path}[{i}]")
            if len(diffs) > MAX_DIFFS:
                break
    elif a != b:
        diffs.append(f"{path}: {a!r} != {b!r}")
    return diffs


def convert(md: str, mode: str, out: str) -> None:
    cmd = [sys.executable, CONVERTER, os.path.join(EXAMPLES, md),
           "-o", out, "--mode", mode, "--no-font-check"]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        sys.stderr.write(r.stdout + r.stderr)
        raise SystemExit(f"转换失败: {md} ({mode})")


def main() -> int:
    ap = argparse.ArgumentParser(description="md-to-docx 回归验证")
    ap.add_argument("--update", action="store_true", help="重新生成基线文件")
    ap.add_argument("--keep", action="store_true", help="保留本次生成的 docx")
    args = ap.parse_args()

    if args.update:
        os.makedirs(BASELINE, exist_ok=True)
        for md, mode in CASES:
            convert(md, mode, os.path.join(BASELINE, baseline_name(md, mode)))
        print(f"基线已更新 -> {BASELINE}")
        return 0

    tmp = tempfile.mkdtemp(prefix="md2docx-verify-")
    failed = 0
    try:
        for md, mode in CASES:
            name = baseline_name(md, mode)
            base = os.path.join(BASELINE, name)
            if not os.path.isfile(base):
                print(f"[FAIL] {name}: 基线文件缺失")
                failed += 1
                continue
            out = os.path.join(tmp, name)
            convert(md, mode, out)
            diffs = compare(summarize(base), summarize(out))
            if diffs:
                failed += 1
                print(f"[FAIL] {name}: {len(diffs)} 处差异")
                for d in diffs[:MAX_DIFFS]:
                    print("       - " + d)
                if len(diffs) > MAX_DIFFS:
                    print(f"       ... 另有 {len(diffs) - MAX_DIFFS} 处")
            else:
                print(f"[PASS] {name}")
        if args.keep:
            print(f"临时产物保留于: {tmp}")
    finally:
        if not args.keep:
            shutil.rmtree(tmp, ignore_errors=True)

    total = len(CASES)
    print(f"合计 {total - failed}/{total} 通过")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
