"""轻量校验：docx 中所有 r:id 是否都能在关系表中找到对应且部件存在。

默认检查 tests/baseline/ 下的基线 docx，也可在命令行传入其他 docx 路径。
更全面的回归比对请用 scripts/verify.py。
"""
import glob
import os
import posixpath
import re
import sys
import zipfile

DEFAULT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tests", "baseline"
)

args = sys.argv[1:]
files = [a for a in args if a.lower().endswith(".docx")] or \
    sorted(glob.glob(os.path.join(DEFAULT_DIR, "*.docx")))

if not files:
    print(f"未找到待检查的 docx（默认目录 {DEFAULT_DIR} 为空）")
    raise SystemExit(1)

for f in files:
    z = zipfile.ZipFile(f)
    names = z.namelist()
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    ids = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    doc = z.read("word/document.xml").decode("utf-8")
    used = set(re.findall(r'r:id="([^"]+)"', doc))
    missing = sorted(i for i in used if i not in ids)
    bad = []
    for _i, t in ids.items():
        if t.startswith(("http", "/")):
            continue
        # 相对路径以 word/_rels/ 为基准解析，再判断是否打包在内
        if posixpath.normpath(posixpath.join("word", t)) not in names:
            bad.append(t)
    n_parts = sum(1 for n in names if n.startswith("word/"))
    status = "OK" if not missing and not bad else f"MISSING={missing} BAD={bad}"
    print(f"{os.path.basename(f):40s} {status}  parts={n_parts}")
