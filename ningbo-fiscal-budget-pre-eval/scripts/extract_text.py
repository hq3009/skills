# -*- coding: utf-8 -*-
"""
读取 .doc / .docx / .pdf / 图片（png、jpg 等）中的文字内容，输出到终端或文件。

用途：阅读部门上报的事前绩效评估报告、预算申报表、绩效目标表与各类佐证材料，
把其中的文字取出，供后续摘录、核对与写作使用。

用法：
    python scripts/extract_text.py <文件或目录> [<文件或目录> ...] [选项]

示例：
    python scripts/extract_text.py 报告.docx                     # 单个文件
    python scripts/extract_text.py 材料目录 -r                    # 递归读取整个目录
    python scripts/extract_text.py 材料目录 -r --out 提取结果.txt  # 结果写入文件
    python scripts/extract_text.py 主报告.doc --plain             # 只输出纯文本
    python scripts/extract_text.py 扫描件.pdf --engine fitz        # 指定 PDF 引擎

选项：
    -r, --recursive        目录递归查找
    --ext LIST             扩展名过滤，逗号分隔（默认 doc,docx,pdf,png,jpg,jpeg,bmp,tif,tiff,gif,webp）
    --out FILE             结果写入该文件（UTF-8），终端只打印统计
    --plain                不加文件头，只输出纯文本（单文件时常用）
    --max-chars N          每个文件最多输出 N 个字符（0 表示不限，默认 0）
    --engine NAME          PDF 引擎：auto(默认)/fitz/pypdf/pdfminer
    --lang LANG            图片 OCR 语言（默认 chi_sim+eng，需已安装对应 tesseract 语言包）
    --list-only            只列出将要处理的文件，不提取

各格式的提取方式（自动降级，无需全部依赖）：
    .docx  标准库 zipfile + XML 解析（无第三方依赖）
    .doc   Windows 下只用 Word / WPS 文字的 COM（pywin32，依次尝试 Word.Application、KWPS.Application、
           WPS.Application）；其他系统用 LibreOffice(soffice) → antiword
    .pdf   PyMuPDF(fitz) → pypdf → pdfminer.six
    图片    pytesseract + tesseract OCR（需另行安装 tesseract 及中文语言包）
若某格式缺少可用引擎，脚本会给出明确的安装提示，不影响其他文件继续提取。
"""
import argparse
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

IS_WINDOWS = sys.platform.startswith("win")

# 读取 .doc 的 COM 程序标识：Microsoft Word 优先，其次 WPS 文字（WPS 兼容 Word COM，亦有自身标识）
WORD_PROGIDS = ["Word.Application", "KWPS.Application", "WPS.Application"]
PROGID_ENGINE = {"Word.Application": "word-com", "KWPS.Application": "wps-com", "WPS.Application": "wps-com"}

DEFAULT_EXTS = ["doc", "docx", "pdf", "png", "jpg", "jpeg", "bmp", "tif", "tiff", "gif", "webp"]
IMAGE_EXTS = ["png", "jpg", "jpeg", "bmp", "tif", "tiff", "gif", "webp"]


class ExtractError(Exception):
    pass


# ---------------- .docx：标准库解析 ----------------

def _paragraph_text(p):
    parts = []
    for node in p.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag in (W + "br", W + "cr"):
            parts.append("\n")
    return "".join(parts)


def extract_docx(path):
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise ExtractError("不是有效的 docx 文件（zip 结构损坏或实为旧版 .doc）")
    with zf:
        names = [n for n in zf.namelist() if n == "word/document.xml"]
        if not names:
            raise ExtractError("docx 中未找到 word/document.xml")
        lines = []
        for name in names:
            root = ET.fromstring(zf.read(name))
            for p in root.iter(W + "p"):
                lines.append(_paragraph_text(p))
    text = "\n".join(lines)
    return _tidy(text), "stdlib-xml"


# ---------------- .doc：Word COM → LibreOffice → antiword ----------------

def _set_attr(obj, name, value):
    """部分 COM 实现（如 WPS）可能不支持某些属性，忽略即可。"""
    try:
        setattr(obj, name, value)
    except Exception:
        pass


def _engine_name(progid, app):
    """按实际调用的应用程序判断引擎名：WPS 会兼容注册 Word.Application，需如实区分。"""
    try:
        loc = str(getattr(app, "Path", "") or "").lower()
    except Exception:
        loc = ""
    if "kingsoft" in loc or "wps" in loc:
        return "wps-com"
    return PROGID_ENGINE.get(progid, progid)


def _doc_via_word_com(path):
    """Windows 下用 Word / WPS 的 COM 读取 .doc，返回 (文本 or None, 引擎名, 失败原因)。"""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        return None, "", "缺少 pywin32（pip install pywin32）：%s" % exc

    pythoncom.CoInitialize()
    errors = []
    try:
        for progid in WORD_PROGIDS:
            app = doc = None
            try:
                app = win32.Dispatch(progid)
                _set_attr(app, "Visible", False)
                _set_attr(app, "DisplayAlerts", 0)
                doc = app.Documents.Open(
                    str(Path(path).resolve()), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False
                )
                text = doc.Content.Text
                if text and text.strip():
                    return text, _engine_name(progid, app), ""
                errors.append("%s(空文本)" % progid)
            except Exception as exc:
                errors.append("%s(%s)" % (progid, type(exc).__name__))
            finally:
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
                try:
                    if app is not None:
                        app.Quit()
                except Exception:
                    pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return None, "", "COM 调用失败：" + "、".join(errors)


def _run(cmd):
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=120)
    except (OSError, subprocess.SubprocessError):
        return None
    if proc.returncode != 0:
        return None
    return proc


def _doc_via_soffice(path):
    import shutil
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        proc = _run([soffice, "--headless", "--convert-to", "txt:Text", "--outdir", tmp, str(path)])
        if proc is None:
            return None
        txts = list(Path(tmp).glob("*.txt"))
        if not txts:
            return None
        return txts[0].read_text(encoding="utf-8", errors="ignore")


def _doc_via_antiword(path):
    import shutil
    if not shutil.which("antiword"):
        return None
    proc = _run(["antiword", str(path)])
    if proc is None:
        return None
    return proc.stdout.decode("utf-8", errors="ignore")


def extract_doc(path):
    if IS_WINDOWS:
        # Windows 上只用 pywin32 的 Word / WPS COM，不依赖 LibreOffice
        text, engine, reason = _doc_via_word_com(path)
        if text:
            return _tidy(text), engine
        raise ExtractError(
            "无法读取 .doc：Windows 下通过 Word / WPS 的 COM（pywin32）提取，%s。"
            "请确认已安装 Microsoft Word 或 WPS 文字，并安装 pywin32（pip install pywin32）；"
            "也可在 Word / WPS 中将文件另存为 .docx 后重试。" % (reason or "未取到文本")
        )
    # 非 Windows 环境：LibreOffice → antiword
    text = _doc_via_soffice(path)
    if text:
        return _tidy(text), "libreoffice"
    text = _doc_via_antiword(path)
    if text:
        return _tidy(text), "antiword"
    raise ExtractError(
        "无法读取 .doc：未找到 LibreOffice / antiword。"
        "请安装其中之一，或将文件另存为 .docx 后重试。"
    )


# ---------------- .pdf：PyMuPDF → pypdf → pdfminer ----------------

def _pdf_via_fitz(path):
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24.3
    except ImportError:
        try:
            import fitz  # 旧版模块名
        except ImportError:
            return None
    doc = fitz.open(str(path))
    try:
        return "\n".join(page.get_text("text") for page in doc)
    finally:
        doc.close()


def _pdf_via_pypdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    reader = PdfReader(str(path))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def _pdf_via_pdfminer(path):
    try:
        from pdfminer.high_level import extract_text as pm_extract
    except ImportError:
        return None
    return pm_extract(str(path))


PDF_ENGINES = {"fitz": _pdf_via_fitz, "pypdf": _pdf_via_pypdf, "pdfminer": _pdf_via_pdfminer}


def extract_pdf(path, engine="auto"):
    order = ["fitz", "pypdf", "pdfminer"] if engine == "auto" else [engine]
    tried = []
    for name in order:
        func = PDF_ENGINES.get(name)
        if not func:
            continue
        try:
            text = func(path)
        except Exception as exc:  # 单个引擎失败继续降级
            tried.append("%s(%s)" % (name, type(exc).__name__))
            continue
        if text and text.strip():
            return _tidy(text), name
        tried.append("%s(空)" % name)
    raise ExtractError(
        "PDF 未提取到文本（%s）。若为扫描件需先做 OCR，或改用 pdfplumber / 安装 PyMuPDF：pip install pymupdf"
        % "、".join(tried)
    )


# ---------------- 图片：OCR ----------------

def extract_image(path, lang="chi_sim+eng"):
    try:
        from PIL import Image
    except ImportError:
        raise ExtractError("图片 OCR 需要 Pillow：pip install pillow")
    try:
        import pytesseract
    except ImportError:
        raise ExtractError(
            "图片 OCR 需要 pytesseract：pip install pytesseract；"
            "并需安装 tesseract 引擎及中文语言包 chi_sim"
            "（Windows 可用 UB-Mannheim/tesseract 安装包，或 choco install tesseract）。"
            "若暂不安装，请人工提供图片中的文字。 "
            "（已检测到 Pillow，可先确认图片是否为纯扫描件：%s）" % Path(path).name
        )
    try:
        with Image.open(str(path)) as img:
            text = pytesseract.image_to_string(img, lang=lang)
    except Exception as exc:
        raise ExtractError("OCR 失败：%s（确认已安装 tesseract 及语言包 %s）" % (exc, lang))
    if not text.strip():
        raise ExtractError("OCR 未识别出文字（可换用 --lang 指定语言，或提高图片清晰度）")
    return _tidy(text), "pytesseract(%s)" % lang


# ---------------- 通用 ----------------

def _tidy(text):
    """压缩连续空行、去掉行尾空白。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_one(path, engine="auto", lang="chi_sim+eng"):
    """返回 (文本, 引擎名)。失败抛 ExtractError。"""
    ext = Path(path).suffix.lower().lstrip(".")
    if ext == "docx":
        return extract_docx(path)
    if ext == "doc":
        return extract_doc(path)
    if ext == "pdf":
        return extract_pdf(path, engine=engine)
    if ext in IMAGE_EXTS:
        return extract_image(path, lang=lang)
    raise ExtractError("不支持的扩展名：.%s" % ext)


def collect_files(inputs, recursive, exts):
    files = []
    for item in inputs:
        p = Path(item)
        if p.is_file():
            files.append(p)
        elif p.is_dir():
            pattern = "**/*" if recursive else "*"
            for f in sorted(p.glob(pattern)):
                if f.is_file() and f.suffix.lower().lstrip(".") in exts:
                    files.append(f)
        else:
            print("警告：路径不存在，已跳过：%s" % p, file=sys.stderr)
    # 去重并保持顺序
    seen = set()
    uniq = []
    for f in files:
        key = str(f.resolve()).lower()
        if key not in seen:
            seen.add(key)
            uniq.append(f)
    return uniq


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description="读取 doc/docx/pdf/图片中的文字内容")
    ap.add_argument("paths", nargs="+", help="文件或目录（可多个）")
    ap.add_argument("-r", "--recursive", action="store_true", help="目录递归查找")
    ap.add_argument("--ext", default=",".join(DEFAULT_EXTS), help="扩展名过滤，逗号分隔")
    ap.add_argument("--out", default=None, help="结果写入文件（UTF-8）")
    ap.add_argument("--plain", action="store_true", help="只输出纯文本，不加文件头")
    ap.add_argument("--max-chars", type=int, default=0, help="每个文件最多输出字符数，0 为不限")
    ap.add_argument("--engine", default="auto", choices=["auto", "fitz", "pypdf", "pdfminer"], help="PDF 引擎")
    ap.add_argument("--lang", default="chi_sim+eng", help="图片 OCR 语言")
    ap.add_argument("--list-only", action="store_true", help="只列出待处理文件")
    args = ap.parse_args()

    exts = [e.strip().lower().lstrip(".") for e in args.ext.split(",") if e.strip()]
    files = collect_files(args.paths, args.recursive, exts)
    if not files:
        print("未找到可处理的文件。")
        return 1

    if args.list_only:
        for f in files:
            print(f)
        print("共 %d 个文件" % len(files))
        return 0

    chunks = []
    ok, fail = 0, 0

    def emit(text):
        if args.out:
            chunks.append(text)
        else:
            print(text)

    for idx, path in enumerate(files, 1):
        head = "" if args.plain else "===== [%d/%d] %s" % (idx, len(files), path)
        try:
            text, engine = extract_one(path, engine=args.engine, lang=args.lang)
        except ExtractError as exc:
            fail += 1
            emit("" if args.plain else head)
            emit("!!! 提取失败：%s" % exc)
            continue
        except Exception as exc:  # 兜底，避免单个文件中断整批
            fail += 1
            emit("" if args.plain else head)
            emit("!!! 提取失败（未预期错误）：%s: %s" % (type(exc).__name__, exc))
            continue

        ok += 1
        if args.max_chars and len(text) > args.max_chars:
            text = text[: args.max_chars] + "\n...[已截断，原文共 %d 字]" % len(text)
        if args.plain:
            emit(text)
        else:
            emit("%s（引擎：%s，%d 字）=====" % (head, engine, len(text)))
            emit(text)
            emit("")

    summary = "----- 完成：成功 %d 个，失败 %d 个，共 %d 个 -----" % (ok, fail, len(files))
    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text("\n".join(chunks) + "\n" + summary + "\n", encoding="utf-8")
        print("已写入：%s" % out_path)
    print(summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
