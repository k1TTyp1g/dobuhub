"""
.doc（旧格式）→ .docx 转换器
=============================
python-docx 只能处理 .docx，旧版 .doc 需要先转换。

转换策略（按优先级）：
  1. win32com  — 调用本机 Microsoft Word（最精准，需安装 Office）
  2. LibreOffice — 跨平台，通过 CLI 转换
  3. 人工提示  — 告知用户需转换后再处理
"""
import os
import subprocess
import sys
from pathlib import Path
from typing import Optional


def can_convert_doc() -> tuple:
    """检测当前环境是否支持 .doc 转换，返回 (支持, 引擎名)"""
    # 1. 检测 win32com (Microsoft Office)
    try:
        import win32com.client
        # 尝试创建 Word 应用实例
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True, "win32com (Microsoft Word)"
    except Exception:
        pass

    # 2. 检测 LibreOffice
    for candidate in [
        "soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "C:/Program Files/LibreOffice/program/soffice.exe",
        "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    ]:
        try:
            subprocess.run([candidate, "--version"],
                           capture_output=True, timeout=10)
            return True, f"LibreOffice ({candidate})"
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    return False, ""


def convert_doc_to_docx(src_path: str, dst_path: Optional[str] = None) -> str:
    """将 .doc 转换为 .docx"""
    if dst_path is None:
        dst_path = str(Path(src_path).with_suffix(".docx"))

    # 如果目标已存在，直接返回
    if os.path.exists(dst_path):
        return dst_path

    # 方法 1: win32com
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        try:
            doc = word.Documents.Open(os.path.abspath(src_path))
            doc.SaveAs(os.path.abspath(dst_path), FileFormat=16)  # 16 = wdFormatDocumentDefault
            doc.Close()
            word.Quit()
            return dst_path
        except Exception:
            word.Quit()
            raise
    except ImportError:
        pass

    # 方法 2: LibreOffice
    for candidate in [
        "soffice",
        "C:/Program Files/LibreOffice/program/soffice.exe",
        "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    ]:
        try:
            subprocess.run(
                [candidate, "--headless", "--convert-to", "docx",
                 "--outdir", str(Path(dst_path).parent),
                 os.path.abspath(src_path)],
                capture_output=True, timeout=120,
            )
            if os.path.exists(dst_path):
                return dst_path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    raise RuntimeError(
        "无法转换 .doc 文件。请安装 Microsoft Office 或 LibreOffice。\n"
        f"源文件：{src_path}"
    )


def batch_convert_to_docx(file_list: list,
                          progress_callback=None) -> dict:
    """
    批量将 .doc 转为 .docx。

    返回：
        {"converted": {"原路径": "新路径", ...},
         "failed": ["原路径", ...],
         "skipped": ["原路径", ...]}
    """
    result = {"converted": {}, "failed": [], "skipped": []}

    for i, src_path in enumerate(file_list):
        ext = Path(src_path).suffix.lower()

        if ext != ".doc":
            result["skipped"].append(src_path)
            continue

        if progress_callback:
            progress_callback(i + 1, len(file_list), Path(src_path).name)

        try:
            dst = convert_doc_to_docx(src_path)
            result["converted"][src_path] = dst
        except Exception as e:
            result["failed"].append(src_path)

    return result
