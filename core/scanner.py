"""
全局文档聚合与搜索引擎
======================
智能扫描预设路径和自定义路径，提取 .doc/.docx 文档。
"""
import os
import time
from pathlib import Path
from typing import Callable, Optional, List
from datetime import datetime

from models.document import DocumentItem


# 常见的文档散落路径关键词（Windows）
COMMON_DOWNLOAD_PATHS = [
    # 系统下载夹
    "Downloads",
    "下载",
    # 桌面
    "Desktop",
    "桌面",
    # 微信文件
    "WeChat Files",
    "WeChat",
    "FileStorage",
    "File",
    # 腾讯会议 / QQ
    "Tencent",
    "QQ",
    # 浏览器
    "Chrome",
    "Edge",
    "360Chrome",
]


# 在 Windows 上通过已知的用户目录扩展搜索范围
def _get_user_profile_paths(username: str = "k1T") -> list:
    """根据常见用户目录结构生成扫描路径候选"""
    base = f"C:/Users/{username}"
    candidates = [
        f"{base}/Desktop",
        f"{base}/Downloads",
        f"{base}/Documents",
        f"{base}/Desktop/文档",
        f"{base}/Desktop/工作文档",
        f"{base}/Downloads/微信文件",
        f"{base}/AppData/Local/Packages",
    ]
    return [p for p in candidates if os.path.isdir(p)]


class DocScanner:
    """文档扫描器：递归遍历目录，收集 .doc/.docx 文件"""

    # 需要过滤的系统目录关键词
    EXCLUDE_DIR_PARTS = [
        "System32", "Windows", "Program Files", "Program Files (x86)",
        "AppData/Local/Temp", "AppData/LocalLow",
        "node_modules", ".git", "__pycache__", ".svn",
        "cache", "Cache", "CACHE",
    ]

    def __init__(self):
        self.documents = []  # type: List[DocumentItem]
        self.paths = set()  # 去重路径

    def scan(self, paths: List[str],
             progress_callback: Optional[Callable] = None):
        """
        扫描指定路径列表，返回文档列表。

        参数：
            paths:         要扫描的目录或文件路径列表
            progress_callback: f(current, total, path_name)

        返回：
            DocumentItem 列表
        """
        self.documents = []
        seen_paths = set()

        # 展开目录 + 收集文件
        file_candidates = []
        for path in paths:
            path = os.path.normpath(path)
            if not os.path.exists(path):
                continue

            if os.path.isfile(path):
                file_candidates.append(path)
            elif os.path.isdir(path):
                file_candidates.extend(self._walk_dir(path, progress_callback))

        # 过滤去重 + 提取元数据
        total = len(file_candidates)
        for i, fpath in enumerate(file_candidates):
            norm = os.path.normpath(fpath).lower()
            if norm in seen_paths:
                continue
            seen_paths.add(norm)

            if progress_callback:
                progress_callback(i + 1, total, os.path.basename(fpath))

            ext = os.path.splitext(fpath)[1].lower()
            if ext not in (".doc", ".docx"):
                continue

            try:
                stat = os.stat(fpath)
                doc = DocumentItem(
                    path=fpath,
                    filename=os.path.basename(fpath),
                    ext=ext,
                    size_bytes=stat.st_size,
                    modified_time=datetime.fromtimestamp(stat.st_mtime),
                    source_dir=self._detect_source(fpath),
                )
                self.documents.append(doc)
            except OSError:
                continue

        return self.documents

    def _walk_dir(self, directory: str,
                  progress_callback: Optional[Callable] = None) -> List[str]:
        """递归遍历目录，收集 .doc/.docx 文件路径"""
        files = []
        try:
            for root, dirs, filenames in os.walk(directory, topdown=True):
                # 过滤掉系统目录
                dirs[:] = [
                    d for d in dirs
                    if not self._should_exclude(os.path.join(root, d))
                ]

                for fn in filenames:
                    ext = os.path.splitext(fn)[1].lower()
                    if ext in (".doc", ".docx"):
                        files.append(os.path.join(root, fn))
        except (PermissionError, OSError):
            pass
        return files

    def _should_exclude(self, dir_path: str) -> bool:
        """判断目录是否应该排除"""
        dir_lower = dir_path.lower()
        for part in self.EXCLUDE_DIR_PARTS:
            if part.lower() in dir_lower:
                return True
        return False

    @staticmethod
    def _detect_source(file_path: str) -> str:
        """根据路径推断来源标签"""
        pl = file_path.lower()

        if "desktop" in pl or "桌面" in pl:
            return "桌面"
        if "download" in pl:
            return "下载"
        if "wechat" in pl or "微信" in pl or "wx" in pl.replace("\\", "/").split("/"):
            return "微信"
        if "qq" in pl:
            return "QQ"
        if "documents" in pl or "文档" in pl:
            return "文档"
        return "其他"

    @staticmethod
    def search(documents: List[DocumentItem], keyword: str = "",
               sort_by: str = "filename", reverse: bool = False) -> List[DocumentItem]:
        """
        在已扫描的文档列表中搜索/排序。

        参数：
            documents:  扫描结果列表
            keyword:    文件名关键词（模糊匹配）
            sort_by:    排序字段 (filename / modified_time / size_bytes)
            reverse:    是否倒序

        返回：
            过滤/排序后的列表
        """
        results = documents[:]

        # 关键词过滤
        if keyword.strip():
            kw = keyword.strip().lower()
            results = [
                d for d in results
                if kw in d.filename.lower()
            ]

        # 排序
        if sort_by == "modified_time":
            results.sort(key=lambda d: d.modified_time or datetime.min,
                         reverse=reverse)
        elif sort_by == "size_bytes":
            results.sort(key=lambda d: d.size_bytes, reverse=reverse)
        else:
            results.sort(key=lambda d: d.filename.lower(), reverse=reverse)

        return results

    def get_stats(self) -> dict:
        """返回扫描统计信息"""
        ext_counts = {".doc": 0, ".docx": 0}
        source_counts = {}

        for doc in self.documents:
            ext_counts[doc.ext] = ext_counts.get(doc.ext, 0) + 1
            src = doc.source_dir
            source_counts[src] = source_counts.get(src, 0) + 1

        return {
            "total": len(self.documents),
            "by_ext": ext_counts,
            "by_source": source_counts,
        }
