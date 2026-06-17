"""文档数据模型"""
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime
from typing import Optional


@dataclass
class DocumentItem:
    """单个文档的元数据"""
    path: str               # 完整路径
    filename: str           # 文件名（含扩展名）
    ext: str                # 扩展名 .doc / .docx
    size_bytes: int = 0
    modified_time: Optional[datetime] = None
    source_dir: str = ""    # 来源目录标签（如"桌面"、"下载"、"微信"）

    def formatted_size(self) -> str:
        """人性化显示文件大小"""
        if self.size_bytes < 1024:
            return f"{self.size_bytes} B"
        elif self.size_bytes < 1024 * 1024:
            return f"{self.size_bytes / 1024:.1f} KB"
        else:
            return f"{self.size_bytes / 1024 / 1024:.1f} MB"

    def formatted_time(self) -> str:
        """格式化修改时间"""
        if self.modified_time:
            return self.modified_time.strftime("%Y-%m-%d %H:%M")
        return "未知"

    @property
    def stem(self) -> str:
        return Path(self.path).stem
