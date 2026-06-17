"""排版模板数据模型"""
import yaml
from dataclasses import dataclass, field, asdict
from typing import Optional


@dataclass
class PageSettings:
    paper_size: str = "A4"
    margin_top: float = 2.54
    margin_bottom: float = 2.54
    margin_left: float = 3.18
    margin_right: float = 3.18


@dataclass
class BodySettings:
    font_name_cn: str = "宋体"
    font_name_en: str = "Times New Roman"
    font_size: int = 12           # pt
    bold: bool = False
    italic: bool = False
    alignment: str = "justify"
    first_line_indent: int = 2    # 字符数
    line_spacing_type: str = "fixed"  # fixed / multiple / single
    line_spacing: float = 20
    space_before: int = 0
    space_after: int = 0


@dataclass
class HeadingSettings:
    font_name_cn: str = "黑体"
    font_name_en: str = "Arial"
    font_size: int = 16
    bold: bool = True
    italic: bool = False
    alignment: str = "center"
    first_line_indent: int = 0
    line_spacing_type: str = "fixed"
    line_spacing: float = 20
    space_before: int = 12
    space_after: int = 6


@dataclass
class FormatTemplate:
    """完整的格式模板"""
    name: str = "默认"
    description: str = ""
    version: float = 1.0
    page: PageSettings = field(default_factory=PageSettings)
    body: BodySettings = field(default_factory=BodySettings)
    headings: dict = field(default_factory=lambda: {
        1: HeadingSettings(),  # 一级标题
        2: HeadingSettings(font_size=14, alignment="left", space_before=8, space_after=4),
        3: HeadingSettings(font_name_cn="楷体", font_name_en="Times New Roman",
                           font_size=12, alignment="left", space_before=6, space_after=2),
    })

    @classmethod
    def from_yaml(cls, path: str) -> "FormatTemplate":
        """从 YAML 文件加载模板"""
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        t = cls(
            name=data.get("name", "未命名"),
            description=data.get("description", ""),
            version=data.get("version", 1.0),
        )

        # 解析 page 节
        p = data.get("page", {})
        t.page = PageSettings(
            paper_size=p.get("paper_size", "A4"),
            margin_top=float(p.get("margin_top", 2.54)),
            margin_bottom=float(p.get("margin_bottom", 2.54)),
            margin_left=float(p.get("margin_left", 3.18)),
            margin_right=float(p.get("margin_right", 3.18)),
        )

        # 解析 body 节
        b = data.get("body", {})
        t.body = BodySettings(
            font_name_cn=b.get("font_name_cn", "宋体"),
            font_name_en=b.get("font_name_en", "Times New Roman"),
            font_size=int(b.get("font_size", 12)),
            bold=b.get("bold", False),
            italic=b.get("italic", False),
            alignment=b.get("alignment", "justify"),
            first_line_indent=int(b.get("first_line_indent", 2)),
            line_spacing_type=b.get("line_spacing_type", "fixed"),
            line_spacing=float(b.get("line_spacing", 20)),
            space_before=int(b.get("space_before", 0)),
            space_after=int(b.get("space_after", 0)),
        )

        # 解析 headings 节
        h = data.get("headings", {})
        for level in [1, 2, 3]:
            hl = h.get(f"level{level}", {})
            t.headings[level] = HeadingSettings(
                font_name_cn=hl.get("font_name_cn", t.body.font_name_cn),
                font_name_en=hl.get("font_name_en", t.body.font_name_en),
                font_size=int(hl.get("font_size", t.body.font_size)),
                bold=hl.get("bold", level == 1),
                italic=hl.get("italic", False),
                alignment=hl.get("alignment", "center" if level == 1 else "left"),
                first_line_indent=int(hl.get("first_line_indent", 0)),
                line_spacing_type=hl.get("line_spacing_type", t.body.line_spacing_type),
                line_spacing=float(hl.get("line_spacing", 20)),
                space_before=int(hl.get("space_before", 0)),
                space_after=int(hl.get("space_after", 0)),
            )

        return t

    def to_yaml(self, path: str):
        """导出模板到 YAML 文件"""
        data = {
            "name": self.name,
            "description": self.description,
            "version": self.version,
            "page": asdict(self.page),
            "body": asdict(self.body),
            "headings": {
                f"level{k}": asdict(v)
                for k, v in self.headings.items()
            },
        }
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
