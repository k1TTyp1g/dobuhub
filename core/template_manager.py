"""
排版模板管理器
==============
管理预设模板和用户自定义模板的加载/保存/删除。
"""
import os
import glob
from typing import Optional, List

from models.template import FormatTemplate


class TemplateManager:
    """模板管理器"""

    def __init__(self, template_dir: str = "templates/"):
        self.template_dir = template_dir
        self._templates: dict[str, FormatTemplate] = {}
        self._refresh()

    # ── 加载 ──────────────────────────────────────────────

    def _refresh(self):
        """扫描模板目录，加载所有 .yaml 模板"""
        os.makedirs(self.template_dir, exist_ok=True)
        self._templates.clear()

        yaml_files = glob.glob(os.path.join(self.template_dir, "*.yaml"))
        yaml_files += glob.glob(os.path.join(self.template_dir, "*.yml"))

        for yf in sorted(yaml_files):
            try:
                tmpl = FormatTemplate.from_yaml(yf)
                self._templates[tmpl.name] = tmpl
            except Exception as e:
                print(f"⚠ 加载模板失败 {yf}: {e}")

    def list_templates(self):
        """返回所有可用模板名称列表"""
        self._refresh()
        return list(self._templates.keys())

    def get_template(self, name: str) -> Optional[FormatTemplate]:
        """按名称获取模板"""
        self._refresh()
        return self._templates.get(name)

    def get_default(self) -> Optional[FormatTemplate]:
        """获取第一个可用模板作为默认"""
        self._refresh()
        if self._templates:
            return next(iter(self._templates.values()))
        return None

    # ── 增删 ──────────────────────────────────────────────

    def save_template(self, template: FormatTemplate) -> str:
        """保存模板到文件"""
        safe_name = template.name.replace(" ", "_").replace("/", "_")
        path = os.path.join(self.template_dir, f"{safe_name}.yaml")
        template.to_yaml(path)
        self._templates[template.name] = template
        return path

    def delete_template(self, name: str) -> bool:
        """删除模板"""
        tmpl = self._templates.get(name)
        if tmpl is None:
            return False

        safe_name = name.replace(" ", "_").replace("/", "_")
        # 尝试 .yaml 和 .yml
        for ext in (".yaml", ".yml"):
            path = os.path.join(self.template_dir, f"{safe_name}{ext}")
            if os.path.exists(path):
                os.remove(path)
                break

        self._templates.pop(name, None)
        return True

    # ── 默认模板路径 ──────────────────────────────────────

    @staticmethod
    def get_builtin_templates() -> dict:
        """
        返回内置模板定义（当模板文件丢失时可自动重建）
        """
        return {
            "default.yaml": """\
name: "默认规范"
description: "标准公文/论文格式"
version: 1.0
page:
  paper_size: "A4"
  margin_top: 2.54
  margin_bottom: 2.54
  margin_left: 3.18
  margin_right: 3.18
body:
  font_name_cn: "宋体"
  font_name_en: "Times New Roman"
  font_size: 12
  bold: false
  italic: false
  alignment: "justify"
  first_line_indent: 2
  line_spacing_type: "fixed"
  line_spacing: 20
  space_before: 0
  space_after: 0
headings:
  level1:
    font_name_cn: "黑体"
    font_name_en: "Arial"
    font_size: 16
    bold: true
    alignment: "center"
    space_before: 12
    space_after: 6
  level2:
    font_name_cn: "黑体"
    font_name_en: "Arial"
    font_size: 14
    bold: true
    alignment: "left"
    space_before: 8
    space_after: 4
  level3:
    font_name_cn: "楷体"
    font_name_en: "Times New Roman"
    font_size: 12
    bold: true
    alignment: "left"
    space_before: 6
    space_after: 2
""",
        }
