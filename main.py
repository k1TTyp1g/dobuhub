#!/usr/bin/env python
"""
文档管家 DocuHub — 桌面端文档聚合工具

启动方式：
    python main.py          # 图形界面模式
    python main.py --cli    # 命令行模式（批量处理）

技术栈：Python 3.8+ | python-docx | Tkinter | PyYAML

架构概览：
    main.py                  ← 入口
    core/
      style_engine.py        ← ⭐ 核心排版引擎（一键标准化）
      scanner.py             ← 文档扫描聚合器
      doc_converter.py       ← .doc → .docx 转换桥
      template_manager.py    ← 模板 CRUD
    models/
      document.py            ← 文档数据模型
      template.py            ← 模板数据模型
    ui/
      main_window.py         ← Tkinter 主窗口（三栏布局）
    templates/
      academic.yaml          ← 学术论文模板
      default.yaml           ← 默认公文模板
"""
import sys
import os

# 确保项目根目录在 sys.path 中
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def resource_path(relative_path: str) -> str:
    """
    获取资源文件的绝对路径。
    在 PyInstaller 打包后，资源文件被解压到 sys._MEIPASS。
    开发环境下直接使用相对路径。
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def main():
    """主入口"""
    # 命令行模式
    if "--cli" in sys.argv:
        _run_cli()
        return

    # GUI 模式
    try:
        from ui.main_window import DocuHubApp
        from core.template_manager import TemplateManager
        app = DocuHubApp()
        # 注入资源路径（PyInstaller 兼容）
        app.template_mgr = TemplateManager(template_dir=resource_path("templates/"))
        app.run()
    except ImportError as e:
        print(f"❌ UI 依赖缺失：{e}")
        print("请确保 Python 版本 >= 3.8，且 Tkinter 可用。")
        sys.exit(1)


def _run_cli():
    """命令行批处理模式"""
    import argparse
    from core.style_engine import StyleEngine
    from core.template_manager import TemplateManager

    parser = argparse.ArgumentParser(description="DocuHub 命令行模式")
    parser.add_argument("files", nargs="+", help="要处理的 .docx 文件")
    parser.add_argument("--template", "-t", default=None,
                        help="模板名称（默认：第一个可用模板）")
    parser.add_argument("--outdir", "-o", default=None,
                        help="输出目录（默认：与源文件同目录）")
    args = parser.parse_args()

    # 加载模板
    tm = TemplateManager(template_dir=resource_path("templates/"))
    tmpl = tm.get_template(args.template) if args.template else tm.get_default()
    if tmpl is None:
        print("❌ 未找到可用模板")
        sys.exit(1)

    print(f"📐 使用模板：{tmpl.name}")
    print(f"📄 待处理文档：{len(args.files)} 个")

    engine = StyleEngine(tmpl)
    engine.batch_format(
        args.files,
        output_dir=args.outdir,
        progress_callback=lambda c, t, n: print(f"  [{c}/{t}] {n}")
    )

    print(f"\n✅ 完成：成功 {engine.stats['ok']} 个，跳过 {engine.stats['skip']} 个，失败 {engine.stats['error']} 个")


if __name__ == "__main__":
    main()
