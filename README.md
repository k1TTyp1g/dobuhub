# DocuHub 文档管家

一款 Windows 桌面文档管理工具，支持多格式文件浏览、预览、搜索。

## 功能特性

- 📁 **文件浏览器** — 多级目录穿透扫描，支持全盘搜索
- 🔍 **类型筛选** — 勾选文件类型即时过滤
- 🖼️ **文件预览** — 支持图片、文档（docx/doc/pdf/ppt）、文本、CSV 等格式
- 🗑️ **安全删除** — 移到回收站，防止误删
- 📌 **收藏文件夹** — 锁定常用目录，快速切换
- 📋 **最近打开** — 自动记录最近文件

## 支持的文件类型

| 类别 | 格式 |
|------|------|
| Word | .doc .docx .wps .rtf .odt |
| Excel | .xls .xlsx .csv .et |
| PPT | .ppt .pptx .pps .dps |
| 图片 | .png .jpg .jpeg .gif .bmp .webp .tiff |
| PDF | .pdf |
| 文本 | .txt .md .py .js .html .css .json .xml .log |
| 压缩包 | .zip .rar .7z |
| 邮件 | .eml .msg |

## 快速开始

### 方法一：直接运行

下载 `DocuHub_Setup.exe` 安装即可。

### 方法二：源码运行

```bash
# 安装依赖
pip install pillow pywin32 python-docx python-pptx PyPDF2 pyyaml ttkthemes

# 运行
python main.py
```

## 环境要求

- Windows 10 / 11
- Python 3.8+（源码运行）
- 无需网络连接

## 技术栈

- **UI 框架**: Tkinter + ttkthemes (Breeze 主题)
- **文档处理**: python-docx, python-pptx, PyPDF2, win32com
- **图像处理**: Pillow
- **打包**: PyInstaller

## 项目结构

```
docuhub/
├── main.py                 # 入口文件
├── core/
│   ├── style_engine.py     # 排版引擎
│   ├── scanner.py          # 文档扫描
│   ├── doc_converter.py    # 格式转换
│   └── template_manager.py # 模板管理
├── ui/
│   └── main_window.py      # 主窗口
├── models/
│   ├── document.py         # 文档模型
│   └── template.py         # 模板模型
├── templates/              # 排版模板
└── resources/              # 资源文件
```

## License

MIT
