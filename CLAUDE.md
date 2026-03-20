# PPT 批量导出图片工具 — CLAUDE.md

## 项目规则

**语言要求：请始终用中文回复，包括所有分析、解释和说明。**

## 项目简介

将文件夹中的 PPT 文件批量导出为 PNG 图片，支持递归扫描子文件夹。
面向不懂技术的普通用户，提供中文图形界面，双击即可使用。

运行方式：
```bash
python3 main.py
```

打包方式：
```bash
bash build_app.sh   # macOS: 生成 dist/PPT转图片.app + .dmg
```

## 项目结构

```
ppt-batch-tool/
├── main.py                     # 入口
├── requirements.txt            # PyQt6, PyMuPDF, pyinstaller
├── build_app.sh                # macOS 打包脚本
├── README.md                   # 中文使用说明
├── core/
│   ├── scanner.py              # 递归扫描 PPT 文件
│   ├── filename_cleaner.py     # 文件名清理（去版权/平台标记）
│   └── converter.py            # 转换引擎（WPS COM / PowerPoint / LibreOffice）
└── ui/
    └── main_window.py          # GUI 主窗口
```

## 技术栈

| 层次 | 库 | 用途 |
|------|-----|------|
| GUI | PyQt6 | 主界面 |
| PDF渲染 | PyMuPDF (fitz) | LibreOffice 后端的 PDF→PNG |
| Windows COM | comtypes | 调用 WPS/PowerPoint |
| 打包 | PyInstaller | 双击运行的 .app / .exe |

## 转换后端

- **Windows**: WPS COM → PowerPoint COM → LibreOffice headless
- **macOS**: PowerPoint AppleScript → LibreOffice headless
