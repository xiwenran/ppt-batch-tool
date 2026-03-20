# PPT 批量导出图片工具

把文件夹中的 PPT 文件批量导出为 PNG 图片，支持递归扫描子文件夹。

## 功能

- 递归扫描指定文件夹中所有 PPT 文件（`.ppt` `.pptx` `.pps` `.ppsx`）
- 每个 PPT 导出前 N 页为 PNG 图片（默认 17 页，可调整）
- 自动清理文件名（去除版权声明、平台标记等）
- 每个 PPT 的图片存入独立文件夹
- 显示详细的成功/失败统计和日志
- 中文图形界面，双击即可使用

## 转换引擎

程序会自动检测并使用你电脑上已安装的办公软件：

| 平台 | 优先使用 | 备选 |
|------|---------|------|
| Windows | WPS Office | Microsoft PowerPoint → LibreOffice |
| macOS | Microsoft PowerPoint | LibreOffice |

> 如果以上都没有，需要安装免费的 [LibreOffice](https://www.libreoffice.org/download/)（安装一次即可）。

## 使用方法

1. 双击打开程序
2. 选择包含 PPT 文件的文件夹
3. 设置导出页数（默认 17 页）
4. 点击「开始转换」
5. 等待完成，查看结果

## 从源码运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行
python main.py
```

## 打包

### macOS

```bash
bash build_app.sh
# 生成 dist/PPT转图片.app 和 dist/PPT转图片_arm64.dmg
```

首次打开未签名的 `.app`：右键点击 → 打开 → 打开。

### Windows

推送到 `main` 分支后，GitHub Actions 自动打包，在 Actions 页面下载 `PPT转图片_windows_x64.zip`。

或本地打包：
```bash
pip install pyinstaller
pyinstaller --windowed --name "PPT转图片" --noconfirm main.py
```

## 输出结构

```
导出图片/
├── 课件名称A/
│   ├── 1.png
│   ├── 2.png
│   └── ...
├── 课件名称B/
│   ├── 1.png
│   └── ...
└── ...
```

## 文件名清理规则

程序会自动从 PPT 文件名中去除：
- 平台标记（公众号、小红书、抖音、微博等）
- 版权声明（侵权删、转载、版权等）
- @用户名
- 各种括号内容（【】、[]、（））

例如：`精美课件【公众号：XXX】侵删.pptx` → `精美课件/`
