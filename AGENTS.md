# PPT批量工具 — Codex 项目说明

## 项目概述

PPT 批量导出图片工具：把文件夹中的 PPT 文件批量导出为 PNG 图片，支持递归扫描子文件夹。

## 重要约束

- 支持 `.ppt` `.pptx` `.pps` `.ppsx` 格式，每个 PPT 默认导出前 17 页
- 转换引擎自动检测：macOS 优先 PowerPoint → LibreOffice；Windows 优先 WPS → PowerPoint → LibreOffice
- CLI 入口：`python main.py`
- 打包：macOS 用 `bash build_app.sh`；Windows 走 GitHub Actions
- 输出结构：`导出图片/{课件名称}/1.png, 2.png, ...`
- 文件名自动清理（去除版权声明、平台标记等）

**本地路径**：`~/ppt-batch-tool/`（GitHub 仓库名 `xiwenran/ppt-batch-tool`，与本地目录名一致）
**主要流程入口**：`python main.py`（批量导出图片 CLI）

---

## 项目坐标（AI 找本项目信息的固定入口）

Obsidian 里本项目名为**ppt-batch-tool**。找本项目的方案/进度/风险 → 直接读这几个路径，不用扫全库：

| 类别 | 路径 |
|---|---|
| Obsidian 项目主页 | `~/Obsidian/PersonalWiki/项目/ppt-batch-tool/README.md` |
| Obsidian changelog / 进度 | `~/Obsidian/PersonalWiki/项目/ppt-batch-tool/changelog/`（按日期命名，取目录列表最新一个） |
| Obsidian 踩坑记录 | `~/Obsidian/PersonalWiki/项目/ppt-batch-tool/踩坑记录.md` |

> 坐标卡只存**常青入口的固定路径**；具体内容会变，查最新进度需要实际读文件。

---

## 项目专属护栏（全局规则外的补充）

- 暂无。
- 通用护栏（冷眼审查 / 圆桌 / Obsidian 捕获 / 脱敏 / 规则同步）均以全局 AGENTS.md 为准，本文件不再重抄，源头改则处处改。
