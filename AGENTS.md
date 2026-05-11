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

---

## ⚡ Echo 会话检查清单（每次任务完成前必查）

完成任何实质性改动后，**在报告完成之前**，按顺序检查：

**① 高风险改动？→ 冷眼审查**
- 触发条件：发布逻辑 / 写入文件或数据库 / 防重复机制
- 操作：说"我将触发冷眼审查"，派 worker 读取 echo-reviewer.md 做审查，或使用 App Review

**② 方向性问题出现？→ 主动建议圆桌讨论**
- 触发条件：项目定位、产品路线、架构主线；用户表达"要不要 / 像 A 还是像 B / 感觉但不确定"；或主会话自己准备列出 ≥2 个互斥方向
- 操作：在回答之前先停下，说"这是方向性问题，建议三省讨论。可以输入「圆桌讨论：[问题]」触发多 Agent 分析。"
- 注意：主动建议圆桌不等于擅自执行圆桌。用户确认后才执行完整 roundtable。

**③ 有实质改动？→ Obsidian 捕获**
- 触发条件：新功能完成 / 重构 / bug 根因找到 / 推送 GitHub
- 操作：展示 changelog 草稿，等用户确认后写入 `~/Obsidian/PersonalWiki/项目/ppt-batch-tool/changelog/`

**④ git push 前 → 脱敏扫描**
- 检查：`/Users/用户名/`绝对路径、token、邮箱、AppID
- 发现敏感内容：立即停下告知用户，修改后再推
