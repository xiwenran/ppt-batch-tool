---
name: ppt-batch-tool
description: PPT批量转图片：把文件夹中的PPT/PPTX文件批量导出为PNG图片，自动清理文件名。触发词：PPT转图片、把PPT转成图片、批量导出PPT、PPT转PNG、把课件转成图片、导出幻灯片截图。
---

# PPT 批量转图片 Skill

将指定文件夹下的所有 PPT/PPTX 文件递归扫描并导出为 PNG 图片。
使用 Microsoft PowerPoint（macOS）或 LibreOffice 作为转换引擎。

## CLI 路径

```
python3 ~/ppt-batch-tool/cli.py <子命令>
```

## 工作流程

### Step 0：确认参数

用户会说类似：
- "把这个文件夹的 PPT 都转成图片"
- "帮我导出 ~/Downloads/课件/ 里所有 PPT 的前10页"

需要确认两件事：
1. **输入文件夹**：哪个目录（递归扫描，会找到所有子文件夹里的 PPT）
2. **输出目录**：没有指定时默认用 `输入文件夹/../PPT图片/`
3. **最多导出页数**：默认 17 页，用户有特殊需求时调整

### Step 1：检测环境（首次使用或有疑问时）

```bash
cd ~/ppt-batch-tool && python3 cli.py detect
```

确认 PowerPoint 或 LibreOffice 可用再继续。

### Step 2：执行转换

```bash
cd ~/ppt-batch-tool && python3 cli.py convert \
  --input <输入文件夹> \
  --output <输出目录> \
  --max-slides <页数>
```

**注意**：macOS 上首次运行会弹出 PowerPoint 授权窗口，告诉用户点「允许」即可，之后本次批量不再弹。

### Step 3：报告结果

完成后告诉用户：
- 成功/失败了多少个文件
- 输出目录在哪里
- 如有失败，说明哪个文件出了什么问题

## 输出结构

```
输出目录/
  课件名称A/     ← 文件名已自动清理（去掉版权声明、@用户名、平台标记等）
    1.png
    2.png
    ...
  课件名称B/
    1.png
    ...
```

## 注意事项

- 支持格式：`.ppt` `.pptx` `.pps` `.ppsx`
- 自动跳过 Office 临时文件（`~$` 开头）
- 文件名清理会自动去掉：`【公众号：XXX】`、`@用户名`、`（转载）` 等常见标记
- 如果输出文件夹已有同名子目录，自动追加 `_2`、`_3` 避免覆盖
