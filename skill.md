---
name: thesis-formatter
description: Complete thesis formatting (西安石油大学模板) — page, styles, TOC, headers/footers, captions, references. Use when user says 格式化论文/论文格式/format thesis.
---

# Thesis Formatter — 基于西安石油大学本科毕业设计(论文)模板

## How to Invoke

**Method 1: Natural language (recommended)**
Just say: `帮我格式化这篇论文：E:\path\to\文件.docx`

**Method 2: CLI**
```bash
python3 ~/.claude/skills/thesis-formatter/thesis_formatter.py "input.docx" ["output.docx"] [--doc-title="标题"] [--no-toc]
```

Options:
- `--doc-title="XXX"` — sets header text and document title (default: "论文")
- `--no-toc` — skip TOC generation message

## What It Does (完整功能)

### 1. Page Setup
- A4 paper, margins: left 3.0cm, right 2.5cm, top 2.5cm, bottom 2.5cm
- 封面+摘要页：无页码
- 正文第一章起：页码从 1 开始（通过 pgNumType start="1" 实现）

### 2. Cover Page Formatting 🆕

**论文中文题目**:
- 宋体, 小三 (15pt), 加粗, 居中
- 1.25倍行距, 段前0.5行, 段后1行

**摘要标题** (自动拆分):
- 支持多种分隔形式：冒号（摘要：）、空格（摘要 ）、无分隔符（摘要内容紧跟）
- 自动拆分为: "摘要" — Heading 1, 宋体, 小三 (15pt), 加粗, 居中
- 摘要正文 — Normal 样式，首行缩进2字符
- 摘要与关键词之间自动插入空行

**关键词**:
- 宋体, 小四 (12pt), 顶格左对齐, 首行缩进0
- "关键词：" 字样加粗，关键词之间用分号分隔
- 1.25倍行距

### 3. Section & Page Breaks

**分节符 (Section Breaks)** — v3.0 仅保留2处:
- **摘要后 / 正文第一章前**: 封面+摘要独立成节
- **参考文献前**: 参考文献与正文分节

文档仅3节，每节页眉页脚配置:
- 第0节 (封面+摘要): **无页眉、无页脚、无页码**
- 第1节 (正文1-N章+附录): 页眉=文档标题 + 页脚=页码（**从1开始**）
- 第2节 (参考文献): 页眉=文档标题 + 页脚=页码（连续编号）
- **所有节均取消"链接到上一节"**

**分页符 (Page Breaks)** — v3.0 已取消:
- 章节之间不再插入分页符 — 章节自然连续流动，页面充实
- 仅通过2处分节符做结构性分隔

### 4. Styles

| Style | Font | Size | Bold | Align | Other |
|-------|------|------|------|-------|-------|
| Normal (正文) | 宋体+TNR | 小四 12pt | No | Justify | 首行缩进2字符, 1.25x行距, 段前段后0 |
| Heading 1 (标题1) | 宋体+TNR | 小三 15pt | Yes | Center | 首行0, 段前12pt/段后24pt |
| Heading 2 (标题2) | 宋体+TNR | 四号 14pt | Yes | Left | 首行0, 段前6pt/段后6pt |
| Heading 3 (标题3) | 宋体+TNR | 小四 12pt | Yes | Left | 首行0, 段前6pt/段后6pt |
| List Paragraph | 宋体+TNR | 五号 10.5pt | No | Justify | 无缩进, 1.25x行距 |
| toc 1 / toc 2 / toc 3 | 继承Normal | 12pt | No | Left | toc1行距1.2x; toc2左缩进24pt; toc3左缩进48pt |

### 5. Table of Contents
- 在Word中手动插入: References → Table of Contents → Auto Table
- 插入后右键TOC → "更新域" → "更新整个目录"

### 6. Headers & Footers
- 封面节: 无页眉、无页脚
- 正文节: 页眉=文档标题 (宋体+TNR, 五号 10.5pt, 居中)
- 正文节: 页脚=PAGE域 (TNR, 五号 10.5pt, 居中)
- 所有节均取消"链接到上一节"

### 7. Auto-Detection

**Heading 1** (一级标题):
- 一、二、三... (中文数字编号) → **自动转为 1. 2. 3. (阿拉伯数字)**
- 1. / 2. / 3. ... (阿拉伯数字编号，与 1.1 的区分：数字后不跟第二个数字)
- 摘要 / ABSTRACT / 参考文献 / 致谢 / 绪论 / 引言 / 前言 / 结论 / 附录X

**Heading 2** (二级标题):
- （一）（二）（三）... (中文括号编号)
- 1.1 / 2.3 (数字编号，需有空格后缀)

**Heading 3** (三级标题):
- 1.1.1 / 2.3.4 (数字编号，需有空格后缀)
- ~~（1）（2）~~ — v3.0 已移除（避免正文枚举被误识别）

**Cover page**: 第一个非摘要 Heading 1 之前的所有段落
**References**: "参考文献" 标题之后的段落 → List Paragraph 样式

### 8. Special Elements
- Figure captions (图X.X): 五号 10.5pt, bold, centered, **位于图片下方**
- Figure images: centered, width ~5.2 inches, PNG format
- Table captions (表X-X): 五号 10.5pt, bold, centered
- Table cells: 五号 10.5pt, centered
- Reference entries: List Paragraph, 五号 10.5pt, not bold, no indent

### 9. Markdown Artifact Cleanup 🆕
- 自动清除从 Markdown 转换后残留的 `###`/`##`/`#` 标题前缀
- 自动清除 `**粗体**` 和 `*斜体*` 标记（转为纯文本）
- 在文档加载后立即执行，确保后续标题检测和样式应用不受干扰
- 覆盖范围：全部段落（标题、摘要、正文、参考文献）

### 10. Figure Insertion from HTML 🆕
- 从 Chart.js HTML 文件提取图表数据
- 使用 matplotlib + 宋体 重新绘制为高分辨率 PNG（200dpi）
- 自动插入到 docx 对应位置（替换【图X-X】占位符）
- 图注格式：图片在上，图注在下（五号 10.5pt、加粗、居中）

## Template Source
`E:\wechat下载\...\西安石油大学本科毕业设计(论文)模板.docx`

## Notes
- H1 中文数字标题（一、二...）自动转为阿拉伯数字（1. 2...），其余文字内容不变
- Auto-creates .backup.docx before overwriting
- If Word has the file open, save to a different filename
- 仅2处分节符（摘要后 + 参考文献前），章节间无分页符
- Cover abstract is auto-split if combined with heading in one paragraph
- Supports standard (1.1/1.1.1) and Chinese (（一）) heading formats; （1）（2）no longer treated as H3
