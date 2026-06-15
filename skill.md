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
python3 thesis_formatter.py "input.docx" ["output.docx"] [--doc-title="标题"] [--no-toc] [--mode={xsyu_thesis|homework}] [--course-title="课程名"]
```

Options:
- `--doc-title="XXX"` — sets header text (default: "论文")
- `--no-toc` — skip TOC generation message
- `--mode=xsyu_thesis|homework` — formatting mode (default: xsyu_thesis)
- `--course-title="XXX"` — course name for header in homework mode

## Formatting Modes

### 🎓 XSYU Thesis Mode (`--mode=xsyu_thesis`, default)
Formal Xi'an Shiyou University bachelor's thesis format:
- **Front matter** (cover + abstract + TOC): Roman numeral page numbers (Ⅰ, Ⅱ, Ⅲ...), no headers
- **Body chapters**: Each chapter in its own section, odd/even page headers
  - Odd pages: chapter title, 宋体五号居中
  - Even pages: "西安石油大学本科毕业设计(论文)", 宋体五号居中
- **Page numbering**: Front matter = Roman; Body = Arabic starting from 1
- **Section breaks**: Before every chapter heading (except 摘要/ABSTRACT)

### 📝 Homework Mode (`--mode=homework`)
Simplified format for daily assignments:
- No section breaks between chapters
- One continuous section
- Header: user-specified course name, 宋体五号居中
- Simple continuous page numbering from page 1
- No odd/even page differentiation

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

**XSYU 毕业论文模式**:
- **每个一级章标题前**插入分节符（除摘要/ABSTRACT外）
- 封面+摘要独立成节（罗马数字页码）
- 参考文献/致谢/附录各自成节
- 每个正文节：奇偶页不同页眉
- 章节之间不插入分页符（自然连续流动）

**日常大作业模式**:
- 不插入任何分节符，全文统一为一节
- 单一页眉（课程名称），连续页码

**分页符 (Page Breaks)** — 已取消:
- 章节之间不再插入分页符 — 章节自然连续流动，页面充实

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

**XSYU 毕业论文模式**:
- 封面+摘要节: 无页眉，罗马数字页码（Ⅰ,Ⅱ,Ⅲ...），宋体五号居中
- 正文各章独立节:
  - 奇数页页眉: 章节标题 (宋体, 五号 10.5pt, 居中)
  - 偶数页页眉: "西安石油大学本科毕业设计(论文)" (宋体, 五号 10.5pt, 居中)
  - 页脚: PAGE域 (阿拉伯数字, 宋体五号, 居中)
- 所有节均取消"链接到上一节"

**日常大作业模式**:
- 全部节: 页眉=课程名称 (宋体, 五号 10.5pt, 居中)
- 页脚: PAGE域 (阿拉伯数字, 宋体五号, 居中, 从1开始)

### 7. Auto-Detection

**Heading 1** (一级标题):
- 第一章 / 第二章 / 第1章... (第X章格式) 🆕
- 一、二、三... / 十一、二十一、 (中文数字编号) → **自动转为 1. 2. 3. (阿拉伯数字)**
- 一 (纯中文数字，后跟空格) 🆕
- 1. / 2. / 3. ... (阿拉伯数字编号，与 1.1 区分)
- 1 (纯数字，后跟空格) 🆕
- 摘要 / ABSTRACT / 参考文献 / 致谢 / 谢辞 / 绪论 / 引言 / 前言 / 结论 / 目录 / 鸣谢 / 附录X

**Heading 2** (二级标题):
- 第一节 / 第二节 / 第1节... (第X节格式) 🆕
- （一）（二）（三）... (中文括号编号)
- 1.1 / 2.3 (数字编号，需有空格后缀)

**Heading 3** (三级标题):
- 1.1.1 / 2.3.4 (数字编号，需有空格后缀)
- ~~（1）（2）~~ — v3.0 已移除（避免正文枚举被误识别）

**Formatting-based fallback** 🆕:
- 当文本匹配失败时，自动检测段落字体格式：
  - 居中 + 加粗 + 字号 ≥ 15pt → Heading 1
  - 加粗 + 字号 ≥ 14pt → Heading 2
  - 加粗 + 字号 ≥ 12pt → Heading 3

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
- 两种排版模式可选：毕业论文模式（分节+奇偶页眉+罗马页码）vs 大作业模式（简单连续）
- Cover abstract is auto-split if combined with heading in one paragraph
- Supports standard (1.1/1.1.1) and Chinese (（一）/第一章/第一节) heading formats
- Formatting-based fallback detects headings by font size/bold/alignment when regex fails
- （1）（2）no longer treated as H3
- Chapter heading auto-detection now includes: 第一章/第X章, 第一节/第X节, bare numerals

## Version History

### v3.2 (2026-06-15)
- **Section structure overhaul**:
  - XSYU mode: Abstract section NO page numbers; TOC section gets Roman numerals (Ⅰ,Ⅱ,Ⅲ...) + odd/even headers ("目录"/"西安石油大学本科毕业设计(论文)"); Body chapters restart Arabic at 1
  - Homework mode: 2 section breaks (after abstract, before references); Abstract section NO page numbers
- **TOC placeholder**: Auto-inserts blank paragraph after keywords in XSYU mode for TOC insertion area
- **Page numbering fix**: Uses flag-based restart logic instead of hardcoded section index
- **Section plan alignment fix**: No phantom TOC entry when placeholder is absent

### v3.1 (2026-06-15)
- **New heading patterns**: 第一章/第X章 (H1), 第一节/第X节 (H2), bare Chinese/Arabic numerals (H1)
- **Formatting inference fallback**: Detects headings by font size/bold/alignment when regex fails
- **Mode system**: `--mode=xsyu_thesis|homework` with different section/header/footer/page-numbering schemes
- **XSYU mode**: Per-chapter section breaks, odd/even page headers, Roman numeral front matter
- **Homework mode**: Simple continuous formatting, course-title header
- **Enhanced markdown cleanup**: Strips `__bold__` and `~~strike~~` artifacts
- **Special headings expanded**: 目录, 鸣谢, 附录一~五
- **Streamlit UI**: Mode selector with conditional inputs

### v3.0 (initial release)
- Core formatting pipeline, 2 section breaks, H1/H2/H3 detection, cover formatting
