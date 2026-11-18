---
name: pdf-to-pptx-omml
description: 从 PDF 教材/讲义提取章节内容生成专业 PPTX（正文 20pt/行距 1.4/Times New Roman，公式转 OMML 在 PowerPoint 中完全可编辑，图表按 caption 自动定位裁剪 300dpi 原图）。适用 AoPS/IB/Cambridge/国内理工教材等 → 课件转换。触发词：教材提取 PPT、OMML 公式、章节转课件、PDF to PPTX、课件制作。先跑 probe 免费诊断，再按推荐管线执行。
---

# PDF 教材 → PPT（OMML 可编辑公式）工作流

环境：Windows + Python（pymupdf / python-pptx / lxml / pillow / pdfplumber / pywin32）+ 本机 PowerPoint 或 WPS。
所有脚本在 `scripts/`，实战坑总表在 `references/pitfalls.md`（执行前必读对应类别）。

## 流程总览

```
probe（诊断）→ 定位章节 → 裁图 → 构建 PPTX → 静态校验 → COM 导出 PDF → 版式审计 → 交付
```

## 第 0 步 · probe 诊断（只读，先判管线）

```bash
python scripts/probe_pdf.py <教材.pdf> --report 诊断.md
```

报告给出：每页有无文本层（文字层/扫描版/OCR 可疑）、嵌入位图数、caption 命中数、
数学符号密度、单/双栏、装订线页。据此选管线：

| 诊断结果 | 管线 |
|---|---|
| 全书文字层 | 原生管线：caption 文本块定位 + clip 渲染裁图 |
| 大量扫描页 | 扫描管线：按页码渲染逐页视觉确认，公式目视转录 |
| OCR 可疑页多 | 混合管线：正文用文本层，**公式一律以渲染页目视为准**（坑 40） |

## 第 1 步 · 定位章节页码

- `pymupdf` 搜索各节标题 + 下一章标题确定页码范围，dump 全文通读。
- **注意 `p.number` 是 0-based**（坑 9）；核对用页面文本匹配，不靠记忆页码。
- 扫描版无文本层：按用户给的页码渲染页面（1-based）视觉确认。

## 第 2 步 · 裁剪图表

- 优先判断图是嵌入位图还是矢量：位图用 `get_image_info(xrefs=True)` + xref 原分辨率提取（坑 16）；
  矢量才用 caption 定位 + clip 渲染（`Matrix(3,3)`，300dpi PNG）。
- caption 匹配规则与 prose 判别三判据见坑 28/41；双栏版式按中缝分左右栏定裁剪框。
- 扫描页先查装订线色条（坑 42，probe 已标出）。PIL 拼 contact sheet 逐张视觉校验，混入正文的重裁。

## 第 3 步 · 构建 PPTX（omml_helpers）

用 `scripts/omml_helpers.py` 构造公式，**不要手写 OMML XML**——库已把坑 1/3/6/10/11/12/35 在 API 层堵死
（裸 str 自动包 T()、空基底抛错、a14 命名空间自动声明、无 m:acc/supHide）。用法见 `references/omml-cookbook.md`。

规格：16:9（13.333×7.5in）；正文 Times New Roman 20pt / line_spacing=1.4；
display 公式**独占段落且行距 1.0**（坑 13/15）；浅色 shape 里公式必须显式设深色（坑 18）。

行高预算（坑 24）：正文 ≈33.9pt/行；无分数 display ≈1.0 行，含分数 ≈1.7 行；
每行字符数 ≈ 宽度(in)×72/8.7。溢出修法优先级：合并公式行 > 精简文案 > 拆页；**禁止全局缩字号**。

## 第 4 步 · 校验闭环（三道闸，缺一不可）

```bash
# 闸 1：静态校验（不打开 PowerPoint，先扫低级错误）
python scripts/validate_omml.py check <成品>.pptx
python scripts/validate_omml.py lint build_ppt.py      # 构建脚本源码自伤扫描

# 闸 2：COM 导出 PDF + 字形级审计
#   PowerPoint: pres.SaveAs(pdf, 32)（公式原生渲染，一个 PDF 同时核验版式+公式）
#   失败备选：WPS COM DispatchEx('KWPP.Application') Slides(i).Export(...)（导出路径纯 ASCII）
python scripts/audit_layout.py <导出的.pdf> --footer "<页脚文案>" --skip-pages 1 --report 审计.md

# 闸 3：contact sheet 目视（审美问题：对齐/留白/图位置）
```

审计目标：collisions=0 / overflow=0 / footer=1（跳过页除外），迭代到全零。

## 交付物

```
<项目目录>/
  诊断.md               # probe 报告
  ch<N>_text.txt        # 章节全文
  crop_figures.py       # 裁图脚本
  build_ppt.py          # PPT 构建脚本（omml_helpers）
  assets/fig*.png       # 300dpi 裁剪图
  preview/              # contact sheet
  <名称>.pptx           # 最终交付
```

## 诚实边界

- 扫描版页的公式为目视转录，需人工核对；OCR 文本层的公式串一律不信（坑 40）。
- 表格结构不自动还原，仅保证阅读顺序；加密 PDF 不处理。
- Fallback 占位文本仅在极旧渲染器显示——交付前必须走完闸 2 的 PowerPoint 原生渲染核验。
- 环境依赖本机 PowerPoint（或 WPS）COM；Linux/macOS 无 COM 时跳过闸 2，用 LibreOffice 渲染粗核。
