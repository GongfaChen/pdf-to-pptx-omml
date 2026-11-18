---
name: pdf-to-pptx-omml
description: 从 PDF 教材/讲义提取章节内容生成专业 PPTX（正文 20pt/行距 1.4/Times New Roman，公式转 OMML 在 PowerPoint 中完全可编辑，图表按 caption 自动定位裁剪 300dpi 原图）。适用 AoPS/IB/Cambridge/国内理工教材等 → 课件转换。触发词：教材提取 PPT、OMML 公式、章节转课件、PDF to PPTX、课件制作。先跑 probe 免费诊断，再按推荐管线执行。
---

# PDF 教材 → PPT（OMML 可编辑公式）工作流

环境：Windows + Python（pymupdf / python-pptx / lxml / pillow / pdfplumber / pywin32）+ 本机 PowerPoint 或 WPS。
所有脚本在 `scripts/`，实战坑总表在 `references/pitfalls.md`（执行前必读对应类别）。

## 流程总览

```
前置 A（环境闸）→ probe（诊断）→ 定位章节 → 裁图 → 构建 PPTX → 静态校验 → COM 导出 PDF → 版式审计 → 交付
```

## 前置 A · 环境闸（scripts/check_env.py）

开工先跑只读环境闸，确认依赖矩阵与 COM 闸能不能跑（不打开 PowerPoint、不导出、不修改任何文件）：

```bash
python scripts/check_env.py
```

输出：7 项依赖（fitz/pptx/lxml/PIL/pdfplumber/win32com/pythoncom）的 OK/MISSING 矩阵与缺失项安装命令、
PowerPoint/WPS COM ProgID 是否注册，以及闸 2（公式原生渲染核验）的结论。命名「前置 A」是为了不与下面的「第 0 步 · probe 诊断」撞号。

- **有 COM**（pywin32 + PowerPoint/WPS 均可用）：闸 2 可执行，交付前必须 `pres.SaveAs(pdf, 32)` 并核验 CambriaMath 字形 / 无 `[formula]` 占位。
- **无 COM**：闸 2 不可执行 → **公式渲染未验收**，不得据此宣布通过或交付；LibreOffice 等只认 `mc:Fallback` 的渲染器导出的 PDF 禁止作为交付物（与下方「诚实边界」一致）。

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

直接调用 `scripts/crop_figures.py`（把坑 16/28/41/42 的裁图规则固化成 CLI，只读 PDF，产出 `assets/` + `preview/` + `crop_report.md`）：

```bash
python scripts/crop_figures.py <教材.pdf> --out <目录> [--dpi 300] [--cols 4]
```

规则链（脚本内已实现，无需手写）：

- **xref 位图优先**：`get_image_info(xrefs=True)` 拿 bbox，按 xref 原分辨率提取（坑 16）；同 bbox 另渲染一张 300dpi clip 作对照。
- **caption clip**：矢量/文字页用带标点的严格 caption 匹配（坑 41）定位，clip 渲染 300dpi PNG。
- **prose 三判据**（坑 28）：`prose_ratio ≥ 0.86` + `longest_word ≥ 4` + `y1 ≥ 58`，命中即判为正文并跳过该裁剪（「混入正文的重裁」）；区域无矢量绘图/位图也跳过。
- **双栏**：按中缝分左右栏定裁剪框（坑 28）。
- **装订线**：像素级扫描边缘饱和色竖线，裁剪 x0 移到线右侧（坑 42）；PIL 差分白边 trim 兜底。

拼图用 `scripts/contact_sheet.py`（也可单独跑，或由 `crop_figures.py` 在裁图末尾自动调用），输出 `preview/contact_sheet.png` 逐张视觉校验：

```bash
python scripts/contact_sheet.py <assets目录|图片...> --out preview/contact_sheet.png [--cols 4]
```

## 第 3 步 · 构建 PPTX（omml_helpers）

用 `scripts/omml_helpers.py` 构造公式，**不要手写 OMML XML**——库已把坑 1/3/6/10/11/12/35 在 API 层堵死
（裸 str 自动包 T()、空基底抛错、a14 命名空间自动声明、无 m:acc/supHide）。用法见 `references/omml-cookbook.md`。

- **大运算符 / 方程组**：`SUM(sub, sup, body)`、`PROD(sub, sup, body)`（m:nary，∑/∏ 上下限必须同时给，坑 12）、`EQARR(*rows)`（左对齐方程组 m:eqArr，不套 `m:d`；列向量仍用 `CV`，坑 34）。示例见 cookbook。
- **每份 deck 前调 `reset_run_style()`**：`set_run_style()` 写的是模块级全局，size/color/font 传 `None` 只表示「不修改」而非「清空」，连续构建多份 deck 会串味；重置回出厂默认后再设本 deck 样式。
- **helper 片段拼接一律用 `M(...)`，不要用 `+`**（坑 45）：`_Frag` 是 str 子类，`T('a') + T('b')` 曾退化成普通 str 被下游当裸文本包 `T()` 转义，公式显示成 XML 字面文本；库已重载 `__add__/__radd__` 修好，但 `+` 跨版本不稳且 `validate_omml.py lint` 会报，统一用 `M(...)`。

规格：16:9（13.333×7.5in）；正文 Times New Roman 20pt / line_spacing=1.4；
display 公式**独占段落且行距 1.0**（坑 13/15）；浅色 shape 里公式必须显式设深色（坑 18）。

行高预算（坑 24）：正文 ≈33.9pt/行；无分数 display ≈1.0 行，含分数 ≈1.7 行；
每行字符数 ≈ 宽度(in)×72/8.7。溢出修法优先级：合并公式行 > 精简文案 > 拆页；**禁止全局缩字号**。

- **带上下限 / 多行的「大结构」必须 `display=True` 独占段落**（坑 13/24/34）：`SUM`/`PROD`（m:nary 上下限）、`EQARR`/`CV`/`MX`（多行结构）不要内联 `display=False` 塞进正文行——smoke demo 实测 EQARR 内联导出后被挤压、贴住相邻文字，`collisions=0` 但稿件难看。
  行高预算把这一行按 **1 行** 算（无分数 display ≈1.0 行，坑 24），而大结构实际要 ≈1.3–1.7 行（两行向量 ≈1.3 行、三行 ≈1.7 行，坑 34；含分数 ≈1.7 行，坑 24）→ 垂直溢出/挤压。若确实要行内表达，退化为单个符号（如 `∑` 或 `a_i`），而不是整段 nary/eqArr。

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
- mc:Fallback 在只认 Fallback 的渲染器（如 LibreOffice）里会替换整段内容——交付前必须走完闸 2 的 PowerPoint 原生渲染核验。
- 环境依赖本机 PowerPoint（或 WPS）COM；Linux/macOS 无 COM 时闸 2 无法执行，**公式渲染未验收**，不得据此宣布通过或交付。LibreOffice 等只认 mc:Fallback 的渲染器仅可用于版式粗看，其导出的 PDF 禁止作为交付物。
