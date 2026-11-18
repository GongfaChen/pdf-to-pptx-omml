#!/usr/bin/env python3
"""audit_layout.py — PPT 版式审计（对 PowerPoint 导出的 PDF 做字形级检查）。

三件套检查（对应 pitfalls.md 坑 23/24/39；坑 27 为输入前提）：
1. 字形级 span 碰撞：两 span 交集面积 / 较小面积 ≥ 0.40 判为真压字
   （block 级会把 display 公式分子/分母拆成相邻块而大量误报；下标与基底 ~0.30 属正常）
2. 内容溢出：非页脚内容最低点 > --bottom-in（I10：默认按页高推导 page.rect.height/72 − 0.52，
   16:9 即 6.98in；旧硬编码 6.98 只对 7.5in 页高成立，4:3/Letter 会误报。允许显式覆盖）
3. 页脚计数：有正文内容的页页脚对象数必须 = --expected-footer（默认 1）；
   标题页/分隔页（--skip-pages 列出的页）不检查页脚
4. 公式渲染核验（I4 / 坑 27）：逐页统计 Cambria Math span 数，命中 [formula]/□ 占位即判违规；
   三件套几何全零并不能证明公式是 PowerPoint 原生渲染

用法：
    python audit_layout.py <导出的.pdf> --footer "页脚文案" [--skip-pages 1,3] \
        [--bottom-in 6.98] [--threshold 0.40] [--report 审计.md]
退出码：0 = 全零通过；1 = 有违规。
"""
import argparse
import sys
from pathlib import Path

try:
    import fitz
except ImportError:
    print("缺少依赖：pip install pymupdf", file=sys.stderr)
    sys.exit(2)

PT_PER_IN = 72.0


def get_spans(page):
    spans = []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for line in b.get("lines", []):
            for s in line.get("spans", []):
                if s["text"].strip():
                    spans.append({
                        "rect": fitz.Rect(s["bbox"]),
                        "text": s["text"],
                        "size": s["size"],
                        "font": s.get("font", ""),  # I4：保留字体，用于识别 Cambria Math 字形
                    })
    return spans


def is_footer(span, footer, max_len=90):
    return bool(footer) and footer in span["text"] and len(span["text"]) <= max_len


def audit_page(page, page_no, args):
    problems = []
    spans = get_spans(page)
    body = [s for s in spans if not is_footer(s, args.footer)]
    footers = [s for s in spans if is_footer(s, args.footer)]

    # 1) 字形碰撞
    for i in range(len(body)):
        for j in range(i + 1, len(body)):
            r = body[i]["rect"] & body[j]["rect"]
            inter = r.get_area()
            smaller = min(body[i]["rect"].get_area(), body[j]["rect"].get_area())
            if smaller > 0 and inter / smaller >= args.threshold:
                problems.append(
                    f"p{page_no} 压字 {inter/smaller:.0%}: {body[i]['text'][:20]!r} × {body[j]['text'][:20]!r}"
                )

    # 2) 溢出（含图片底边）。I10：未显式给 --bottom-in 时按本页页高推导，
    #    避免把 4:3/Letter 页面按 16:9 的 6.98in 误判为溢出。
    page_h_in = page.rect.height / PT_PER_IN
    bottom_in = args.bottom_in if args.bottom_in is not None else page_h_in - 0.52
    limit = bottom_in * PT_PER_IN
    lowest = max((s["rect"].y1 for s in body), default=0)
    for info in page.get_image_info():
        lowest = max(lowest, info["bbox"][3])
    if lowest > limit + 0.5:
        problems.append(f"p{page_no} 内容底边 {lowest/PT_PER_IN:.2f}in > {bottom_in:.2f}in（溢出/压页脚区）")

    # 3) 页脚计数（跳过 title/divider 页；无正文内容的页也跳过）
    skipped = page_no in args.skip_pages
    if not skipped and body:
        if len(footers) != args.expected_footer:
            problems.append(f"p{page_no} 页脚数 {len(footers)} ≠ {args.expected_footer}（双页脚/缺页脚）")

    # 4) 公式渲染核验（I4 / 坑 27）：原生 OMML 走 Cambria Math 字形；
    #    只认 mc:Fallback 的渲染器（如 LibreOffice）整页只剩 [formula]/□，几何三件套照样全零。
    math_count = sum(
        1 for s in spans if "cambriamath" in (s["font"] or "").replace(" ", "").lower()
    )
    placeholder = sum(s["text"].count("[formula]") + s["text"].count("□") for s in body)
    if placeholder:
        problems.append(
            f"p{page_no} 命中占位文本 {placeholder} 处（[formula]/□）—— PDF 非 PowerPoint 原生渲染"
        )
    return problems, lowest, math_count


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="PowerPoint 导出的 PDF（pres.SaveAs(pdf, 32)）")
    ap.add_argument("--footer", default="", help="页脚文案（用于区分页脚与正文）")
    ap.add_argument("--skip-pages", default="", help="不检查页脚的页码（1-based，逗号分隔，如 1,5）")
    ap.add_argument("--bottom-in", type=float, default=None,
                    help="内容底边上限（英寸，默认按页高推导 page.rect.height/72 − 0.52；显式传值覆盖）")
    ap.add_argument("--threshold", type=float, default=0.40, help="碰撞判定阈值（默认 0.40）")
    ap.add_argument("--expected-footer", type=int, default=1)
    ap.add_argument("--report")
    args = ap.parse_args()
    args.skip_pages = {int(x) for x in args.skip_pages.split(",") if x.strip()}

    doc = fitz.open(args.pdf)
    all_problems = []
    math_counts = {}
    for i, page in enumerate(doc, start=1):
        problems, _, math_count = audit_page(page, i, args)
        all_problems += problems
        math_counts[i] = math_count

    lines = [f"# 版式审计：{Path(args.pdf).name}", ""]
    if args.bottom_in is None:
        h = doc[0].rect.height / PT_PER_IN
        lines += [f"> --bottom-in 未显式指定，按页高推导：{h:.2f}in − 0.52 = **{h - 0.52:.2f}in**"
                  f"（16:9 页高 7.5in 时 = 6.98in；4:3/Letter 自动放宽）。", ""]
    if all_problems:
        lines += [f"**{len(all_problems)} 项违规：**", ""] + [f"- {p}" for p in all_problems]
    else:
        lines += [f"collisions=0 / overflow=0 / footer={args.expected_footer} —— 三件套全零，通过（公式渲染核验见下）。"]
    lines += ["", "## 公式渲染核验（Cambria Math span 计数）", ""]
    lines += [f"- p{pno}：CambriaMath {c} 处" for pno, c in math_counts.items()]
    total_math = sum(math_counts.values())
    if total_math == 0:
        lines += [
            "",
            "> **公式渲染未验收**：全篇 0 个 CambriaMath span，无法证明公式是 PowerPoint 原生渲染；"
            "确认输入 PDF 来自 PowerPoint COM SaveAs，勿把只认 mc:Fallback 的渲染器导出物当交付物。",
        ]
    else:
        lines += ["", f"> 全篇 CambriaMath {total_math} 处；[formula]/□ 命中即判违规（坑 27/33）。"]
    report = "\n".join(lines)
    print(report)
    if args.report:
        Path(args.report).write_text(report, encoding="utf-8")
    return 0 if not all_problems else 1


if __name__ == "__main__":
    sys.exit(main())
