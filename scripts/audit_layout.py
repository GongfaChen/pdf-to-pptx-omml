#!/usr/bin/env python3
"""audit_layout.py — PPT 版式审计（对 PowerPoint 导出的 PDF 做字形级检查）。

三件套检查（对应 pitfalls.md 坑 23/24/27/39）：
1. 字形级 span 碰撞：两 span 交集面积 / 较小面积 ≥ 0.40 判为真压字
   （block 级会把 display 公式分子/分母拆成相邻块而大量误报；下标与基底 ~0.30 属正常）
2. 内容溢出：非页脚内容最低点 > --bottom-in（默认 6.98in，页脚在 7.08in）
3. 页脚计数：有正文内容的页页脚对象数必须 = --expected-footer（默认 1）；
   标题页/分隔页（--skip-pages 列出的页）不检查页脚

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
                    spans.append({"rect": fitz.Rect(s["bbox"]), "text": s["text"], "size": s["size"]})
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

    # 2) 溢出（含图片底边）
    limit = args.bottom_in * PT_PER_IN
    lowest = max((s["rect"].y1 for s in body), default=0)
    for info in page.get_image_info():
        lowest = max(lowest, info["bbox"][3])
    if lowest > limit + 0.5:
        problems.append(f"p{page_no} 内容底边 {lowest/PT_PER_IN:.2f}in > {args.bottom_in}in（溢出/压页脚区）")

    # 3) 页脚计数（跳过 title/divider 页；无正文内容的页也跳过）
    skipped = page_no in args.skip_pages
    if not skipped and body:
        if len(footers) != args.expected_footer:
            problems.append(f"p{page_no} 页脚数 {len(footers)} ≠ {args.expected_footer}（双页脚/缺页脚）")
    return problems, lowest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="PowerPoint 导出的 PDF（pres.SaveAs(pdf, 32)）")
    ap.add_argument("--footer", default="", help="页脚文案（用于区分页脚与正文）")
    ap.add_argument("--skip-pages", default="", help="不检查页脚的页码（1-based，逗号分隔，如 1,5）")
    ap.add_argument("--bottom-in", type=float, default=6.98, help="内容底边上限（英寸，默认 6.98）")
    ap.add_argument("--threshold", type=float, default=0.40, help="碰撞判定阈值（默认 0.40）")
    ap.add_argument("--expected-footer", type=int, default=1)
    ap.add_argument("--report")
    args = ap.parse_args()
    args.skip_pages = {int(x) for x in args.skip_pages.split(",") if x.strip()}

    doc = fitz.open(args.pdf)
    all_problems = []
    for i, page in enumerate(doc, start=1):
        problems, _ = audit_page(page, i, args)
        all_problems += problems

    lines = [f"# 版式审计：{Path(args.pdf).name}", ""]
    if all_problems:
        lines += [f"**{len(all_problems)} 项违规：**", ""] + [f"- {p}" for p in all_problems]
    else:
        lines += ["collisions=0 / overflow=0 / footer=1 —— 三件套全零，通过。"]
    report = "\n".join(lines)
    print(report)
    if args.report:
        Path(args.report).write_text(report, encoding="utf-8")
    return 0 if not all_problems else 1


if __name__ == "__main__":
    sys.exit(main())
