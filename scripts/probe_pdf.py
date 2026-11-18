#!/usr/bin/env python3
"""probe_pdf.py — 免费诊断 PDF 教材的可提取性（只读，不修改任何文件）。

诊断内容：
- 每页有无文本层（原生文字 / 扫描版 / OCR 可疑）
- 嵌入位图与矢量图形数量（图/表估算）
- Figure / Table caption 数量
- 数学符号密度（估算 OMML 转录工作量）
- 单栏 / 双栏版式探测
- 扫描页边缘装订线（彩色竖线）探测

用法：
    python probe_pdf.py <pdf路径...> [--report 报告.md] [--json 输出.json]

退出码恒为 0（probe 免费且只读）。
"""
import argparse
import json
import re
import sys
from pathlib import Path

try:
    import fitz  # pymupdf
except ImportError:
    print("缺少依赖：pip install pymupdf", file=sys.stderr)
    sys.exit(2)

MATH_CHARS = set(
    "∑∫∏√±×÷≤≥≠≈∞∂∇°′″"
    "αβγδεζηθικλμνξπρστυφχψω"
    "ΓΔΘΛΞΠΣΥΦΨΩ"
    "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻ⁿ"
    "₀₁₂₃₄₅₆₇₈₉₊₋"
)

CAPTION_RE = re.compile(r"^\s*(Figure|Fig\.?|Table|Example|Problem)\s*\d+", re.IGNORECASE)

# ---- 单页分析 ---------------------------------------------------------------


def page_text_stats(page):
    text = page.get_text()
    stripped = text.strip()
    n = len(stripped)
    words = re.findall(r"[A-Za-z]{2,}", stripped)
    tokens = stripped.split()
    good_ratio = (len(words) / len(tokens)) if tokens else 0.0
    math_count = sum(1 for ch in stripped if ch in MATH_CHARS)
    math_density = math_count / n * 1000 if n else 0.0
    return stripped, n, good_ratio, math_count, math_density


def detect_columns(page):
    blocks = [b for b in page.get_text("blocks") if b[4].strip()]
    if len(blocks) < 6:
        return "?"
    mid = page.rect.width / 2
    left = sum(1 for b in blocks if b[0] < mid - 20)
    right = sum(1 for b in blocks if b[0] > mid + 20)
    if left >= 5 and right >= 5:
        return "2栏"
    return "1栏"


def detect_binding_lines(page):
    """渲染低分辨率后扫描左右边缘的饱和色竖线（装订线/印章边）。"""
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), alpha=False)
    except Exception:
        return False
    w, h, n = pix.width, pix.height, pix.n
    s = pix.samples
    stripe_h = 0

    def sat_col(x):
        cnt = 0
        for y in range(h):
            o = (y * w + x) * n
            r, g, b = s[o], s[o + 1], s[o + 2]
            if (r > 150 and g < 110 and b < 110) or (b > 150 and r < 110 and g < 110) or (g > 150 and r < 110 and b < 110):
                cnt += 1
        return cnt

    for x in list(range(0, 12)) + list(range(w - 12, w)):
        stripe_h = max(stripe_h, sat_col(x))
    return stripe_h > h / 3


def analyze_page(doc, idx):
    page = doc[idx]
    text, n, good_ratio, math_count, math_density = page_text_stats(page)
    images = len(page.get_image_info(xrefs=True))
    captions = sum(1 for line in text.splitlines() if CAPTION_RE.match(line))

    if n < 20:
        layer = "扫描版" if images or page.get_drawings() else "空白页"
    elif good_ratio < 0.35 and n > 50:
        layer = "OCR可疑"  # 有文本层但几乎拼不出单词 → OCR 错字/乱码
    else:
        layer = "文字层"

    return {
        "page": idx + 1,
        "layer": layer,
        "chars": n,
        "images": images,
        "vector_drawings": len(page.get_drawings()),
        "captions": captions,
        "math_chars": math_count,
        "math_density": round(math_density, 1),
        "columns": detect_columns(page) if layer != "扫描版" else "-",
        "binding_line": detect_binding_lines(page) if layer == "扫描版" else False,
    }


# ---- 汇总判定 ---------------------------------------------------------------


def summarize(pdf_path, pages):
    scan_pages = [p["page"] for p in pages if p["layer"] == "扫描版"]
    ocr_pages = [p["page"] for p in pages if p["layer"] == "OCR可疑"]
    text_pages = [p["page"] for p in pages if p["layer"] == "文字层"]
    total_fig = sum(p["images"] for p in pages)
    total_captions = sum(p["captions"] for p in pages)
    dens = [p["math_density"] for p in pages if p["layer"] == "文字层"]
    avg_dens = sum(dens) / len(dens) if dens else 0.0
    cols = {}
    for p in pages:
        if p["columns"] not in ("?", "-"):
            cols[p["columns"]] = cols.get(p["columns"], 0) + 1
    main_col = max(cols, key=cols.get) if cols else "?"

    if scan_pages:
        pipeline = "扫描版管线（无文本层：按页码渲染 + 视觉定位 + OCR/目视转录公式，图片按像素裁剪）"
        difficulty = "高（公式须人工目视转录为 OMML）"
    elif ocr_pages and len(ocr_pages) > len(text_pages) / 2:
        pipeline = "OCR 修正管线（文本层错字多：正文以文本层为准，公式以渲染页目视为准）"
        difficulty = "高"
    else:
        pipeline = "原生文本层管线（文本层可用，caption 定位 + clip 渲染裁图）"
        difficulty = "中" if avg_dens > 8 else "低"

    return {
        "file": str(pdf_path),
        "total_pages": len(pages),
        "text_layer_pages": text_pages,
        "scan_pages": scan_pages,
        "ocr_suspect_pages": ocr_pages,
        "embedded_images": total_fig,
        "captions_detected": total_captions,
        "avg_math_density_per_kchars": round(avg_dens, 1),
        "dominant_layout": main_col,
        "binding_line_pages": [p["page"] for p in pages if p["binding_line"]],
        "recommended_pipeline": pipeline,
        "omml_difficulty": difficulty,
    }


def render_markdown(pdf_path, pages, summary):
    lines = [
        f"# PDF 可提取性诊断报告：{Path(pdf_path).name}",
        "",
        f"- 页数 **{summary['total_pages']}**｜文字层 {len(summary['text_layer_pages'])} 页｜"
        f"扫描版 {len(summary['scan_pages'])} 页｜OCR 可疑 {len(summary['ocr_suspect_pages'])} 页",
        f"- 嵌入位图 **{summary['embedded_images']}** 个｜caption 命中 **{summary['captions_detected']}** 个",
        f"- 平均数学符号密度 **{summary['avg_math_density_per_kchars']}** /千字符｜主导版式 **{summary['dominant_layout']}**",
        f"- 推荐管线：**{summary['recommended_pipeline']}**",
        f"- OMML 公式转录难度：**{summary['omml_difficulty']}**",
        "",
        "## 逐页明细",
        "",
        "| 页 | 类型 | 字数 | 位图 | caption | 数学字符/千字 | 版式 |",
        "|---:|---|---:|---:|---:|---:|---|",
    ]
    for p in pages:
        lines.append(
            f"| {p['page']} | {p['layer']} | {p['chars']} | {p['images']} "
            f"| {p['captions']} | {p['math_density']} | {p['columns']} |"
        )
    if summary["scan_pages"]:
        lines += ["", f"> 扫描版页（无文本层，公式只能目视转录）：{summary['scan_pages']}"]
    if summary["ocr_suspect_pages"]:
        lines += ["", f"> OCR 可疑页（文本层错字风险，公式勿信 get_text）：{summary['ocr_suspect_pages']}"]
    if summary["binding_line_pages"]:
        lines += ["", f"> 检测到装订线/边缘色条的页（裁图时需避开）：{summary['binding_line_pages']}"]
    lines += [
        "",
        "## 诚实边界",
        "",
        "- 本报告为只读诊断，未提取任何正文；提不出来的页不会被猜。",
        "- 扫描版页的公式识别精度取决于扫描清晰度，OMML 转录后需人工核对。",
        "- 表格结构不会自动还原，仅保证阅读顺序。",
    ]
    return "\n".join(lines)


def main():
    from license_check import require_license_or_exit  # probe 永远免费（FREE_FEATURES）
    require_license_or_exit("probe")

    ap = argparse.ArgumentParser(description="PDF 教材可提取性诊断（免费、只读）")
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("--report", help="输出 Markdown 报告路径（单文件时）")
    ap.add_argument("--json", help="输出 JSON 摘要路径")
    args = ap.parse_args()

    all_summaries = []
    for pdf in args.pdfs:
        doc = fitz.open(pdf)
        pages = [analyze_page(doc, i) for i in range(len(doc))]
        summary = summarize(pdf, pages)
        all_summaries.append(summary)
        print(render_markdown(pdf, pages, summary))
        print("\n" + "-" * 60 + "\n")
        if args.report and len(args.pdfs) == 1:
            Path(args.report).write_text(render_markdown(pdf, pages, summary), encoding="utf-8")
        doc.close()

    if args.json:
        Path(args.json).write_text(
            json.dumps(all_summaries, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
