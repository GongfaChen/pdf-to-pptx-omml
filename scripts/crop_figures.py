#!/usr/bin/env python3
"""crop_figures.py — 把坑 16/28/41/42 的裁图规则固化成 CLI（只读 PDF，产出 assets + 报告）。

规则链：
  1. 位图优先：get_image_info(xrefs=True) 拿 bbox，按 xref 原分辨率提取（坑 16）；
     同一 bbox 另渲染一张 --dpi（默认 300）的 clip 作对照，演示「原图比 clip 干净」。
  2. 矢量/文字页：caption 严格匹配（坑 41）定位；双栏按中缝分左右栏（坑 28）定裁剪框。
  3. prose 三判据（坑 28）：prose_ratio ≥ 0.86 + longest_word ≥ 4 + y1 ≥ 58；
     命中即判为正文，跳过该裁剪（「混入正文的重裁」）；区域无矢量绘图/位图也跳过。
  4. 装订线像素扫描（坑 42）：低分辨率渲染后扫边缘饱和色竖线，裁剪 x0 移到线右侧。
  5. 白边 trim（坑 28 兜底）：PIL 差分裁掉纯白边。
  6. 末尾调 contact_sheet.build_contact_sheet() 拼图（坑 16/28 的目视闭环）。

用法：
    python scripts/crop_figures.py <教材.pdf> --out <目录> [--dpi 300] [--cols 4]
退出码恒为 0（诊断式产出，缺图不报错）。
"""
import argparse
import io
import re
import sys
from pathlib import Path

try:
    import fitz
except ImportError:
    print("缺少依赖：pip install pymupdf", file=sys.stderr)
    sys.exit(2)
from PIL import Image, ImageChops

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from contact_sheet import build_contact_sheet
except ImportError:
    build_contact_sheet = None

# 坑 41：带标点的严格 caption（排除正文引用 "Table 1.2 shows ..."）。
CAPTION_RE = re.compile(
    r"^\s*(?:Figure|Fig\.?|Table|Example|Problem)\s+\d+(?:\.\d+)?\s*[:.](?:\s|$)",
    re.IGNORECASE,
)
PROSE_RATIO_MIN = 0.86
LONGEST_WORD_MIN = 4
PROSE_Y1_MIN = 58
WRAP_PAD = 6
TRIM_TOL = 18  # 与纯白差异 >18 才算内容（容忍 JPEG/扫描噪点）


def prose_metrics(text):
    words = re.findall(r"[A-Za-z]{2,}", text)
    tokens = text.split()
    ratio = len(words) / len(tokens) if tokens else 0.0
    longest = max((len(w) for w in words), default=0)
    return ratio, longest


def detect_columns(page):
    """双栏按中缝分栏；返回 (栏数, 各栏 x 区间)。"""
    blocks = [b for b in page.get_text("blocks") if b[4].strip()]
    mid = page.rect.width / 2
    if len(blocks) >= 6:
        left = sum(1 for b in blocks if b[0] < mid - 20)
        right = sum(1 for b in blocks if b[0] > mid + 20)
        if left >= 5 and right >= 5:
            return 2, [(0, mid), (mid, page.rect.width)]
    return 1, [(0, page.rect.width)]


def scan_binding_lines(page):
    """坑 42：低分辨率渲染后扫左右边缘饱和色竖线，返回需避开的 x0（PDF 坐标，0=无线）。"""
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), alpha=False)
    except Exception:
        return 0.0
    w, h, n = pix.width, pix.height, pix.n
    s = pix.samples

    def sat_col(x):
        cnt = 0
        for y in range(h):
            o = (y * w + x) * n
            r, g, b = s[o], s[o + 1], s[o + 2]
            if ((r > 150 and g < 110 and b < 110) or
                    (b > 150 and r < 110 and g < 110) or
                    (g > 150 and r < 110 and b < 110)):
                cnt += 1
        return cnt

    hit = [x for x in list(range(0, 12)) + list(range(w - 12, w)) if sat_col(x) > h / 3]
    if not hit:
        return 0.0
    return (max(hit) + 2) * 2.0  # 0.5 倍渲染 → ×2 还原到 PDF 坐标


def trim_white(im):
    """坑 28 兜底：PIL 差分裁掉纯白边，四周各留 WRAP_PAD 像素。"""
    rgb = im.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L").point(lambda v: 255 if v > TRIM_TOL else 0)
    bbox = diff.getbbox()
    if not bbox:
        return rgb
    l, t, r, b = bbox
    return rgb.crop((max(0, l - WRAP_PAD), max(0, t - WRAP_PAD),
                     min(rgb.width, r + WRAP_PAD), min(rgb.height, b + WRAP_PAD)))


def render_clip(page, rect, dpi):
    """按 rect 渲染 --dpi（默认 300dpi）PNG，返回 PIL Image。"""
    rect = fitz.Rect(rect) & page.rect
    if rect.is_empty:
        raise ValueError("裁剪框为空")
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect, alpha=False)
    return pix.pil_image().convert("RGB")


def fmt_rect(r):
    return "(" + ",".join(f"{v:.0f}" for v in (r.x0, r.y0, r.x1, r.y1)) + ")"


def column_of(x, col_ranges):
    for lo, hi in col_ranges:
        if lo <= x < hi:
            return lo, hi
    return col_ranges[-1]


def clip_in_column(bbox, col_ranges, bind_x):
    lo, hi = column_of((bbox.x0 + bbox.x1) / 2, col_ranges)
    return fitz.Rect(max(lo, bind_x, bbox.x0 - 2), bbox.y0 - 2,
                     min(hi, bbox.x1 + 2), bbox.y1 + 2)


def region_has_graphics(page, region, infos):
    for d in page.get_drawings():
        if fitz.Rect(d["rect"]).intersects(region):
            return True
    for i in infos:
        if fitz.Rect(i["bbox"]).intersects(region):
            return True
    return False


def crop_pdf(pdf_path, out_dir, dpi, cols):
    doc = fitz.open(pdf_path)
    out_dir = Path(out_dir)
    assets = out_dir / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    produced = []
    report = []

    def add(name, im, source, page_no, **criteria):
        im2 = trim_white(im)
        path = assets / name
        im2.save(path)
        produced.append(path)
        report.append({"file": name, "page": page_no, "source": source,
                       "result": f"已产出，尺寸 {im2.width}x{im2.height}", **criteria})

    def skip(name, page_no, source, reason, **criteria):
        report.append({"file": name, "page": page_no, "source": source,
                       "result": "跳过：" + reason, **criteria})

    for pno in range(1, len(doc) + 1):
        page = doc[pno - 1]
        infos = [i for i in page.get_image_info(xrefs=True) if i.get("xref")]
        ncols, col_ranges = detect_columns(page)
        bind_x = scan_binding_lines(page)
        captions = [b for b in page.get_text("blocks")
                    if b[4].strip() and CAPTION_RE.match(b[4].strip().splitlines()[0])]

        # 1) 位图优先（坑 16）：xref 原图 + 同 bbox 的 300dpi clip 对照
        for k, info in enumerate(infos, 1):
            xref = info["xref"]
            bbox = fitz.Rect(info["bbox"])
            try:
                data = doc.extract_image(xref)
                im = Image.open(io.BytesIO(data["image"])).convert("RGB")
                add(f"fig_p{pno}_xref{xref}.png", im, "位图 xref", pno,
                    bbox=fmt_rect(bbox), column=f"{ncols}栏",
                    binding_line_x=round(bind_x, 1), prose="n/a（xref 原图不混正文）")
            except Exception as e:
                skip(f"fig_p{pno}_xref{xref}.png", pno, "位图 xref", f"提取失败：{e}")
            clip = clip_in_column(bbox, col_ranges, bind_x)
            try:
                rim = render_clip(page, clip, dpi)
                add(f"fig_p{pno}_clip{k}.png", rim, "bbox clip 300dpi（对照）", pno,
                    bbox=fmt_rect(clip), column=f"{ncols}栏",
                    binding_line_x=round(bind_x, 1))
            except Exception as e:
                skip(f"fig_p{pno}_clip{k}.png", pno, "bbox clip", f"渲染失败：{e}")

        # 2) 文字/矢量页：caption 定位 + clip 渲染（坑 41/28）
        for ci, cap in enumerate(captions, 1):
            cap_rect = fitz.Rect(cap[:4])
            x0, x1 = column_of(cap_rect.x0, col_ranges)
            x0 = max(x0, bind_x, 0)
            y_bottom = max(PROSE_Y1_MIN + 4, cap_rect.y0 - 2)
            region = fitz.Rect(x0, PROSE_Y1_MIN, x1, y_bottom)
            blocks = [b for b in page.get_text("blocks")
                      if b[4].strip() and fitz.Rect(b[:4]).intersects(region)]
            ratio, longest = prose_metrics(" ".join(b[4] for b in blocks))
            max_y1 = max((b[3] for b in blocks), default=0)
            prose_hit = (ratio >= PROSE_RATIO_MIN and
                         longest >= LONGEST_WORD_MIN and max_y1 >= PROSE_Y1_MIN)
            graphic = region_has_graphics(page, region, infos)
            crit = {"prose_ratio": round(ratio, 2), "longest_word": longest,
                    "y1": round(max_y1, 1), "prose_hit": prose_hit,
                    "column": f"{ncols}栏", "has_graphics": graphic,
                    "binding_line_x": round(bind_x, 1)}
            reasons = []
            if prose_hit:
                reasons.append("prose 三判据命中（混入正文，坑 28）")
            if not graphic:
                reasons.append("区域无矢量绘图/位图（纯文字页）")
            if reasons:
                skip(f"fig_p{pno}_cap{ci}.png", pno, "caption clip", "；".join(reasons), **crit)
                continue
            try:
                rim = render_clip(page, region, dpi)
                add(f"fig_p{pno}_cap{ci}.png", rim, "caption clip 300dpi", pno,
                    bbox=fmt_rect(region), **crit)
            except Exception as e:
                skip(f"fig_p{pno}_cap{ci}.png", pno, "caption clip", f"渲染失败：{e}", **crit)

        if not infos and not captions:
            report.append({"file": "—", "page": pno, "source": "—",
                           "result": "无嵌入位图、无 caption → 无图可裁",
                           "column": f"{ncols}栏", "binding_line_x": round(bind_x, 1)})

    # contact sheet（坑 16/28 的目视闭环）
    sheet = out_dir / "preview" / "contact_sheet.png"
    sheet_n = 0
    if build_contact_sheet is not None:
        _, sheet_n = build_contact_sheet(produced, sheet, cols=cols)

    _write_report(out_dir, pdf_path, dpi, produced, report, sheet, sheet_n)
    doc.close()
    return produced, report, sheet


def _write_report(out_dir, pdf_path, dpi, produced, report, sheet, sheet_n):
    lines = [f"# 裁图报告：{Path(pdf_path).name}", "",
             f"- --dpi = {dpi}（Matrix({dpi}/72, {dpi}/72)）",
             f"- 产出资产 **{len(produced)}** 张；contact sheet：`{sheet.as_posix()}`（{sheet_n} 张图）", "",
             "## 每张图的来源与判据命中", ""]
    for r in report:
        lines += [f"### `{r['file']}`（p{r['page']}）", "",
                  f"- 来源：**{r['source']}**", f"- 结果：{r['result']}"]
        for k, v in r.items():
            if k not in ("file", "page", "source", "result"):
                lines.append(f"- {k} = {v}")
        lines.append("")
    lines += ["## 诚实边界", "",
              "- 位图按 xref 原分辨率提取（坑 16），clip 仅作同区域对照；二者均做白边 trim。",
              "- caption 区域若命中 prose 三判据或不含图形内容则跳过，不产出误导性「正文图」。",
              "- 装订线为像素级扫描结果（坑 42），复杂底纹页仍需 contact sheet 人工复核。", ""]
    out = Path(out_dir) / "crop_report.md"
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


def main():
    ap = argparse.ArgumentParser(description="按坑 16/28/41/42 固化裁图规则（只读 PDF）")
    ap.add_argument("pdf", help="教材 PDF")
    ap.add_argument("--out", default="crop_out", help="输出目录（assets/ + preview/ + 报告）")
    ap.add_argument("--dpi", type=int, default=300, help="clip 渲染分辨率（默认 300）")
    ap.add_argument("--cols", type=int, default=4, help="contact sheet 列数")
    args = ap.parse_args()

    # 闸门放在 argparse 之后：--help 永远可用（P2-1）
    from license_check import require_license_or_exit
    require_license_or_exit("core")

    produced, report, sheet = crop_pdf(args.pdf, args.out, args.dpi, args.cols)
    print(f"裁图完成：assets {len(produced)} 张 | contact sheet {sheet.as_posix()}")
    print(f"候选/跳过记录 {len(report)} 条，详见 {Path(args.out) / 'crop_report.md'}")
    for p in produced:
        print("  asset:", p.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
