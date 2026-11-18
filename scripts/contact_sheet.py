#!/usr/bin/env python3
"""contact_sheet.py — 把裁出的图拼成带标签的 contact sheet，逐张视觉校验（坑 28/16）。

可单独用：python scripts/contact_sheet.py <assets_dir|图片...> --out sheet.png [--cols 4]
也可被 crop_figures.py 直接 import build_contact_sheet() 在裁图末尾自动拼图。
"""
import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}


def _font():
    try:
        return ImageFont.truetype("arial.ttf", 16)
    except Exception:
        return ImageFont.load_default()


def build_contact_sheet(image_paths, out_path, cols=4, cell=(360, 300), label_h=26):
    """把 image_paths 拼成网格。返回 (写出路径, 图片数)。无图也写一张提示 sheet。"""
    paths = [Path(p) for p in image_paths]
    font = _font()
    cell_w, cell_h = cell
    box_h = cell_h - label_h
    rows = (len(paths) + cols - 1) // cols if paths else 1
    sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), (245, 245, 245))
    draw = ImageDraw.Draw(sheet)
    for idx, p in enumerate(paths):
        r, c = divmod(idx, cols)
        x, y = c * cell_w, r * cell_h
        try:
            im = Image.open(p).convert("RGB")
            im.thumbnail((cell_w - 12, box_h - 12), Image.Resampling.LANCZOS)
            sheet.paste(im, (x + (cell_w - im.width) // 2, y + (box_h - im.height) // 2))
        except Exception as e:
            draw.text((x + 8, y + 8), f"打不开: {e}", fill=(180, 0, 0), font=font)
        label = p.name if len(p.name) <= 44 else p.name[:41] + "..."
        draw.text((x + 6, y + box_h + 3), label, fill=(20, 20, 20), font=font)
    if not paths:
        draw.text((12, 12), "（无裁图资产）", fill=(120, 120, 120), font=font)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    return out, len(paths)


def _collect(targets):
    out = []
    for t in targets:
        p = Path(t)
        if p.is_dir():
            out += sorted(q for q in p.iterdir() if q.suffix.lower() in EXTS)
        elif p.suffix.lower() in EXTS:
            out.append(p)
    return out


def main():
    ap = argparse.ArgumentParser(description="拼 contact sheet（逐张视觉校验裁图）")
    ap.add_argument("targets", nargs="+", help="图片文件或 assets 目录")
    ap.add_argument("--out", default="preview/contact_sheet.png")
    ap.add_argument("--cols", type=int, default=4)
    args = ap.parse_args()
    paths = _collect(args.targets)
    out, n = build_contact_sheet(paths, args.out, cols=args.cols)
    print(f"contact sheet: {out}（{n} 张图）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
