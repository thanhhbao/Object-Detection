#!/usr/bin/env python3
"""Build numbered contact sheets from page images for visual DOCX/PDF QA."""

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("image_dir", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--prefix", default="page-")
    parser.add_argument("--per-sheet", type=int, default=12)
    args = parser.parse_args()

    pages = sorted(args.image_dir.glob(f"{args.prefix}*.jpg"))
    if not pages:
        raise SystemExit("No page JPG files found")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    font = ImageFont.load_default()
    cols = 3
    rows = (args.per_sheet + cols - 1) // cols
    thumb_w, thumb_h = 420, 594
    label_h, gap = 24, 16
    sheet_w = cols * thumb_w + (cols + 1) * gap
    sheet_h = rows * (thumb_h + label_h) + (rows + 1) * gap
    for sheet_index, start in enumerate(range(0, len(pages), args.per_sheet), 1):
        sheet = Image.new("RGB", (sheet_w, sheet_h), "#d8d8d8")
        draw = ImageDraw.Draw(sheet)
        for slot, path in enumerate(pages[start : start + args.per_sheet]):
            row, col = divmod(slot, cols)
            x = gap + col * (thumb_w + gap)
            y = gap + row * (thumb_h + label_h + gap)
            with Image.open(path) as page:
                page = page.convert("RGB")
                page.thumbnail((thumb_w, thumb_h))
                px = x + (thumb_w - page.width) // 2
                py = y + label_h
                sheet.paste(page, (px, py))
            page_number = start + slot + 1
            draw.text((x + 4, y + 4), f"Page {page_number}", fill="black", font=font)
        sheet.save(args.output_dir / f"contact_{sheet_index:02d}.jpg", quality=88)


if __name__ == "__main__":
    main()
