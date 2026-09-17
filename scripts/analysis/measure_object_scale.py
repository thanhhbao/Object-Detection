#!/usr/bin/env python3
"""Measure ground-truth object scale in a way that is comparable across datasets.

Comparing box sizes in stored pixels is invalid when datasets are stored at
different resolutions: ACDC is written at 640x640 while XWOD keeps its native
1280x720, so the same real object yields very different stored pixel counts.

Two resolution-independent views are reported instead:

  * normalised side — sqrt(w*h) as a fraction of the image, straight from the
    YOLO label, independent of how the file happens to be stored;
  * side at model input — what the detector actually sees after letterboxing
    to --imgsz, which is what determines whether an object is detectable.

The second is the one that matters. An object below roughly 32 px at the input
falls in the COCO "small" band, where recall collapses.

Usage:
  python scripts/measure_object_scale.py \
    --dataset acdc=/workspace/datasets_noleak/acdc_6cls_yolo/dataset.yaml \
    --dataset xwod=/workspace/datasets_noleak/xwod_6cls_yolo/dataset.yaml \
    --split val --imgsz 640 --out /workspace/runs/evals/object_scale
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from tune_conf_thresholds import CLASS_NAMES, resolve_split  # noqa: E402

SMALL_PX = 32  # COCO small/medium boundary, on the side length


def collect(data_yaml: Path, split: str, imgsz: int) -> dict:
    from PIL import Image

    img_dir, lbl_dir = resolve_split(data_yaml, split)
    dims: dict[str, tuple] = {}
    per_class: dict[int, list] = defaultdict(list)
    img_sizes: list = []

    for lbl in sorted(lbl_dir.rglob("*.txt")):
        img = next((p for ext in (".jpg", ".jpeg", ".png")
                    for p in [img_dir / f"{lbl.stem}{ext}"] if p.is_file()), None)
        if img is None:
            hits = list(img_dir.rglob(f"{lbl.stem}.*"))
            img = hits[0] if hits else None
        if img is None:
            continue
        key = str(img)
        if key not in dims:
            with Image.open(img) as im:      # header only, no decode
                dims[key] = im.size
        W, H = dims[key]
        img_sizes.append((W, H))
        # Letterbox keeps aspect ratio: one scale for both axes.
        scale = min(imgsz / W, imgsz / H)

        for line in lbl.read_text(encoding="utf-8").splitlines():
            parts = line.split()
            if len(parts) < 5:
                continue
            c, nw, nh = int(parts[0]), float(parts[3]), float(parts[4])
            per_class[c].append((
                float(np.sqrt(nw * nh)),                    # normalised side
                float(np.sqrt(nw * W * scale * nh * H * scale)),  # side at input
            ))

    return {"per_class": per_class, "img_sizes": img_sizes,
            "n_images": len(img_sizes)}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dataset", action="append", required=True,
                    metavar="NAME=DATA_YAML")
    ap.add_argument("--split", default="val", choices=["train", "val", "test"])
    ap.add_argument("--imgsz", type=int, default=640)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    datasets = {}
    for v in args.dataset:
        if "=" not in v:
            raise SystemExit(f"--dataset expects name=path, got {v!r}")
        name, path = v.split("=", 1)
        if not Path(path).is_file():
            raise SystemExit(f"Dataset {name!r}: yaml not found: {path}")
        datasets[name] = Path(path)

    report = {"split": args.split, "imgsz": args.imgsz, "datasets": {}}

    for name, yml in datasets.items():
        print(f"\n{'='*74}\n{name}  ({args.split})\n{'='*74}")
        d = collect(yml, args.split, args.imgsz)
        sizes = d["img_sizes"]
        uniq = sorted({s for s in sizes}, key=lambda s: -sizes.count(s))[:3]
        print(f"{d['n_images']} anh; kich thuoc luu pho bien: "
              + ", ".join(f"{w}x{h}" for w, h in uniq))
        print(f"\n{'lop':12s} {'hop':>7s} {'canh chuan hoa':>15s} "
              f"{'canh tai dau vao':>17s} {'% nho (<32px)':>14s}")
        entry = {"n_images": d["n_images"],
                 "common_sizes": [f"{w}x{h}" for w, h in uniq]}
        for c, name_c in enumerate(CLASS_NAMES):
            vals = d["per_class"].get(c, [])
            if not vals:
                continue
            norm = np.array([v[0] for v in vals])
            inp = np.array([v[1] for v in vals])
            frac_small = float(np.mean(inp < SMALL_PX))
            print(f"{name_c:12s} {len(vals):7d} {np.median(norm):15.4f} "
                  f"{np.median(inp):14.1f} px {frac_small*100:13.1f}%")
            entry[name_c] = {
                "boxes": len(vals),
                "median_normalised_side": round(float(np.median(norm)), 4),
                "median_side_at_input_px": round(float(np.median(inp)), 1),
                "frac_small_at_input": round(frac_small, 4),
            }
        report["datasets"][name] = entry

    # ── cross-dataset ratio, on the only comparable quantity ────────────────
    names = list(datasets)
    if len(names) >= 2:
        base = names[0]
        print(f"\n{'='*74}\nTI LE canh tai dau vao, lay '{base}' lam moc\n{'='*74}")
        print(f"{'lop':12s}" + "".join(f"{n:>12s}" for n in names)
              + "".join(f"{'x'+n:>12s}" for n in names[1:]))
        for c, name_c in enumerate(CLASS_NAMES):
            v = {n: report["datasets"][n].get(name_c, {}).get(
                "median_side_at_input_px") for n in names}
            if any(x is None for x in v.values()):
                continue
            ratios = "".join(f"{v[n]/v[base]:12.2f}" for n in names[1:])
            print(f"{name_c:12s}" + "".join(f"{v[n]:12.1f}" for n in names)
                  + ratios)

    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "object_scale.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {args.out / 'object_scale.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
