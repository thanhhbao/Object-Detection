#!/usr/bin/env python3
"""Tăng cường thiếu sáng (offline) cho ảnh BDD replay trong Phase 2 dataset — E4-dark.

Chỉ áp dụng cho ảnh bdd_* gốc (loại _osK), tạo một bản dark cho mỗi ảnh được chọn.
Không chạm vào XWOD hay ACDC; không kết hợp với weather augmentation (E3).
Mục tiêu: kiểm tra độc lập giả thuyết "mất thông tin do thiếu sáng" là nút thắt ACDC.

Hai kiểu biến đổi:
  gamma   — tối ảnh qua lũy thừa pixel: pixel^gamma (gamma>1 → tối hơn)
  bc      — giảm brightness trực tiếp; contrast có thể thay đổi nhẹ

Cả hai đều lấy tham số chính xác theo py_rng (không phải dải ngẫu nhiên của Albumentations)
nên manifest ghi được giá trị thực tế, không chỉ tên kiểu biến đổi.

Kiểm tra histogram: ảnh augmented có mean pixel < MIN_MEAN_BRIGHTNESS được bỏ qua
để tránh sinh ảnh gần như hoàn toàn đen.

Chạy:
  python scripts/augment_dark.py \\
      --src /workspace/datasets/phase2_v3 \\
      --dst /workspace/datasets/phase2_v3_dark \\
      --bdd-ratio 0.60 \\
      --seed 42 --clean

Kiểm tra sau khi chạy:
  1. Xem histogram một số ảnh _dark0: phải còn chi tiết, không hoàn toàn đen.
  2. Kiểm tra waug_dark_manifest.json: giá trị gamma/brightness phân bố đều.
  3. Xác nhận chỉ ảnh bdd_* (không có _osK) trong manifest.
"""

from __future__ import annotations

import argparse
import json
import random
import shutil
from pathlib import Path

import albumentations as A
import cv2
import numpy as np
import yaml

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

# Ảnh quá tối sau augmentation bị loại (mean pixel / 255)
MIN_MEAN_BRIGHTNESS = 15.0 / 255.0   # ~6% of full range

import re
_OS_RE = re.compile(r"_os\d+$")


# ---------------------------------------------------------------------------
# Named dark transforms — tham số được lấy tường minh để ghi manifest
# ---------------------------------------------------------------------------

def _apply_gamma_dark(
    image_rgb: np.ndarray, py_rng: random.Random, aug_seed: int
) -> tuple[np.ndarray, dict]:
    """Tối ảnh qua gamma > 1. Gamma 1.2–1.8 cho dải tối hợp lý."""
    gamma_val = round(py_rng.uniform(1.20, 1.80), 3)
    # Albumentations RandomGamma nhận gamma_limit theo đơn vị *100
    gamma_int = int(gamma_val * 100)
    tf = A.RandomGamma(gamma_limit=(gamma_int, gamma_int), p=1.0)
    random.seed(aug_seed)
    np.random.seed(aug_seed)
    result = A.Compose([tf])(image=image_rgb)["image"]
    return result, {"transform": "gamma", "gamma": gamma_val}


def _apply_bc_dark(
    image_rgb: np.ndarray, py_rng: random.Random, aug_seed: int
) -> tuple[np.ndarray, dict]:
    """Tối ảnh qua brightness âm + điều chỉnh contrast nhẹ."""
    brightness = round(py_rng.uniform(-0.35, -0.12), 3)
    contrast   = round(py_rng.uniform(-0.10,  0.20), 3)
    tf = A.RandomBrightnessContrast(
        brightness_limit=(brightness, brightness),
        contrast_limit=(contrast, contrast),
        p=1.0,
    )
    random.seed(aug_seed)
    np.random.seed(aug_seed)
    result = A.Compose([tf])(image=image_rgb)["image"]
    return result, {"transform": "brightness_contrast", "brightness": brightness, "contrast": contrast}


_DARK_TRANSFORMS = [_apply_gamma_dark, _apply_bc_dark]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _label_dir_of(images_split_dir: Path) -> Path:
    parts = list(images_split_dir.parts)
    for i in range(len(parts) - 1, -1, -1):
        if parts[i] == "images":
            parts[i] = "labels"
            break
    return Path(*parts)


def _copy_tree(src_dir: Path, dst_dir: Path) -> None:
    if src_dir.exists():
        shutil.copytree(src_dir, dst_dir, dirs_exist_ok=True)


def _is_oversampled(stem: str) -> bool:
    return bool(_OS_RE.search(stem))


# ---------------------------------------------------------------------------
# Core loop
# ---------------------------------------------------------------------------

def _augment_split(
    src_images: Path,
    dst_images: Path,
    dst_labels: Path,
    bdd_ratio: float,
    seed: int,
) -> tuple[list[dict], int]:
    """Tạo dark copies cho ảnh bdd_* gốc; trả về (manifest, n_skipped_too_dark)."""
    py_rng = random.Random(seed)
    np_rng = np.random.default_rng(seed)
    src_labels = _label_dir_of(src_images)

    images = sorted(p for p in src_images.rglob("*") if p.suffix.lower() in IMAGE_SUFFIXES)
    manifest: list[dict] = []
    n_skipped_too_dark = 0

    for image_path in images:
        stem = image_path.stem

        # Chỉ lấy ảnh bdd_* gốc
        if not stem.startswith("bdd_"):
            continue
        if _is_oversampled(stem):
            continue
        if py_rng.random() > bdd_ratio:
            continue

        bgr = cv2.imread(str(image_path))
        if bgr is None:
            continue
        image_rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)

        label_path = src_labels / f"{stem}.txt"
        aug_seed = int(np_rng.integers(0, 2**31))

        # Chọn ngẫu nhiên một trong hai kiểu dark transform
        tf_fn = py_rng.choice(_DARK_TRANSFORMS)

        try:
            aug_rgb, params = tf_fn(image_rgb, py_rng, aug_seed)
        except Exception:
            continue

        # Kiểm tra histogram: bỏ ảnh gần như hoàn toàn đen
        mean_brightness = aug_rgb.mean() / 255.0
        if mean_brightness < MIN_MEAN_BRIGHTNESS:
            n_skipped_too_dark += 1
            continue

        aug_stem = f"{stem}_dark0"
        out_img = dst_images / f"{aug_stem}{image_path.suffix}"
        aug_bgr = cv2.cvtColor(aug_rgb, cv2.COLOR_RGB2BGR)
        cv2.imwrite(str(out_img), aug_bgr, [cv2.IMWRITE_JPEG_QUALITY, 95])

        if label_path.exists():
            shutil.copy2(label_path, dst_labels / f"{aug_stem}.txt")
        else:
            (dst_labels / f"{aug_stem}.txt").write_text("", encoding="utf-8")

        manifest.append({
            "aug_file": out_img.name,
            "src_file": image_path.name,
            "source": "bdd",
            "aug_seed": aug_seed,
            "mean_brightness_after": round(mean_brightness, 4),
            **params,
        })

    return manifest, n_skipped_too_dark


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="E4-dark: offline dark augmentation on BDD replay images only."
    )
    parser.add_argument("--src", type=Path, required=True)
    parser.add_argument("--dst", type=Path, required=True)
    parser.add_argument("--split", default="train")
    parser.add_argument("--bdd-ratio", type=float, default=0.60,
                        help="Tỉ lệ ảnh bdd_* gốc nhận dark aug (không tính _osK).")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--clean", action="store_true",
                        help="Xóa dst trước nếu không rỗng.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    src, dst = args.src.resolve(), args.dst.resolve()

    if src == dst:
        raise SystemExit("ERROR: --src và --dst phải khác nhau.")

    if dst.exists() and any(dst.iterdir()):
        if args.clean:
            shutil.rmtree(dst)
            print(f"Đã xóa {dst} (--clean)")
        else:
            raise SystemExit(
                f"ERROR: {dst} đã tồn tại và không rỗng.\n"
                "Dùng --clean hoặc chọn đường dẫn mới."
            )

    dst.mkdir(parents=True, exist_ok=True)
    for sub in ("images", "labels"):
        _copy_tree(src / sub, dst / sub)

    dst_img = dst / "images" / args.split
    dst_lbl = dst / "labels" / args.split
    dst_img.mkdir(parents=True, exist_ok=True)
    dst_lbl.mkdir(parents=True, exist_ok=True)

    manifest, n_too_dark = _augment_split(
        src_images=src / "images" / args.split,
        dst_images=dst_img,
        dst_labels=dst_lbl,
        bdd_ratio=args.bdd_ratio,
        seed=args.seed,
    )

    # Ghi manifest
    manifest_path = dst / "waug_dark_manifest.json"
    manifest_path.write_text(
        json.dumps({
            "seed": args.seed,
            "bdd_ratio": args.bdd_ratio,
            "bdd_only": True,
            "os_images_excluded": True,
            "min_mean_brightness_threshold": MIN_MEAN_BRIGHTNESS,
            "skipped_too_dark": n_too_dark,
            "albumentations_version": A.__version__,
            "augmentations": manifest,
        }, indent=2),
        encoding="utf-8",
    )

    # Cập nhật dataset.yaml
    data_yaml = src / "dataset.yaml"
    if data_yaml.exists():
        data = yaml.safe_load(data_yaml.read_text(encoding="utf-8"))
        data["path"] = str(dst)
        (dst / "dataset.yaml").write_text(
            yaml.safe_dump(data, sort_keys=False), encoding="utf-8"
        )

    # Thống kê
    n_gamma = sum(1 for e in manifest if e["transform"] == "gamma")
    n_bc    = sum(1 for e in manifest if e["transform"] == "brightness_contrast")
    total   = sum(1 for _ in dst_img.rglob("*") if _.suffix.lower() in IMAGE_SUFFIXES)
    brightness_vals = [e["mean_brightness_after"] for e in manifest]

    print(f"\nĐã tạo {len(manifest)} ảnh dark từ bdd_* gốc:")
    print(f"  gamma             : {n_gamma}")
    print(f"  brightness_contrast: {n_bc}")
    print(f"  bỏ qua (quá tối) : {n_too_dark}")
    if brightness_vals:
        print(f"  brightness after  : mean={sum(brightness_vals)/len(brightness_vals):.3f}"
              f"  min={min(brightness_vals):.3f}  max={max(brightness_vals):.3f}")
    print(f"Train tổng cộng: {total} ảnh")
    print(f"Dataset: {dst}/dataset.yaml")
    print(f"Manifest: {manifest_path}")


if __name__ == "__main__":
    main()
