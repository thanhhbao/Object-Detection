#!/usr/bin/env python3
"""Tăng cường dữ liệu thời tiết (offline) cho dataset YOLO — có hỗ trợ lọc theo prefix nguồn.

Sinh thêm các bản suy giảm quang học để tạo Phase 2 weather-aug.
Nhãn bounding box GIỮ NGUYÊN vì các biến đổi đều là quang học (không làm dời hộp bao).

Chế độ 1 — toàn bộ dataset (tương thích cũ):
  python scripts/augment_weather.py --src /data/phase2_v3 --dst /data/phase2_v3_waug
      --copies 1 --seed 42

Chế độ 2 — lọc theo prefix nguồn (khuyến nghị cho Phase 2):
  python scripts/augment_weather.py --src /data/phase2_v3 --dst /data/phase2_v3_waug
      --bdd-ratio 0.70 --xwod-ratio 0.12 --acdc-ratio 0.0 --seed 42

Quy ước prefix: build_phase2_dataset.py gán bdd_ / xwod_ / acdc_ cho mỗi ảnh.
Ảnh oversampling (bdd_foo_os1, xwod_bar_os2) bị loại khỏi nguồn augmentation —
bản gốc vẫn có trong dataset, nhưng không dùng làm nguồn tạo thêm weather copy.

Đầu ra:
  dst/images/train/  — ảnh gốc + ảnh tăng cường (_waug{k})
  dst/labels/train/  — nhãn tương ứng (bản sao)
  dst/images/val|test — giữ nguyên từ src
  dst/dataset.yaml   — cập nhật path
  dst/waug_manifest.json — provenance đầy đủ: tên file, nguồn, transform đã áp
"""

from __future__ import annotations

import argparse
import json
import random
import re
import shutil
from pathlib import Path

import albumentations as A
import cv2
import numpy as np
import yaml

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

# Ảnh oversampling từ build_phase2_dataset.py kết thúc bằng _os<số>
_OS_RE = re.compile(r"_os\d+$")


# ---------------------------------------------------------------------------
# Named transform groups — chọn thủ công để ghi tên vào manifest
# ---------------------------------------------------------------------------

_STRONG_PRIMARIES = [
    ("fog",    lambda: A.RandomFog(p=1.0)),
    ("rain",   lambda: A.RandomRain(p=1.0)),
    ("snow",   lambda: A.RandomSnow(p=1.0)),
    ("shadow", lambda: A.RandomShadow(p=1.0)),
]

_LIGHT_OPTIONS = [
    ("brightness_contrast", lambda: A.RandomBrightnessContrast(
        brightness_limit=0.2, contrast_limit=0.2, p=1.0)),
    ("hue_saturation",      lambda: A.HueSaturationValue(
        hue_shift_limit=5, sat_shift_limit=20, val_shift_limit=20, p=1.0)),
    ("gaussian_blur",       lambda: A.GaussianBlur(blur_limit=(3, 5), p=1.0)),
]


def _seed_all(aug_seed: int) -> None:
    """Seed cả global random lẫn numpy để Albumentations 1.x tái lập được."""
    random.seed(aug_seed)
    np.random.seed(aug_seed)


def _apply_strong(
    image_rgb: np.ndarray, py_rng: random.Random, aug_seed: int
) -> tuple[np.ndarray, dict]:
    """Áp strong transform; trả về ảnh và dict ghi đầy đủ mọi phép biến đổi đã áp.

    Mọi quyết định nhị phân (có/không áp brightness-contrast, secondary transform)
    đều dùng py_rng (deterministic theo seed chọn ảnh), không phụ thuộc vào
    trạng thái global sau khi seed augmentation pixel.
    """
    primary_name, primary_factory = py_rng.choice(_STRONG_PRIMARIES)
    apply_bc      = py_rng.random() < 0.6
    apply_second  = py_rng.random() < 0.4
    second_name   = py_rng.choice(["motion_blur", "gauss_noise"]) if apply_second else None

    transforms = [primary_factory()]
    if apply_bc:
        transforms.append(A.RandomBrightnessContrast(
            brightness_limit=0.3, contrast_limit=0.3, p=1.0))
    if second_name == "motion_blur":
        transforms.append(A.MotionBlur(blur_limit=5, p=1.0))
    elif second_name == "gauss_noise":
        transforms.append(A.GaussNoise(p=1.0))

    _seed_all(aug_seed)
    result = A.Compose(transforms)(image=image_rgb)["image"]

    tracking = {
        "primary": primary_name,
        "brightness_contrast": apply_bc,
        "secondary": second_name,
    }
    return result, tracking


def _apply_light(
    image_rgb: np.ndarray, py_rng: random.Random, aug_seed: int
) -> tuple[np.ndarray, dict]:
    """Áp light transform; trả về ảnh và dict mô tả phép biến đổi."""
    name, tf_factory = py_rng.choice(_LIGHT_OPTIONS)
    _seed_all(aug_seed)
    result = A.Compose([tf_factory()])(image=image_rgb)["image"]
    tracking = {"primary": name, "brightness_contrast": False, "secondary": None}
    return result, tracking


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


def _detect_prefix(stem: str) -> str:
    for prefix in ("bdd_", "xwod_", "acdc_"):
        if stem.startswith(prefix):
            return prefix.rstrip("_")
    return ""


def _is_oversampled(stem: str) -> bool:
    return bool(_OS_RE.search(stem))


# ---------------------------------------------------------------------------
# Core augmentation loop
# ---------------------------------------------------------------------------

def _augment_split(
    src_images: Path,
    dst_images: Path,
    dst_labels: Path,
    copies: int,
    bdd_ratio: float,
    xwod_ratio: float,
    acdc_ratio: float,
    seed: int,
) -> list[dict]:
    """Tạo ảnh tăng cường cho split; trả về manifest entries."""
    py_rng = random.Random(seed)
    # Seed riêng cho augmentation pixel (tách khỏi seed chọn ảnh)
    np_rng = np.random.default_rng(seed)

    src_labels = _label_dir_of(src_images)
    images = sorted(p for p in src_images.rglob("*") if p.suffix.lower() in IMAGE_SUFFIXES)
    if not images:
        return []

    use_prefix_mode = (bdd_ratio > 0 or xwod_ratio > 0 or acdc_ratio > 0)

    manifest: list[dict] = []

    for image_path in images:
        stem = image_path.stem
        prefix = _detect_prefix(stem)

        # Loại ảnh oversampling ra khỏi nguồn augmentation
        if _is_oversampled(stem):
            continue

        if use_prefix_mode:
            ratio_map = {"bdd": bdd_ratio, "xwod": xwod_ratio, "acdc": acdc_ratio}
            ratio = ratio_map.get(prefix, 0.0)
            if ratio == 0.0 or py_rng.random() > ratio:
                continue
            use_strong = (prefix in ("bdd", ""))
            n_copies = 1
        else:
            use_strong = True
            n_copies = copies

        # Đọc ảnh BGR và chuyển sang RGB cho Albumentations
        bgr = cv2.imread(str(image_path))
        if bgr is None:
            continue
        image_rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)

        label_path = src_labels / f"{stem}.txt"

        for k in range(n_copies):
            # Seed augmentation dựa trên tên file + copy index để tái lập
            aug_seed = int(np_rng.integers(0, 2**31))

            try:
                if use_strong:
                    aug_rgb, tracking = _apply_strong(image_rgb, py_rng, aug_seed)
                else:
                    aug_rgb, tracking = _apply_light(image_rgb, py_rng, aug_seed)
            except Exception:
                # Bỏ qua ảnh quá nhỏ hoặc lỗi Albumentations
                continue

            aug_stem = f"{stem}_waug{k}"
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
                "source": prefix or "unknown",
                "transform_group": "strong" if use_strong else "light",
                "transform_primary": tracking["primary"],
                "transform_brightness_contrast": tracking["brightness_contrast"],
                "transform_secondary": tracking["secondary"],
                "aug_seed": aug_seed,
                "copy_index": k,
            })

    return manifest


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Offline weather augmentation for YOLO dataset")
    parser.add_argument("--src", type=Path, required=True, help="Dataset YOLO gốc (Phase 2 v3)")
    parser.add_argument("--dst", type=Path, required=True, help="Dataset đầu ra với augmentation")
    parser.add_argument("--split", default="train", help="Split cần tăng cường (mặc định: train)")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--clean", action="store_true",
                        help="Xóa dst nếu đã tồn tại trước khi tạo mới (mặc định: từ chối nếu dst không rỗng)")

    # Chế độ toàn bộ dataset (tương thích cũ)
    parser.add_argument("--copies", type=int, default=1,
                        help="Số bản copy mỗi ảnh (chế độ không dùng prefix-ratio)")

    # Chế độ lọc theo prefix
    parser.add_argument("--bdd-ratio", type=float, default=0.0,
                        help="Tỉ lệ ảnh gốc bdd_* nhận strong aug (0=bỏ qua). Ảnh _osK bị loại.")
    parser.add_argument("--xwod-ratio", type=float, default=0.0,
                        help="Tỉ lệ ảnh gốc xwod_* nhận light aug (0=bỏ qua). Ảnh _osK bị loại.")
    parser.add_argument("--acdc-ratio", type=float, default=0.0,
                        help="Tỉ lệ ảnh gốc acdc_* nhận aug (0=bỏ qua).")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    src, dst = args.src.resolve(), args.dst.resolve()

    # Ngăn vô tình xóa dataset nguồn
    if src == dst:
        raise SystemExit("ERROR: --src và --dst phải là hai đường dẫn khác nhau.")

    # Kiểm tra dst trước khi ghi đè
    if dst.exists() and any(dst.iterdir()):
        if args.clean:
            shutil.rmtree(dst)
            print(f"Đã xóa {dst} (--clean)")
        else:
            raise SystemExit(
                f"ERROR: {dst} đã tồn tại và không rỗng.\n"
                "Dùng --clean để xóa và tạo lại, hoặc chọn đường dẫn khác."
            )

    # Sao chép toàn bộ dataset gốc (giữ val/test + ảnh train gốc)
    dst.mkdir(parents=True, exist_ok=True)
    for sub in ("images", "labels"):
        _copy_tree(src / sub, dst / sub)

    dst_images_split = dst / "images" / args.split
    dst_labels_split = dst / "labels" / args.split
    dst_images_split.mkdir(parents=True, exist_ok=True)
    dst_labels_split.mkdir(parents=True, exist_ok=True)

    manifest = _augment_split(
        src_images=src / "images" / args.split,
        dst_images=dst_images_split,
        dst_labels=dst_labels_split,
        copies=args.copies,
        bdd_ratio=args.bdd_ratio,
        xwod_ratio=args.xwod_ratio,
        acdc_ratio=args.acdc_ratio,
        seed=args.seed,
    )

    # Ghi provenance manifest (aug_seed cho phép tái tạo pixel từng ảnh)
    manifest_path = dst / "waug_manifest.json"
    import albumentations as _A
    manifest_path.write_text(
        json.dumps({
            "seed": args.seed,
            "bdd_ratio": args.bdd_ratio,
            "xwod_ratio": args.xwod_ratio,
            "acdc_ratio": args.acdc_ratio,
            "copies": args.copies,
            "os_images_excluded": True,
            "albumentations_version": _A.__version__,
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
    total_train = sum(1 for _ in dst_images_split.rglob("*")
                      if _.suffix.lower() in IMAGE_SUFFIXES)
    by_source: dict[str, int] = {}
    by_transform: dict[str, int] = {}
    for entry in manifest:
        src_key = entry["source"]
        by_source[src_key] = by_source.get(src_key, 0) + 1
        t_key = entry["transform_primary"]
        by_transform[t_key] = by_transform.get(t_key, 0) + 1

    print(f"\nĐã tạo {len(manifest)} ảnh tăng cường:")
    for src_key, count in sorted(by_source.items()):
        print(f"  nguồn {src_key:8s}: +{count}")
    print("  theo transform:")
    for t_key, count in sorted(by_transform.items()):
        print(f"    {t_key:30s}: {count}")
    print(f"Train tổng cộng: {total_train} ảnh")
    print(f"Dataset mới: {dst}/dataset.yaml")
    print(f"Manifest   : {manifest_path}")


if __name__ == "__main__":
    main()
