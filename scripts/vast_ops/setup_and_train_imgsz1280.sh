#!/usr/bin/env bash
# Setup VAST và train Phase 2 với imgsz=1280
# Re-prepare ACDC ở 1280px, rebuild merged dataset, train RT-DETR-L
#
# Yêu cầu: rclone đã config với remote "gdrive"
# Chạy: bash scripts/vast_ops/setup_and_train_imgsz1280.sh

set -euo pipefail

WORKSPACE="/workspace"
DATASETS_DIR="$WORKSPACE/datasets_noleak"
RUNS_DIR="$WORKSPACE/runs"
REPO_DIR="$WORKSPACE/repo"
DRIVE_ROOT="gdrive:object_detection_in_adverse_weather"

WEIGHTS="$RUNS_DIR/phase2_final_rtdetr/weights/best.pt"
ACDC_RAW="$WORKSPACE/datasets/acdc_raw"
ACDC_1280="$DATASETS_DIR/acdc_6cls_yolo_1280"
PHASE2_1280="$DATASETS_DIR/phase2_merged_1280"

# ── 1. Cài dependencies ──────────────────────────────────────────────────────
echo "=== 1. Cài dependencies ==="
pip install ultralytics opencv-python-headless Pillow numpy --quiet
if ! command -v rclone &>/dev/null; then
    curl https://rclone.org/install.sh | bash
fi

# ── 2. Clone repo ────────────────────────────────────────────────────────────
echo "=== 2. Clone repo ==="
if [ ! -d "$REPO_DIR" ]; then
    git clone https://github.com/thanhhbao/Object-Detection.git "$REPO_DIR"
fi
cd "$REPO_DIR"
git pull
export PYTHONPATH="$REPO_DIR/src:${PYTHONPATH:-}"

# ── 3. Download từ Drive ─────────────────────────────────────────────────────
echo "=== 3. Download datasets + weights từ Drive ==="
mkdir -p "$DATASETS_DIR"

# Dataset đã chuẩn hóa (XWOD + BDD, không cần prepare lại)
if [ ! -d "$DATASETS_DIR/xwod_6cls_yolo" ]; then
    echo "  Downloading datasets_noleak_official.tar (~3.8GB)..."
    rclone copy \
        "$DRIVE_ROOT/official_backup_from_old_machine/datasets_noleak_official.tar" \
        "$WORKSPACE/" --progress
    tar -xf "$WORKSPACE/datasets_noleak_official.tar" -C "$WORKSPACE/"
fi

# ACDC raw (cần để re-prepare ở 1280)
if [ ! -d "$ACDC_RAW" ]; then
    echo "  Downloading ACDC raw..."
    mkdir -p "$WORKSPACE/datasets"
    rclone copy \
        "$DRIVE_ROOT/official_backup_from_old_machine/acdc_raw.tar.gz" \
        "$WORKSPACE/" --progress
    tar -xzf "$WORKSPACE/acdc_raw.tar.gz" -C "$WORKSPACE/datasets/"
fi

# Model weights (Stage 2 checkpoint để tiếp tục train)
if [ ! -f "$WEIGHTS" ]; then
    echo "  Downloading phase2_final_rtdetr weights..."
    rclone copy \
        "$DRIVE_ROOT/runs/phase2/phase2_final_rtdetr/phase2_final_rtdetr.tar.gz" \
        "$WORKSPACE/" --progress
    tar -xzf "$WORKSPACE/phase2_final_rtdetr.tar.gz" -C "$RUNS_DIR/"
fi
[ -f "$WEIGHTS" ] || { echo "ERROR: weights không tìm thấy: $WEIGHTS"; exit 1; }

# ── 4. Re-prepare ACDC ở imgsz=1280 ─────────────────────────────────────────
echo "=== 4. Re-prepare ACDC ở 1280px ==="
if [ ! -d "$ACDC_1280" ]; then
    python scripts/data_prep/raw_to_yolo/prepare_acdc.py \
        --raw-dir "$ACDC_RAW" \
        --output-dir "$ACDC_1280" \
        --imgsz 1280 \
        --seed 42 --clean
else
    echo "  ACDC 1280 đã có, bỏ qua"
fi

# ── 5. Rebuild Phase 2 merged dataset ────────────────────────────────────────
echo "=== 5. Rebuild Phase 2 merged dataset ==="
if [ ! -d "$PHASE2_1280" ]; then
    python scripts/data_prep/phase2_build/build_phase2_dataset.py \
        --acdc-dir "$ACDC_1280" \
        --xwod-dir "$DATASETS_DIR/xwod_6cls_yolo" \
        --bdd-dir  "$DATASETS_DIR/bdd_6cls_yolo" \
        --output-dir "$PHASE2_1280" \
        --seed 42 --bdd-use-all
else
    echo "  Phase2 merged 1280 đã có, bỏ qua"
fi

# ── 6. Preflight check ───────────────────────────────────────────────────────
echo "=== 6. Preflight check ==="
python scripts/train/preflight/preflight_phase2.py \
    --config configs/ultralytics/phase2_rtdetr_imgsz1280.yaml \
    || { echo "WARN: preflight fail — kiểm tra lại dataset"; }

# ── 7. Train ─────────────────────────────────────────────────────────────────
echo "=== 7. Train RT-DETR-L imgsz=1280 ==="
python scripts/train/train_ultralytics.py \
    --config configs/ultralytics/phase2_rtdetr_imgsz1280.yaml \
    --data "$PHASE2_1280/dataset.yaml"

# ── 8. Upload weights + results lên Drive ────────────────────────────────────
echo "=== 8. Upload kết quả lên Drive ==="
RESULT_RUN="$RUNS_DIR/phase2_rtdetr_imgsz1280"
DEST="$DRIVE_ROOT/runs/phase2/phase2_rtdetr_imgsz1280"
tar -czf "$WORKSPACE/phase2_rtdetr_imgsz1280.tar.gz" -C "$RUNS_DIR" phase2_rtdetr_imgsz1280/
rclone copy "$WORKSPACE/phase2_rtdetr_imgsz1280.tar.gz" "$DEST/" --progress
echo "  Saved to Drive: $DEST"
echo "=== Hoàn tất ==="
