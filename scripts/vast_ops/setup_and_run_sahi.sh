#!/usr/bin/env bash
# Setup VAST machine mới và chạy SAHI experiment
# Chạy trên VAST ngay sau khi machine khởi động:
#   bash scripts/vast_ops/setup_and_run_sahi.sh
#
# Yêu cầu: rclone đã được config với remote tên "gdrive" trỏ đến
#   Google Drive 2200001341@nttu.edu.vn
#   Nếu chưa config: chạy `rclone config` và làm theo hướng dẫn

set -euo pipefail

WORKSPACE="/workspace"
DATASETS_DIR="$WORKSPACE/datasets_noleak"
RUNS_DIR="$WORKSPACE/runs"
EVALS_DIR="$RUNS_DIR/evals"
WEIGHTS="$RUNS_DIR/phase2_final_rtdetr/weights/best.pt"
DRIVE_ROOT="gdrive:object_detection_in_adverse_weather"

# ── 1. Cài dependencies ──────────────────────────────────────────────────────
echo "=== 1. Cài dependencies ==="
pip install ultralytics opencv-python-headless Pillow numpy --quiet
# Cài rclone nếu chưa có
if ! command -v rclone &>/dev/null; then
    curl https://rclone.org/install.sh | bash
fi

# ── 2. Clone repo ────────────────────────────────────────────────────────────
echo "=== 2. Clone repo ==="
REPO_DIR="$WORKSPACE/repo"
if [ ! -d "$REPO_DIR" ]; then
    git clone https://github.com/thanhhbao/Object-Detection.git "$REPO_DIR"
fi
cd "$REPO_DIR"
export PYTHONPATH="$REPO_DIR/src:${PYTHONPATH:-}"

# ── 3. Download datasets từ Drive ────────────────────────────────────────────
echo "=== 3. Download datasets từ Drive ==="
mkdir -p "$DATASETS_DIR"
if [ ! -d "$DATASETS_DIR/acdc_6cls_yolo" ]; then
    echo "  Downloading datasets_noleak_official.tar (~3.8GB)..."
    rclone copy \
        "$DRIVE_ROOT/official_backup_from_old_machine/datasets_noleak_official.tar" \
        "$WORKSPACE/" \
        --progress
    tar -xf "$WORKSPACE/datasets_noleak_official.tar" -C "$WORKSPACE/"
    echo "  Giải nén xong"
else
    echo "  Dataset đã có, bỏ qua"
fi

# ── 4. Download model weights từ Drive ───────────────────────────────────────
echo "=== 4. Download model weights từ Drive ==="
mkdir -p "$RUNS_DIR"
if [ ! -f "$WEIGHTS" ]; then
    echo "  Downloading phase2_final_rtdetr.tar.gz (~122MB)..."
    rclone copy \
        "$DRIVE_ROOT/runs/phase2/phase2_final_rtdetr/phase2_final_rtdetr.tar.gz" \
        "$WORKSPACE/" \
        --progress
    tar -xzf "$WORKSPACE/phase2_final_rtdetr.tar.gz" -C "$RUNS_DIR/"
fi
[ -f "$WEIGHTS" ] || { echo "ERROR: weights không tìm thấy tại $WEIGHTS"; exit 1; }
echo "  OK: $WEIGHTS"

# ── 5. SAHI eval trên ACDC val ───────────────────────────────────────────────
echo "=== 5. SAHI — ACDC val ==="
python scripts/eval/eval_tiled_inference.py \
    --weights "$WEIGHTS" \
    --data "$DATASETS_DIR/acdc_6cls_yolo/dataset.yaml" \
    --split val \
    --tile 640 \
    --overlap 0.2 \
    --nms-iou 0.6 \
    --device 0 \
    --batch 8 \
    --out "$EVALS_DIR/sahi_acdc_val"

# ── 6. SAHI eval trên XWOD val (so sánh object lớn) ─────────────────────────
echo "=== 6. SAHI — XWOD val ==="
python scripts/eval/eval_tiled_inference.py \
    --weights "$WEIGHTS" \
    --data "$DATASETS_DIR/xwod_6cls_yolo/dataset.yaml" \
    --split val \
    --tile 640 \
    --overlap 0.2 \
    --nms-iou 0.6 \
    --device 0 \
    --batch 8 \
    --out "$EVALS_DIR/sahi_xwod_val"

# ── 7. Đo kích thước object ──────────────────────────────────────────────────
echo "=== 7. Object scale ==="
python scripts/analysis/measure_object_scale.py \
    --dataset acdc="$DATASETS_DIR/acdc_6cls_yolo/dataset.yaml" \
    --dataset xwod="$DATASETS_DIR/xwod_6cls_yolo/dataset.yaml" \
    --split val \
    --imgsz 640 \
    --out "$EVALS_DIR/object_scale"

# ── 8. In tóm tắt kết quả ────────────────────────────────────────────────────
echo ""
echo "=== KẾT QUẢ SAHI ACDC ==="
python3 -c "
import json
d = json.load(open('$EVALS_DIR/sahi_acdc_val/tiled_results.json'))
print(f'  mean F1 whole-image : {d[\"mean_f1_full\"]:.4f}')
print(f'  mean F1 tiled (SAHI): {d[\"mean_f1_tiled\"]:.4f}')
print(f'  delta               : {d[\"mean_f1_delta\"]:+.4f}')
print()
print('  Per class:')
for cls, v in d['per_class'].items():
    print(f'    {cls:12s} full={v[\"full\"][\"f1\"]:.4f}  tiled={v[\"tiled\"][\"f1\"]:.4f}  delta={v[\"tiled\"][\"f1\"]-v[\"full\"][\"f1\"]:+.4f}')
"

echo ""
echo "=== KẾT QUẢ SAHI XWOD ==="
python3 -c "
import json
d = json.load(open('$EVALS_DIR/sahi_xwod_val/tiled_results.json'))
print(f'  mean F1 whole-image : {d[\"mean_f1_full\"]:.4f}')
print(f'  mean F1 tiled (SAHI): {d[\"mean_f1_tiled\"]:.4f}')
print(f'  delta               : {d[\"mean_f1_delta\"]:+.4f}')
"

# ── 9. Upload kết quả lên Drive NGAY ────────────────────────────────────────
echo ""
echo "=== 9. Upload kết quả lên Drive ==="
DEST="$DRIVE_ROOT/evaluations/tiled_acdc/$(date +%Y%m%d)"
rclone copy "$EVALS_DIR/sahi_acdc_val"  "$DEST/sahi_acdc_val"  --progress
rclone copy "$EVALS_DIR/sahi_xwod_val"  "$DEST/sahi_xwod_val"  --progress
rclone copy "$EVALS_DIR/object_scale"   "$DEST/object_scale"   --progress
echo "  Saved to Drive: $DEST"
echo ""
echo "=== Hoàn tất ==="
