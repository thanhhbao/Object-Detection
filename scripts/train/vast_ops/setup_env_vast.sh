#!/usr/bin/env bash
# setup_env_vast.sh — Fresh VAST instance bootstrap
#
# Idempotent: safe to re-run if interrupted.
#
# Usage:
#   bash scripts/setup_env_vast.sh [--skip-drive] [--skip-datasets]
#
# Google Drive layout:
#   Account : 2200001341@nttu.edu.vn  (rclone remote: gdrive)
#   Root    : object_detection_in_adverse_weather/
#     official_backup_from_old_machine/
#       datasets_noleak_official.tar        ← 3.8 GB — xwod, acdc, bdd30k, dawn (all-in-one)
#       dataset_protocol.tar.gz             ← BDD pool protocol files
#     runs/
#       phase2/phase2_final_rtdetr/
#         phase2_final_rtdetr.tar.gz        ← 122 MB — Phase2 RT-DETR checkpoint
#     backups/
#       official_eval_20260901T180820Z.tar.gz

set -euo pipefail

# ── Config ────────────────────────────────────────────────────────────────────
REPO="${REPO:-/workspace/Object-Detection}"
REPO_URL="${REPO_URL:-https://github.com/thanhhbao/Object-Detection-In-Adverse.Weather-Conditions.git}"
GIT_BRANCH="${GIT_BRANCH:-feat/p2-a1-bdd-retrieval}"

PYTHON="${PYTHON:-python3}"
PIP="${PIP:-pip3}"

# Google Drive (rclone remote name: gdrive, account: 2200001341@nttu.edu.vn)
DRIVE_REMOTE="${DRIVE_REMOTE:-gdrive}"
DRIVE_ROOT="${DRIVE_ROOT:-object_detection_in_adverse_weather}"
DRIVE_BACKUP="${DRIVE_REMOTE}:${DRIVE_ROOT}/official_backup_from_old_machine"
DRIVE_RUNS="${DRIVE_REMOTE}:${DRIVE_ROOT}/runs"

# Local workspace
DS_NOLEAK="/workspace/datasets_noleak"
ZIP_CACHE="/workspace/datasets_zip"
RUNS="/workspace/runs"
PHASE2_CKPT="${RUNS}/phase2_final_rtdetr/weights/best.pt"

SKIP_DRIVE=0
SKIP_DATASETS=0
for arg in "$@"; do
  [[ "$arg" == "--skip-drive"    ]] && SKIP_DRIVE=1
  [[ "$arg" == "--skip-datasets" ]] && SKIP_DATASETS=1
done

echo "================================================================"
echo " VAST Environment Bootstrap"
echo " Repo:       ${REPO}  (${GIT_BRANCH})"
echo " Drive:      ${DRIVE_REMOTE}:${DRIVE_ROOT}/"
echo " Account:    2200001341@nttu.edu.vn"
echo "================================================================"
echo ""

# ── Step 1: System packages ───────────────────────────────────────────────────
echo "[1/7] System packages..."
apt-get update -qq
apt-get install -y -qq git curl wget unzip fuse3 tmux rclone \
  libgl1 libglib2.0-0 libsm6 libxext6 libxrender-dev ffmpeg build-essential \
  2>&1 | tail -3
echo "  Done."

# ── Step 2: Python packages ───────────────────────────────────────────────────
echo ""
echo "[2/7] Python packages..."

CUDA_VER=$(nvidia-smi 2>/dev/null | grep -oP "CUDA Version: \K[\d\.]+" | head -1 || echo "")
if [[ "$CUDA_VER" == 12* ]]; then
  TORCH_INDEX="https://download.pytorch.org/whl/cu121"
elif [[ "$CUDA_VER" == 11* ]]; then
  TORCH_INDEX="https://download.pytorch.org/whl/cu118"
else
  TORCH_INDEX="https://download.pytorch.org/whl/cpu"
  echo "  WARNING: CUDA not detected."
fi
echo "  CUDA=${CUDA_VER:-none} → ${TORCH_INDEX}"

$PIP install --quiet torch torchvision --index-url "${TORCH_INDEX}"
$PIP install --quiet \
  "ultralytics==8.3.*" \
  opencv-python-headless \
  numpy pandas matplotlib Pillow pyyaml tqdm \
  scikit-learn albumentations pytest transformers

$PYTHON -c "
import torch
print(f'  torch {torch.__version__}  cuda={torch.cuda.is_available()}  gpus={torch.cuda.device_count()}')
"
echo "  Packages: OK"

# ── Step 3: Clone / update repo ───────────────────────────────────────────────
echo ""
echo "[3/7] Repo..."
if [ ! -d "${REPO}/.git" ]; then
  git clone --branch "${GIT_BRANCH}" "${REPO_URL}" "${REPO}"
else
  cd "${REPO}"
  git fetch origin
  git checkout "${GIT_BRANCH}"
  git pull origin "${GIT_BRANCH}"
  cd -
fi
echo "  $(git -C "${REPO}" log -1 --oneline)"

export PYTHONPATH="${PYTHONPATH:-}:${REPO}/src"
grep -qF "Object-Detection/src" ~/.bashrc 2>/dev/null \
  || echo "export PYTHONPATH=\"\${PYTHONPATH:-}:${REPO}/src\"" >> ~/.bashrc

# ── Step 4: Protocol tests ────────────────────────────────────────────────────
echo ""
echo "[4/7] Protocol tests..."
$PYTHON -m pytest -q \
  "${REPO}/tests/test_retrieve_dinov2.py" \
  "${REPO}/tests/test_build_random_rare_control.py" \
  2>&1 | tee /tmp/setup_env_pytest.log || true

if grep -qE "^FAILED|^ERROR" /tmp/setup_env_pytest.log; then
  echo "  WARNING: some tests failed — see /tmp/setup_env_pytest.log"
else
  PASSED=$(grep -oP "\d+ passed" /tmp/setup_env_pytest.log | head -1 || echo "")
  echo "  Tests: PASS  ${PASSED}"
fi

# ── Step 5: Download from Google Drive ───────────────────────────────────────
echo ""
echo "[5/7] Google Drive downloads (${DRIVE_REMOTE}:${DRIVE_ROOT}/)..."

if [ "$SKIP_DRIVE" -eq 1 ]; then
  echo "  Skipped (--skip-drive)"
else
  if ! rclone lsd "${DRIVE_REMOTE}:" &>/dev/null; then
    echo ""
    echo "  ERROR: rclone remote '${DRIVE_REMOTE}' not configured."
    echo "  Run interactively:"
    echo "    rclone config"
    echo "  → n → name: gdrive → type: drive → scope: 1 → no auto config"
    echo "  → copy auth link → open on local machine → paste token"
    echo ""
    echo "  Then re-run this script, or use --skip-drive."
    echo ""
  else
    mkdir -p "${DS_NOLEAK}" "${ZIP_CACHE}" "${RUNS}"

    # ── 5a. Datasets (all-in-one tar: xwod, acdc, bdd30k, dawn) ──────────
    DATASETS_TAR="${ZIP_CACHE}/datasets_noleak_official.tar"
    if [ ! -d "${DS_NOLEAK}/xwod_6cls_yolo" ] || [ ! -d "${DS_NOLEAK}/acdc_6cls_yolo" ]; then
      if [ ! -f "${DATASETS_TAR}" ]; then
        echo "  Downloading datasets_noleak_official.tar (3.8 GB)..."
        rclone copy \
          "${DRIVE_BACKUP}/datasets_noleak_official.tar" \
          "${ZIP_CACHE}/" --progress
      fi
      echo "  Extracting datasets_noleak_official.tar → ${DS_NOLEAK}/ ..."
      tar -xf "${DATASETS_TAR}" -C /workspace/
      echo "  Datasets: Done"
    else
      echo "  Datasets already extracted — skipping."
    fi

    # ── 5b. Phase2 RT-DETR checkpoint ─────────────────────────────────────
    if [ ! -f "${PHASE2_CKPT}" ]; then
      CKPT_TAR="${ZIP_CACHE}/phase2_final_rtdetr.tar.gz"
      if [ ! -f "${CKPT_TAR}" ]; then
        echo "  Downloading phase2_final_rtdetr.tar.gz (122 MB)..."
        rclone copy \
          "${DRIVE_RUNS}/phase2/phase2_final_rtdetr/phase2_final_rtdetr.tar.gz" \
          "${ZIP_CACHE}/" --progress
      fi
      echo "  Extracting phase2_final_rtdetr.tar.gz → ${RUNS}/ ..."
      tar -xzf "${CKPT_TAR}" -C "${RUNS}/"
      # Verify expected path
      if [ ! -f "${PHASE2_CKPT}" ]; then
        echo "  WARNING: ${PHASE2_CKPT} not found after extraction."
        echo "  Check tar contents: tar -tzf ${CKPT_TAR} | head -20"
      else
        echo "  Phase2 checkpoint: OK"
      fi
    else
      SIZE=$(du -sh "${PHASE2_CKPT}" | cut -f1)
      echo "  Phase2 checkpoint already exists (${SIZE}) — skipping."
    fi

    echo "  Drive downloads: Done"
  fi
fi

# ── Step 6: Verify datasets ───────────────────────────────────────────────────
echo ""
echo "[6/7] Dataset verification..."
ALL_OK=1
for ds in \
  "${DS_NOLEAK}/xwod_6cls_yolo" \
  "${DS_NOLEAK}/acdc_6cls_yolo" \
  "${DS_NOLEAK}/bdd100k_6cls_yolo"
do
  if [ -d "${ds}" ]; then
    CNT=$(find "${ds}/images/train" \( -type f -o -type l \) 2>/dev/null | wc -l | tr -d ' ')
    echo "  OK      $(basename "${ds}")  (train: ${CNT})"
  else
    echo "  MISSING ${ds}"
    ALL_OK=0
  fi
done

if [ -f "${PHASE2_CKPT}" ]; then
  SIZE=$(du -sh "${PHASE2_CKPT}" | cut -f1)
  echo "  OK      phase2_final_rtdetr/weights/best.pt  (${SIZE})"
else
  echo "  MISSING ${PHASE2_CKPT}"
  ALL_OK=0
fi

# ── Step 7: BDD pool reminder ─────────────────────────────────────────────────
echo ""
echo "[7/7] BDD remaining pool..."
POOL_STATS="/workspace/datasets_noleak/bdd_remaining_pool/pool_stats.json"
if [ -f "${POOL_STATS}" ]; then
  POOL_N=$($PYTHON -c "import json; print(json.load(open('${POOL_STATS}'))['candidate_pool'])" 2>/dev/null || echo "?")
  echo "  Pool exists: ${POOL_N} candidates — skipping rebuild."
else
  echo "  Pool not found — will be built by setup_p2_a1_vast.sh."
  echo "  (Requires full BDD100K raw or bdd100k_6cls_full_yolo)"
fi

# ── Summary ───────────────────────────────────────────────────────────────────
echo ""
echo "================================================================"
echo " Bootstrap complete."
echo " Drive: ${DRIVE_REMOTE}:${DRIVE_ROOT}/"
echo ""
if [ "$ALL_OK" -eq 1 ]; then
  echo " Status: READY"
  echo ""
  echo "   # Build BDD remaining pool (first time only):"
  echo "   bash ${REPO}/scripts/setup_p2_a1_vast.sh"
  echo ""
  echo "   # Build retrieval arms + merge + audit:"
  echo "   bash ${REPO}/scripts/setup_ablation_vast.sh"
  echo ""
  echo "   # Train A1-DINO v2:"
  echo "   bash ${REPO}/scripts/train_p2_a1_vast.sh a1_dino"
else
  echo " Status: INCOMPLETE — check missing items above."
fi
echo "================================================================"
