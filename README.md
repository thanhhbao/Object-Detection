# People and Vehicle Detection in Adverse Weather Conditions

Source code for the undergraduate thesis *"Thiết kế và phát triển mô hình học sâu
phát hiện người và phương tiện trong điều kiện thời tiết bất lợi phục vụ giám sát
giao thông thông minh"* (Nguyen Tat Thanh University, 2026).

The repository contains a progressive multi-domain fine-tuning pipeline, a
four-architecture comparison, and two controlled data-selection ablations.

Target classes:

```text
0 person   1 bicycle   2 car   3 motorcycle   4 bus   5 truck
```

## Research pipeline

```text
COCO pretrained
      │
      ▼  Stage 1 — driving-domain adaptation
   BDD100K train (30 000 images)
      │
      ▼  Stage 2 — adverse-weather specialisation
   XWOD train (6 006 images)
      │
      ▼  Phase 2 — simultaneous multi-domain fine-tuning
   XWOD + ACDC + BDD100K (37 188 base → 50 985 after rare-class oversampling)
      │
      ├──▶ A0R      + 5 000 randomly sampled rare-class images
      └──▶ A1-DINO  + 5 000 images retrieved by DINOv2 similarity
```

All configurations are evaluated on four frozen test sets: **XWOD** (target
domain), **BDD100K** (source domain, measures catastrophic forgetting), **DAWN**
(out-of-domain, never used for training at any stage), and **ACDC**.

Four architectures are compared under an identical protocol at Stage 1 and
Stage 2: YOLOv8n, YOLO11n, Faster R-CNN and RT-DETR-L. RT-DETR-L achieved the
highest mAP50-95 on XWOD validation and was carried forward to Phase 2.

## Evaluation protocol

The separation between validation and test data is the constraint the whole
study rests on:

- **validation** drives early stopping, checkpoint selection and the choice of
  architecture;
- **test** is touched only after a configuration is frozen;
- **DAWN** never participates in training, not even in the candidate-selection
  step of the ablations;
- BDD100K validation and test never enter the retrieval candidate pool.

`tests/` enforces these invariants. The suite fails loudly rather than silently
continuing when an overlap is detected.

## Project structure

```text
├── configs/
│   ├── ultralytics/    Stage 1, Stage 2, Phase 2 and ablation configs
│   ├── torchvision/    Faster R-CNN configs
│   ├── ablation/       Preprocessing / architecture variants
│   ├── eval/           Evaluation-only dataset configs
│   └── common/         Shared class names, paths and training defaults
├── scripts/            Data preparation, training, evaluation, analysis
├── src/dawn_ablation/  Shared data-preparation and evaluation utilities
├── models/             Custom YOLO architecture definitions
├── tests/              Leak-protocol and dataset-builder tests
├── demo/               Streamlit inference demo
├── docs/               Experiment protocols and result reports
├── datasets/           Local datasets (git-ignored)
└── runs/               Training outputs (git-ignored)
```

Datasets, checkpoints and training outputs are deliberately not tracked. See the
thesis appendix for the archive containing them.

## Environment

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Experiments were run on a single NVIDIA RTX 5090 with PyTorch 2.7.1 + CUDA 12.8
and Ultralytics 8.3.159. Library versions matter: the same checkpoint evaluated
under Ultralytics 8.3.159 and 8.4.142 differs by up to 0.0147 mAP50-95, which is
larger than the gap between the configurations being compared. Pin the version
when reproducing.

## Configs

Experiment configs stay small; shared values live in `configs/common/`.

```yaml
# configs/ultralytics/stage2_xwod_rtdetr_from_bdd30k.yaml
defaults: configs/common/train_defaults.yaml
paths: configs/common/paths_colab.yaml

from_run: stage1_bdd30k_rtdetr    # resolves to <project>/<from_run>/weights/best.pt
dataset: xwod
name: stage2_xwod_rtdetr_from_bdd30k

lr0: 0.0005
patience: 15
```

Use `configs/common/paths_vast.yaml` instead of `paths_colab.yaml` when running
on a Vast.ai instance, or point `OD_PATHS` at either file.

## Reproducing the pipeline

**1. Prepare the datasets** — each script normalises one source into the shared
six-class YOLO layout:

```bash
python scripts/prepare_bdd100k_yolo.py --src <raw> --dst <out> --subset-size 30000
python scripts/prepare_xwod.py         --raw-dir <raw> --output-dir <out>
python scripts/prepare_acdc.py         --raw-dir <raw> --output-dir <out>
python scripts/prepare_dawn.py         --raw-dir <raw> --output-dir <out>
```

`convert_bdd100k_json_to_yolo.py` handles the official BDD100K release, which
ships JSON annotations instead of pre-converted YOLO labels.

**2. Train Stage 1 and Stage 2:**

```bash
python scripts/train_ultralytics.py --config configs/ultralytics/stage1_bdd30k_rtdetr.yaml
python scripts/train_ultralytics.py --config configs/ultralytics/stage2_xwod_rtdetr_from_bdd30k.yaml
```

**3. Build the Phase 2 dataset and train:**

```bash
python scripts/build_phase2_dataset.py --xwod-root <…> --acdc-root <…> --bdd-root <…> --out-root <…>
python scripts/train_ultralytics.py --config configs/ultralytics/phase2_final_rtdetr_from_xwod.yaml
```

**4. Run the two data-selection ablations.** `setup_ablation_vast.sh` chains the
whole sequence: protocol tests, candidate-pool construction, both selection
strategies, dataset merging and a strict invariant audit.

```bash
bash scripts/setup_ablation_vast.sh
bash scripts/train_p2_a1_vast.sh a0r
bash scripts/train_p2_a1_vast.sh a1_dino
```

**5. Evaluate:**

```bash
python scripts/evaluate.py --config <config> --weights <ckpt> --data <yaml> --split test
python scripts/eval_f1_matrix.py --arm name=<ckpt> --dataset name=<yaml> --split val
```

## Analysis tools

Beyond training and evaluation, several scripts support the diagnostic analysis
reported in Chapter 4:

| Script | Purpose |
|---|---|
| `eval_f1_matrix.py` | Per-class F1 across arms and datasets, with a val/test agreement check |
| `diagnose_weak_classes.py` | Why a class fails: by condition, by object size, blind miss vs class confusion |
| `measure_object_scale.py` | Object size at model input, comparable across datasets stored at different resolutions |
| `tune_conf_thresholds.py` | Per-class confidence thresholds, fitted on validation and verified on test |
| `eval_tiled_inference.py` | Sliced (SAHI-style) inference compared against whole-image inference |
| `report_f1_ablation.py` | Generates the F1 report from the official metrics archive |

## Attention-block ablation

An earlier, separate line of work inserts a single attention block into YOLOv8n
to measure the effect of one lightweight architectural change. It is not part of
the results reported in the thesis, but the code is kept because it is complete
and documented: `docs/ABLATION_SE.md`, `configs/yolov8n_se.yaml`,
`models/yolov8n_cbam_neck.yaml` and `src/dawn_ablation/attention.py`.

`register_custom_modules()` registers those blocks with Ultralytics and is
called by every evaluation script, so that checkpoints containing them load
correctly regardless of which experiment produced the weights.

## Retrieval code layout

`scripts/active_retrieval.py` plays two roles and is easy to misread as a
duplicate of `retrieve_dinov2.py`:

- as a **library**, it provides the shared retrieval infrastructure — label
  validation, class filtering of the candidate pool, ground-truth-aware hard
  mining, embedding hooks and pool caching. Both ablation arms import from it,
  as do `preflight_p2_a1.py` and three test modules;
- as a **script**, it implements the first retrieval variant, which ranks
  candidates by whole-image embedding similarity under a global Top-K budget.

`scripts/retrieve_dinov2.py` implements the variant reported in the thesis:
DINOv2 features taken from ground-truth object crops, a rare-class density
factor in the ranking score, and a per-class budget filled scarcest class first.
It reuses the library half of `active_retrieval.py` rather than duplicating it.

## Figure and utility scripts

Run manually rather than from the pipeline. They regenerate artefacts used in
the thesis or support inspection during development:

| Script | Purpose |
|---|---|
| `draw_pipeline.py` | Renders the training-pipeline diagram (Figure 3.2) |
| `draw_rtdetr_pipeline.py` | Renders the RT-DETR architecture diagram |
| `visualize_dataset_samples.py` | Contact sheet of images and boxes from a dataset split |
| `plot_torchvision_history.py` | Training curves for the Faster R-CNN runs |
| `evaluate_torchvision_by_weather.py` | Per-weather evaluation for Faster R-CNN, mirrors `evaluate_by_weather.py` |
| `error_analysis.py` | Per-image failure inspection |
| `copy_paste_augment.py` | Copy-paste augmentation for rare classes; implemented but not used in the reported experiments |
| `augment_weather.py`, `augment_dark.py` | Synthetic weather and low-light augmentation, used only in `configs/ablation/` |
| `sanity_se.py`, `smoke_p2_a1_embedding.py` | Smoke checks for the custom modules and the embedding step |

## Results

Stage 1 → Phase 2, mAP50-95 on the frozen test sets (RT-DETR-L):

| Test set | Stage 1 | Phase 2 | Change |
|---|---:|---:|---:|
| XWOD | 0.291 | 0.510 | +75 % |
| DAWN | 0.344 | 0.515 | +50 % |
| ACDC | 0.174 | 0.247 | +42 % |
| BDD100K | 0.362 | 0.354 | −2 % |

The BDD100K drop is the cost of holding several domains in one set of weights;
Phase 2 recovers about 96 % of what sequential fine-tuning lost at Stage 2.

Neither data-selection strategy produced a difference beyond run-to-run noise.
`docs/PHASE2_SCALE_DIAGNOSIS_20260907.md` traces this to object scale rather
than sample count: recall reaches 1.000 for large objects across all six classes
while falling to 0.00–0.70 for small ones, and both strategies act on sample
count instead.

Detailed reports: `docs/RESULTS.md`, `docs/f1_ablation_report.md`,
`docs/phase2_protocol.md`, `docs/p2_a1_bdd_retrieval_protocol.md`.

## Tests

```bash
python -m pytest tests/ -q
```

122 tests cover the leak protocol, dataset builders and the retrieval pipeline.

## Demo

```bash
cd demo && pip install -r requirements.txt && streamlit run app.py
```

## References

- BDD100K — https://arxiv.org/abs/1805.04687
- DAWN — https://arxiv.org/abs/2008.05402
- ACDC — https://arxiv.org/abs/2104.13395
- RT-DETR — https://arxiv.org/abs/2304.08069
- DINOv2 — https://arxiv.org/abs/2304.07193
- Ultralytics — https://github.com/ultralytics/ultralytics
