"""Aperture — object detection dashboard (image + video)."""

from __future__ import annotations

import io
import json
import os
import sys
import tempfile
from collections import Counter
from pathlib import Path

import hashlib

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image
from streamlit_paste_button import paste_image_button

sys.path.insert(0, str(Path(__file__).parent))
from utils.inference import (
    load_model,
    list_local_models,
    detect_image_boxes,
    process_video_to_file,
    download_video_from_url,
)
from utils.visualize import (
    CLASS_COLORS_HEX, CLASSES, resize_for_display,
    draw_detections, pil_to_bgr, bgr_to_pil,
)

st.set_page_config(page_title="Aperture · Object Detection",
                   page_icon="◉", layout="wide", initial_sidebar_state="collapsed")

css = (Path(__file__).parent / "assets" / "style.css").read_text()
st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

CLASS_LABELS = [c.capitalize() for c in CLASSES]

MODELS = list_local_models()
_model_names = list(MODELS)
_NO_MODEL = "— Không —"


def _pick(tokens: tuple[str, ...], fallback: str) -> str:
    for n in _model_names:
        if all(t in n.lower() for t in tokens):
            return n
    return fallback


def _model_label(name: str) -> str:
    lo = name.lower()
    if "a0r" in lo or "rare" in lo:
        return "RT-DETR-L A0R"
    if ("rtdetr" in lo or "rt_detr" in lo or "rt-detr" in lo) and ("coco" in lo or "baseline" in lo):
        return "RT-DETR-L COCO Baseline"
    if "rtdetr" in lo or "rt_detr" in lo or "rt-detr" in lo:
        return "RT-DETR-L Phase 2"
    if "coco" in lo and "baseline" in lo:
        return "YOLOv8n COCO Baseline"
    if "v3_960" in lo or "phase2" in lo:
        return "YOLOv8n Phase 2"
    if "v1" in lo or "stage1" in lo:
        return "YOLOv8n Stage 1"
    return name.replace("_", " ")

# ── Session init ──────────────────────────────────────────────────────────────
st.session_state.setdefault("recents", [])       # list of {name, kind, bytes, meta}
st.session_state.setdefault("active_name", None)
st.session_state.setdefault("last_upload", None)
st.session_state.setdefault("stop_requested", False)
st.session_state.setdefault("is_running", False)


def upsert_recent(name: str, kind: str, data: bytes, meta: str) -> None:
    recents = [r for r in st.session_state["recents"] if r["name"] != name]
    recents.insert(0, {"name": name, "kind": kind, "bytes": data, "meta": meta})
    st.session_state["recents"] = recents[:5]


def get_active() -> dict | None:
    for r in st.session_state["recents"]:
        if r["name"] == st.session_state["active_name"]:
            return r
    return None


def donut(pct: int) -> str:
    return (f'<div class="donut" style="background:conic-gradient(var(--accent) {pct}%, '
            f'var(--line) 0);"><span>{pct}%</span></div>')


def object_rows(items: list[dict]) -> str:
    html = ""
    for it in items:
        html += (
            f'<div class="obj-row"><div class="obj-top">'
            f'<span class="obj-name"><span class="obj-dot" style="background:{it["color"]}"></span>'
            f'{it["label"]}</span><span class="obj-conf">{it["right"]}</span></div>'
            f'<div class="obj-bar"><span style="width:{it["pct"]}%;background:{it["color"]}"></span></div></div>'
        )
    return html


# ═══════════════════════════════════════════════════════════════════════════════
#  TOOLBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    tb_brand, tb_mode, tb_actions = st.columns([2.2, 1.5, 6.3], vertical_alignment="center")
    with tb_brand:
        st.markdown("""
        <div class="brand">
            <div class="brand-logo">◉</div>
            <div>
                <div class="brand-title">Aperture</div>
                <div class="brand-sub">Adverse Weather · KLTN NTT</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with tb_mode:
        mode = st.segmented_control("mode", ["Image", "Video"], default="Image",
                                    label_visibility="collapsed")
        mode = mode or "Image"
    with tb_actions:
        a1, a2, a3 = st.columns([3.2, 3.2, 1.6], vertical_alignment="center")
        with a1:
            baseline_name = st.selectbox(
                "Baseline", _model_names,
                index=_model_names.index(_pick(("coco", "baseline"), _model_names[0])),
                format_func=_model_label, label_visibility="collapsed",
            )
        with a2:
            _final_opts = [_NO_MODEL] + _model_names
            final_name = st.selectbox(
                "Final", _final_opts,
                index=_final_opts.index(_pick(("v3_960",), _model_names[-1])) if _pick(("v3_960",), _model_names[-1]) in _final_opts else 0,
                format_func=lambda n: "Không (1 model)" if n == _NO_MODEL else _model_label(n),
                label_visibility="collapsed",
            )
        with a3:
            if st.session_state.get("is_running"):
                stop_btn = st.button("Stop", key="stop_btn", use_container_width=True)
                if stop_btn:
                    st.session_state["stop_requested"] = True
                run = False
            else:
                run = st.button("Run detection", key="run_btn", use_container_width=True)
                st.session_state["stop_requested"] = False

KIND = "image" if mode == "Image" else "video"

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN — 3 columns
# ═══════════════════════════════════════════════════════════════════════════════
left, center, right = st.columns([1.15, 2.3, 1.35])

# ── LEFT: input source ────────────────────────────────────────────────────────
with left:
    with st.container(border=True):
        st.markdown('<div class="panel-label">Input source</div>', unsafe_allow_html=True)

        if KIND == "image":
            up = st.file_uploader("Drag & drop an image, or click to browse",
                                  type=["jpg", "jpeg", "png", "webp", "bmp"],
                                  label_visibility="collapsed", key="img_up")
        else:
            up = st.file_uploader("Drag & drop a video, or click to browse",
                                  type=["mp4", "avi", "mov", "mkv"],
                                  label_visibility="collapsed", key="vid_up")

        if up is not None and st.session_state["last_upload"] != (KIND, up.name, up.size):
            data = up.getvalue()
            if KIND == "image":
                im = Image.open(io.BytesIO(data))
                meta = f"{im.width}×{im.height} · {up.size/1e6:.1f}MB"
            else:
                meta = f"{up.size/1e6:.1f}MB"
            upsert_recent(up.name, KIND, data, meta)
            st.session_state["active_name"] = up.name
            st.session_state["last_upload"] = (KIND, up.name, up.size)
            st.session_state.pop("img_result", None)
            st.session_state.pop("vid_result", None)

        # Paste image from clipboard (screenshots)
        if KIND == "image":
            pasted = paste_image_button("Paste from clipboard", key="paste_btn",
                                        background_color="#111827",
                                        hover_background_color="#000000",
                                        text_color="#ffffff", errors="ignore")
            if pasted is not None and pasted.image_data is not None:
                pimg = pasted.image_data.convert("RGB")
                buf = io.BytesIO(); pimg.save(buf, format="PNG")
                pbytes = buf.getvalue()
                digest = hashlib.md5(pbytes).hexdigest()[:8]
                if st.session_state.get("last_paste") != digest:
                    name = f"pasted-{digest}.png"
                    upsert_recent(name, "image", pbytes,
                                  f"{pimg.width}×{pimg.height} · pasted")
                    st.session_state["active_name"] = name
                    st.session_state["last_paste"] = digest
                    st.session_state.pop("img_result", None)
                    st.rerun()

        # Import video from URL
        if KIND == "video":
            with st.form("url_form", clear_on_submit=True):
                url = st.text_input("Video URL",
                                    placeholder="https://…/clip.mp4 or a YouTube link",
                                    label_visibility="collapsed")
                fetch = st.form_submit_button("Fetch", use_container_width=True)
            if fetch and url.strip():
                try:
                    with st.spinner("Downloading…"):
                        fp = download_video_from_url(url, max_mb=200)
                    with open(fp, "rb") as f:
                        vdata = f.read()
                    os.remove(fp)
                    vname = Path(url).name or "video.mp4"
                    upsert_recent(vname, "video", vdata, f"{len(vdata)/1e6:.1f}MB · url")
                    st.session_state["active_name"] = vname
                    st.session_state.pop("vid_result", None)
                    st.rerun()
                except RuntimeError as exc:
                    st.error(str(exc))

        # Recent files (matching current mode)
        recents = [r for r in st.session_state["recents"] if r["kind"] == KIND]
        if recents:
            st.markdown('<div class="panel-label" style="margin-top:0.9rem;">Recent files</div>',
                        unsafe_allow_html=True)
            for i, r in enumerate(recents):
                is_active = r["name"] == st.session_state["active_name"]
                label = f"{'Active — ' if is_active else ''}{r['name']}"
                if st.button(label, key=f"rec_{i}", use_container_width=True):
                    st.session_state["active_name"] = r["name"]
                    st.session_state.pop("img_result", None)
                    st.session_state.pop("vid_result", None)
                    st.rerun()
                st.caption(r["meta"])

        # Confidence threshold
        st.markdown("<div style='margin-top:0.9rem;'></div>", unsafe_allow_html=True)
        conf_pct = st.slider("Confidence threshold", 10, 95, 30, 1, format="%d%%")
        conf = conf_pct / 100.0

        with st.expander("Advanced"):
            iou = st.slider("IoU (NMS)", 0.10, 0.90, 0.45, 0.05)
            sel_labels = st.multiselect("Classes", CLASS_LABELS, default=CLASS_LABELS)
            sel_ids = [CLASS_LABELS.index(l) for l in sel_labels]
            sel_set = set(sel_ids) if 0 < len(sel_ids) < len(CLASSES) else set(range(len(CLASSES)))

    if KIND == "video":
        with st.container(border=True):
            st.markdown('<div class="panel-label">Duration</div>', unsafe_allow_html=True)
            DUR = {"10s": 10.0, "30s": 30.0, "60s": 60.0, "Entire": None}
            dur_label = st.radio("Duration", list(DUR.keys()), index=1,
                                 horizontal=True, label_visibility="collapsed")
            max_seconds = DUR[dur_label]

active = get_active()

# ── Run detection ─────────────────────────────────────────────────────────────
_single_mode = (final_name == _NO_MODEL)

if run and active and active["kind"] == "image":
    img = Image.open(io.BytesIO(active["bytes"])).convert("RGB")
    prog = st.progress(0, text="Running…")
    base_ms, base_boxes = detect_image_boxes(load_model(MODELS[baseline_name]), img, iou=iou)
    if _single_mode:
        final_ms, final_boxes = None, None
    else:
        prog.progress(50, text="Running Final model…")
        final_ms, final_boxes = detect_image_boxes(load_model(MODELS[final_name]), img, iou=iou)
    prog.progress(100, text="Done")
    prog.empty()
    st.session_state["img_result"] = {
        "name": active["name"],
        "baseline_boxes": base_boxes, "baseline_ms": base_ms,
        "final_boxes": final_boxes, "final_ms": final_ms,
        "baseline_name": baseline_name, "final_name": final_name,
        "w": img.width, "h": img.height,
    }
    st.session_state["scroll_to_results"] = True

do_video_run = bool(run and active and active["kind"] == "video")
if do_video_run:
    st.session_state["is_running"] = True

# ── Build current results ─────────────────────────────────────────────────────
base_summary = final_summary = None
base_annotated = final_annotated = None
obj_items: list[dict] = []   # kept for video compat in bottom section

img_res = st.session_state.get("img_result")
vid_res = st.session_state.get("vid_result")


def _img_summary(boxes: list, elapsed_ms: float) -> dict:
    counts = Counter(CLASSES[b[0]] for b in boxes)
    confs = [b[1] for b in boxes]
    return {
        "avg": sum(confs) / len(confs) if confs else 0.0,
        "total": len(boxes), "n_classes": len(counts),
        "time_ms": elapsed_ms, "counts": dict(counts), "confidences": confs,
    }


if KIND == "image" and active and img_res and img_res["name"] == active["name"]:
    src_img = Image.open(io.BytesIO(active["bytes"])).convert("RGB")
    base_boxes = [b for b in img_res["baseline_boxes"] if b[1] >= conf and b[0] in sel_set]
    base_annotated = bgr_to_pil(draw_detections(pil_to_bgr(src_img), base_boxes, conf_threshold=0))
    base_summary = _img_summary(base_boxes, img_res["baseline_ms"])
    if img_res.get("final_boxes") is not None:
        final_boxes = [b for b in img_res["final_boxes"] if b[1] >= conf and b[0] in sel_set]
        final_annotated = bgr_to_pil(draw_detections(pil_to_bgr(src_img), final_boxes, conf_threshold=0))
        final_summary = _img_summary(final_boxes, img_res["final_ms"])
    else:
        final_boxes = []
        final_annotated = None
        final_summary = None

elif KIND == "video" and active and vid_res and vid_res["name"] == active["name"]:
    def _vid_s(stats: dict) -> dict:
        return {
            "avg": stats["conf_mean"], "total": stats["total_detections"],
            "n_classes": len(stats["class_counts"]), "time_ms": stats["elapsed_s"] * 1000,
            "counts": stats["class_counts"], "confidences": stats.get("confidences", []),
        }
    base_summary = _vid_s(vid_res["baseline_stats"])
    final_summary = _vid_s(vid_res["final_stats"])

# summary alias used by the bottom section
summary = final_summary

# ── CENTER: side-by-side preview ─────────────────────────────────────────────
with center:
    with st.container(border=True):
        st.markdown(f"""
        <div class="canvas-head">
            <div class="canvas-status">{"Static image mode" if KIND == "image" else "Video mode"} · comparison</div>
            <div class="canvas-zoom">{active["name"] if active else "no source"}</div>
        </div>
        """, unsafe_allow_html=True)

        if not active:
            st.markdown("""
            <div class="empty-state">
                <div class="empty-state-icon">◇</div>
                <div class="empty-state-title">No source selected</div>
                <div class="empty-state-hint">Upload a file on the left to begin</div>
            </div>
            """, unsafe_allow_html=True)
        elif KIND == "image":
            bl = _model_label(img_res["baseline_name"]) if img_res else _model_label(baseline_name)
            res_single = img_res and img_res.get("final_boxes") is None
            if res_single or _single_mode and not img_res:
                # single model — full width
                st.markdown(f'<div class="cmp-label"><span class="dot" style="background:#64748b"></span><span class="name">MODEL</span> · {bl}</div>', unsafe_allow_html=True)
                disp = base_annotated or Image.open(io.BytesIO(active["bytes"])).convert("RGB")
                st.image(resize_for_display(disp, 900), use_container_width=True)
            else:
                fl = _model_label(img_res["final_name"]) if img_res else _model_label(final_name)
                col_b, col_f = st.columns(2, gap="small")
                with col_b:
                    st.markdown(f'<div class="cmp-label"><span class="dot" style="background:#64748b"></span><span class="name">BASELINE</span> · {bl}</div>', unsafe_allow_html=True)
                    if base_annotated:
                        st.image(resize_for_display(base_annotated, 600), use_container_width=True)
                    else:
                        st.image(resize_for_display(Image.open(io.BytesIO(active["bytes"])).convert("RGB"), 600),
                                 use_container_width=True)
                with col_f:
                    st.markdown(f'<div class="cmp-label"><span class="dot" style="background:#0ea5e9"></span><span class="name">FINAL</span> · {fl}</div>', unsafe_allow_html=True)
                    if final_annotated:
                        st.image(resize_for_display(final_annotated, 600), use_container_width=True)
                    else:
                        st.image(resize_for_display(Image.open(io.BytesIO(active["bytes"])).convert("RGB"), 600),
                                 use_container_width=True)
        else:  # video
            if do_video_run:
                st.video(active["bytes"])
                prog = st.progress(0, text="Baseline · processing…")

                def _cb_b(idx, total):
                    prog.progress(min(0.45 * (idx + 1) / max(total, 1), 0.45),
                                  text=f"Baseline · frame {idx+1}/{total}")

                tmp = tempfile.NamedTemporaryFile(suffix=".mp4", delete=False)
                tmp.write(active["bytes"]); tmp.flush()
                base_path, base_stats = process_video_to_file(
                    load_model(MODELS[baseline_name]), tmp.name, conf=conf, iou=iou,
                    classes=(list(sel_set) if len(sel_set) < len(CLASSES) else None),
                    max_seconds=max_seconds, progress_callback=_cb_b,
                )
                if st.session_state.get("stop_requested"):
                    prog.empty()
                    st.session_state["is_running"] = False
                    st.session_state["stop_requested"] = False
                    st.warning("Đã dừng sau Baseline.")
                    st.stop()

                prog.progress(50, text="Final model · processing…")

                def _cb_f(idx, total):
                    prog.progress(min(0.5 + 0.5 * (idx + 1) / max(total, 1), 1.0),
                                  text=f"Final model · frame {idx+1}/{total}")

                final_path, final_stats = process_video_to_file(
                    load_model(MODELS[final_name]), tmp.name, conf=conf, iou=iou,
                    classes=(list(sel_set) if len(sel_set) < len(CLASSES) else None),
                    max_seconds=max_seconds, progress_callback=_cb_f,
                )
                os.remove(tmp.name)
                with open(base_path, "rb") as f:
                    base_vid = f.read()
                with open(final_path, "rb") as f:
                    final_vid = f.read()
                os.remove(base_path); os.remove(final_path)
                prog.empty()
                st.session_state["vid_result"] = {
                    "name": active["name"],
                    "baseline_bytes": base_vid, "baseline_stats": base_stats,
                    "baseline_name": baseline_name,
                    "final_bytes": final_vid, "final_stats": final_stats,
                    "final_name": final_name,
                }
                st.session_state["scroll_to_results"] = True
                st.session_state["is_running"] = False
                st.rerun()
            elif vid_res and vid_res["name"] == active["name"]:
                col_b, col_f = st.columns(2, gap="small")
                with col_b:
                    st.markdown(f'<div class="cmp-label"><span class="dot" style="background:#64748b"></span><span class="name">BASELINE</span> · {_model_label(vid_res["baseline_name"])}</div>', unsafe_allow_html=True)
                    st.video(vid_res["baseline_bytes"])
                with col_f:
                    st.markdown(f'<div class="cmp-label"><span class="dot" style="background:#0ea5e9"></span><span class="name">FINAL</span> · {_model_label(vid_res["final_name"])}</div>', unsafe_allow_html=True)
                    st.video(vid_res["final_bytes"])
            else:
                st.video(active["bytes"])

# ── RIGHT: comparison table ───────────────────────────────────────────────────
with right:
    with st.container(border=True):
        has_base = base_summary is not None
        has_both = has_base and final_summary is not None
        n_base = base_summary["total"] if base_summary else 0
        n_final = final_summary["total"] if final_summary else 0
        label = "Kết quả" if not has_both else "Comparison"
        count_txt = str(n_base) if not has_both else f"{n_base} → {n_final}"
        st.markdown(f'<div class="panel-label">{label} <span class="count">{count_txt}</span></div>',
                    unsafe_allow_html=True)

        if not has_base:
            st.markdown("""
            <div class="empty-state">
                <div class="empty-state-icon">◎</div>
                <div class="empty-state-title">Nothing yet</div>
                <div class="empty-state-hint">Click "Run detection" to analyze this file</div>
            </div>
            """, unsafe_allow_html=True)
        elif not has_both:
            # single model
            rows = [{"Lớp": c.capitalize(), "Số lượng": int(base_summary["counts"].get(c, 0))}
                    for c in CLASSES]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                         column_config={
                             "Lớp": st.column_config.TextColumn("Lớp"),
                             "Số lượng": st.column_config.NumberColumn("Số lượng", format="%d"),
                         })
            st.markdown(
                f'<div style="font-size:0.65rem;color:var(--muted);margin-top:0.3rem;">'
                f'Confidence TB: {base_summary["avg"]*100:.0f}% · '
                f'Thời gian: {base_summary["time_ms"]:.0f} ms</div>',
                unsafe_allow_html=True,
            )
        else:
            delta = n_final - n_base
            sign = f"+{delta}" if delta > 0 else str(delta)
            st.markdown(
                f'<div style="font-size:0.75rem;color:var(--muted);margin-bottom:0.6rem;">'
                f'Baseline <b>{n_base}</b> → Final <b>{n_final}</b> &nbsp;'
                f'<span style="color:{"#22c55e" if delta>=0 else "#ef4444"};font-weight:700;">{sign}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            rows = [
                {"Lớp": c.capitalize(),
                 "Base": int(base_summary["counts"].get(c, 0)),
                 "Final": int(final_summary["counts"].get(c, 0)),
                 "Δ": int(final_summary["counts"].get(c, 0)) - int(base_summary["counts"].get(c, 0))}
                for c in CLASSES
            ]
            st.dataframe(
                pd.DataFrame(rows), use_container_width=True, hide_index=True,
                column_config={
                    "Lớp": st.column_config.TextColumn("Lớp"),
                    "Base": st.column_config.NumberColumn("Base", format="%d"),
                    "Final": st.column_config.NumberColumn("Final", format="%d"),
                    "Δ": st.column_config.NumberColumn("Δ", format="%+d"),
                },
            )
            st.markdown(
                f'<div style="font-size:0.65rem;color:var(--muted);margin-top:0.3rem;">'
                f'Thời gian: Base {base_summary["time_ms"]:.0f} ms · Final {final_summary["time_ms"]:.0f} ms</div>',
                unsafe_allow_html=True,
            )


# ═══════════════════════════════════════════════════════════════════════════════
#  BOTTOM — post-detection analysis (collapsed by default to keep screen clean)
# ═══════════════════════════════════════════════════════════════════════════════
if base_summary and final_summary:
    avg_b, avg_f = base_summary["avg"], final_summary["avg"]

    with st.expander("Post-detection analysis", expanded=False):
        mc = st.columns(4)
        with mc[0]:
            with st.container(border=True):
                avg = (avg_b + avg_f) / 2
                st.markdown(f"""
                <div class="donut-card">
                    {donut(int(avg*100))}
                    <div>
                        <div class="metric-label" style="font-size:0.8rem;color:var(--muted);">Avg confidence (both)</div>
                        <div style="font-size:1.6rem;font-weight:800;color:var(--ink);">{avg*100:.0f}<span style="font-size:0.9rem;color:var(--muted);">%</span></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
        mc[1].metric("Baseline objects", base_summary["total"])
        mc[2].metric("Final objects", final_summary["total"],
                     delta=final_summary["total"] - base_summary["total"])
        mc[3].metric("Processing time", f"B {base_summary['time_ms']:.0f} ms · F {final_summary['time_ms']:.0f} ms")

        p1, p2, p3 = st.columns(3)
        with p1:
            with st.container(border=True):
                st.markdown('<div class="panel-label">Distribution by class</div>', unsafe_allow_html=True)
                chart_df = pd.DataFrame(
                    {"Baseline": [base_summary["counts"].get(c, 0) for c in CLASSES],
                     "Final":    [final_summary["counts"].get(c, 0) for c in CLASSES]},
                    index=[c.capitalize() for c in CLASSES],
                )
                st.bar_chart(chart_df, height=180)

        with p2:
            with st.container(border=True):
                st.markdown('<div class="panel-label">Confidence distribution</div>', unsafe_allow_html=True)
                bc, edges = np.histogram(base_summary["confidences"], bins=10, range=(0.0, 1.0))
                fc, _ = np.histogram(final_summary["confidences"], bins=10, range=(0.0, 1.0))
                if bc.sum() or fc.sum():
                    hist = pd.DataFrame({"Baseline": bc, "Final": fc},
                                        index=[f"{edges[i]:.1f}" for i in range(len(bc))])
                    st.bar_chart(hist, height=180)
                else:
                    st.caption("No detections.")

        with p3:
            with st.container(border=True):
                st.markdown('<div class="panel-label">Export</div>', unsafe_allow_html=True)
                export_df = pd.DataFrame([
                    {"class": c.capitalize(),
                     "baseline": base_summary["counts"].get(c, 0),
                     "final": final_summary["counts"].get(c, 0),
                     "delta": final_summary["counts"].get(c, 0) - base_summary["counts"].get(c, 0)}
                    for c in CLASSES
                ])
                st.download_button("Export CSV", export_df.to_csv(index=False).encode("utf-8-sig"),
                                   file_name="baseline_vs_final.csv", mime="text/csv", use_container_width=True)
                report = (
                    f"Aperture comparison report\nSource: {active['name'] if active else '-'}\n\n"
                    f"Baseline: {base_summary['total']} objects, {avg_b*100:.1f}% avg conf\n"
                    f"Final:    {final_summary['total']} objects, {avg_f*100:.1f}% avg conf\n\n"
                    "Per class:\n" +
                    "\n".join(f"  {c}: Base={base_summary['counts'].get(c,0)} Final={final_summary['counts'].get(c,0)}" for c in CLASSES)
                )
                st.download_button("Export report", report.encode(),
                                   file_name="report.txt", mime="text/plain", use_container_width=True)
                st.markdown(
                    '<div class="notes" style="margin-top:0.5rem;">Δ dương = Final nhiều box hơn. '
                    'Kết luận định lượng cần mAP50–95 trên test set.</div>',
                    unsafe_allow_html=True,
                )

    if st.session_state.pop("scroll_to_results", False):
        pass  # no-op: content is now above the fold, no scroll needed
