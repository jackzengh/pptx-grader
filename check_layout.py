#!/usr/bin/env python3
"""
check_layout.py - compact geometry checker for PowerPoint decks.

This module focuses on three deterministic checks:
  1) alignment precision (near-miss edge alignment)
  2) collisions (bounding-box intersections)
  3) frame adherence (content outside the inner margin frame)

It intentionally drops the old YAML rule engine and ASCII map renderer so the
core geometry logic stays small and maintainable.
"""

import os
import sys
from dataclasses import dataclass, field
from typing import Optional

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.exc import PackageNotFoundError
except ImportError:
    sys.stderr.write("Missing dependency: python-pptx. Run: pip install python-pptx\n")
    sys.exit(2)


EMU_PER_INCH = 914400
DEFAULT_MARGIN_IN = 0.5
DEFAULT_ALIGNMENT_TOL_PX = 3
DEFAULT_FRAME_TOL_PX = 2
DEFAULT_COLLISION_TOL_PX = 1
DEFAULT_PX_DPI = 96


def emu_to_in(v: Optional[int]) -> Optional[float]:
    return None if v is None else round(v / EMU_PER_INCH, 3)


def in_to_emu(v: float) -> int:
    return round(v * EMU_PER_INCH)


def px_to_emu(px: float, dpi: int = DEFAULT_PX_DPI) -> int:
    return round((px / dpi) * EMU_PER_INCH)


@dataclass
class Box:
    left: Optional[int]
    top: Optional[int]
    width: Optional[int]
    height: Optional[int]

    @property
    def right(self) -> Optional[int]:
        if self.left is None or self.width is None:
            return None
        return self.left + self.width

    @property
    def bottom(self) -> Optional[int]:
        if self.top is None or self.height is None:
            return None
        return self.top + self.height

    def is_resolved(self) -> bool:
        return None not in (self.left, self.top, self.width, self.height)


@dataclass
class ResolvedShape:
    name: str
    kind: str
    box: Box
    source: str
    ph_type: Optional[str] = None
    ph_idx: Optional[int] = None
    has_text: bool = False
    rotation: float = 0.0
    geometry_inherited: bool = False


@dataclass
class SlideModel:
    index: int
    layout_name: str
    slide_w: int
    slide_h: int
    shapes: list = field(default_factory=list)


@dataclass
class CheckResult:
    check_id: str
    slide_index: Optional[int]
    status: str
    message: str
    shape_names: list = field(default_factory=list)


_PH_TYPE_MAP = {
    "TITLE": "title",
    "CENTER_TITLE": "title",
    "BODY": "body",
    "SUBTITLE": "subTitle",
    "DATE": "dt",
    "FOOTER": "ftr",
    "SLIDE_NUMBER": "sldNum",
    "OBJECT": "obj",
}


def ph_type_str(pf) -> Optional[str]:
    try:
        return _PH_TYPE_MAP.get(pf.type.name if pf.type is not None else None)
    except (AttributeError, KeyError):
        return None


def shape_kind(shape) -> str:
    try:
        st = shape.shape_type
    except (AttributeError, ValueError):
        st = None
    if st == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if st == MSO_SHAPE_TYPE.GROUP:
        return "group"
    if getattr(shape, "is_placeholder", False):
        return "placeholder"
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        return "text"
    if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
        return "auto"
    return "other"


def _local_box(shape) -> Box:
    return Box(shape.left, shape.top, shape.width, shape.height)


def _match_placeholder(ph_type, ph_idx, candidates, by="idx_then_type"):
    if by in ("idx_then_type",) and ph_idx is not None:
        for c in candidates:
            if c.placeholder_format.idx == ph_idx:
                return c
    for c in candidates:
        if ph_type_str(c.placeholder_format) == ph_type:
            return c
    return None


def resolve_geometry(shape, layout, master) -> tuple:
    box = _local_box(shape)
    inherited = False
    if box.is_resolved() or not getattr(shape, "is_placeholder", False):
        return box, inherited

    pf = shape.placeholder_format
    ph_type = ph_type_str(pf)
    ph_idx = pf.idx

    def fill_from(src_shape):
        nonlocal inherited
        for attr in ("left", "top", "width", "height"):
            if getattr(box, attr) is None:
                val = getattr(src_shape, attr)
                if val is not None:
                    setattr(box, attr, val)
                    inherited = True

    lay_ph = _match_placeholder(ph_type, ph_idx, list(layout.placeholders), by="idx_then_type")
    if lay_ph is not None and not box.is_resolved():
        fill_from(lay_ph)
    if not box.is_resolved() and master is not None:
        mas_ph = _match_placeholder(ph_type, ph_idx, list(master.placeholders), by="type")
        if mas_ph is not None:
            fill_from(mas_ph)
    return box, inherited


def to_resolved_shape(shape, source, layout, master) -> ResolvedShape:
    kind = shape_kind(shape)
    if source == "slide":
        box, inherited = resolve_geometry(shape, layout, master)
    else:
        box, inherited = _local_box(shape), False
    pf = shape.placeholder_format if getattr(shape, "is_placeholder", False) else None
    has_text = getattr(shape, "has_text_frame", False) and bool(shape.text_frame.text.strip())
    try:
        rot = float(shape.rotation)
    except (AttributeError, TypeError, ValueError):
        rot = 0.0
    return ResolvedShape(
        name=shape.name,
        kind=kind,
        box=box,
        source=source,
        ph_type=ph_type_str(pf) if pf is not None else None,
        ph_idx=pf.idx if pf is not None else None,
        has_text=has_text,
        rotation=rot,
        geometry_inherited=inherited,
    )


def collect_inherited_pictures(slide, layout, master) -> list:
    out = []
    if slide._element.get("showMasterSp") == "0":
        return out
    for src_name, container in (("layout", layout), ("master", master)):
        if container is None:
            continue
        for sh in container.shapes:
            try:
                if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    out.append(to_resolved_shape(sh, src_name, layout, master))
            except (AttributeError, KeyError, ValueError):
                continue
    return out


def build_slide_models(prs, slide_filter=None) -> list:
    models = []
    w, h = prs.slide_width, prs.slide_height
    for i, slide in enumerate(prs.slides, start=1):
        if slide_filter is not None and i not in slide_filter:
            continue
        layout = slide.slide_layout
        master = layout.slide_master
        shapes = [to_resolved_shape(sh, "slide", layout, master) for sh in slide.shapes]
        seen = {(s.name, s.box.left, s.box.top) for s in shapes}
        for ps in collect_inherited_pictures(slide, layout, master):
            key = (ps.name, ps.box.left, ps.box.top)
            if key not in seen:
                shapes.append(ps)
                seen.add(key)
        models.append(SlideModel(i, layout.name, w, h, shapes))
    return models


def _run_color_hex(run) -> Optional[str]:
    try:
        c = run.font.color
        if c and c.type is not None and c.rgb is not None:
            return f"#{str(c.rgb)}"
    except (AttributeError, TypeError):
        pass
    return None


def _shape_text_style(shape) -> dict:
    if not getattr(shape, "has_text_frame", False):
        return {}
    tf = shape.text_frame
    text = tf.text.strip()
    if not text:
        return {}
    fonts, sizes, colors, aligns = {}, {}, {}, {}

    def bump(d, k):
        if k is not None:
            d[k] = d.get(k, 0) + 1

    for para in tf.paragraphs:
        bump(aligns, para.alignment.name if para.alignment is not None else None)
        for run in para.runs:
            bump(fonts, run.font.name)
            bump(sizes, run.font.size.pt if run.font.size is not None else None)
            bump(colors, _run_color_hex(run))

    def top(d):
        return max(d, key=d.get) if d else None

    return {
        "text": text,
        "font_family": top(fonts),
        "font_size_pt": top(sizes),
        "font_color": top(colors),
        "alignment": top(aligns),
    }


def _raw_shapes_by_name(slide):
    out = {}
    for sh in slide.shapes:
        out.setdefault(sh.name, sh)
    return out


def _iter_resolved_content_shapes(model: SlideModel):
    for sh in model.shapes:
        if sh.kind == "group":
            continue
        if sh.box.is_resolved():
            yield sh


def check_collisions(model: SlideModel, tol_emu: int) -> list:
    cid = "collisions"
    shapes = list(_iter_resolved_content_shapes(model))
    results = []
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            ox = min(a.box.right, b.box.right) - max(a.box.left, b.box.left)
            oy = min(a.box.bottom, b.box.bottom) - max(a.box.top, b.box.top)
            if ox > tol_emu and oy > tol_emu:
                results.append(CheckResult(
                    cid,
                    model.index,
                    "fail",
                    f"{a.name} overlaps {b.name} by {emu_to_in(ox)}in x {emu_to_in(oy)}in",
                    [a.name, b.name],
                ))
    if not results:
        results.append(CheckResult(cid, model.index, "pass", "no collisions detected"))
    return results


def check_frame_adherence(model: SlideModel, margin_emu: int, tol_emu: int) -> list:
    cid = "frame_adherence"
    L = margin_emu
    T = margin_emu
    R = model.slide_w - margin_emu
    B = model.slide_h - margin_emu
    results = []
    checked = 0
    for sh in _iter_resolved_content_shapes(model):
        checked += 1
        b = sh.box
        violations = []
        if b.left < L - tol_emu:
            violations.append(f"left by {emu_to_in(L - b.left)}in")
        if b.top < T - tol_emu:
            violations.append(f"top by {emu_to_in(T - b.top)}in")
        if b.right > R + tol_emu:
            violations.append(f"right by {emu_to_in(b.right - R)}in")
        if b.bottom > B + tol_emu:
            violations.append(f"bottom by {emu_to_in(b.bottom - B)}in")
        if violations:
            results.append(CheckResult(
                cid,
                model.index,
                "fail",
                f"{sh.name} crosses frame ({', '.join(violations)})",
                [sh.name],
            ))
    if not results:
        results.append(CheckResult(cid, model.index, "pass", f"{checked} shape(s) inside frame"))
    return results


def check_alignment_precision(model: SlideModel, tol_emu: int, near_window_emu: Optional[int] = None) -> list:
    cid = "alignment_precision"
    near = near_window_emu if near_window_emu is not None else max(tol_emu * 6, in_to_emu(0.12))
    shapes = list(_iter_resolved_content_shapes(model))
    if len(shapes) < 2:
        return [CheckResult(cid, model.index, "skip", "not enough resolved shapes")]

    edge_getters = {
        "left": lambda s: s.box.left,
        "right": lambda s: s.box.right,
        "top": lambda s: s.box.top,
        "bottom": lambda s: s.box.bottom,
    }
    failures = []
    for edge, get_val in edge_getters.items():
        vals = sorted([(s.name, get_val(s)) for s in shapes], key=lambda x: x[1])
        for (n1, v1), (n2, v2) in zip(vals, vals[1:]):
            delta = v2 - v1
            if tol_emu < delta <= near:
                failures.append(CheckResult(
                    cid,
                    model.index,
                    "fail",
                    f"near-miss {edge} alignment: {n1} vs {n2} (delta {emu_to_in(delta)}in, tol {emu_to_in(tol_emu)}in)",
                    [n1, n2],
                ))
    if failures:
        return failures
    return [CheckResult(cid, model.index, "pass", "no near-miss edge alignment issues")]


def run_checks(
    models: list,
    *,
    margin_in: float = DEFAULT_MARGIN_IN,
    alignment_tol_px: int = DEFAULT_ALIGNMENT_TOL_PX,
    frame_tol_px: int = DEFAULT_FRAME_TOL_PX,
    collision_tol_px: int = DEFAULT_COLLISION_TOL_PX,
    px_dpi: int = DEFAULT_PX_DPI,
) -> list:
    margin_emu = in_to_emu(margin_in)
    alignment_tol_emu = px_to_emu(alignment_tol_px, px_dpi)
    frame_tol_emu = px_to_emu(frame_tol_px, px_dpi)
    collision_tol_emu = px_to_emu(collision_tol_px, px_dpi)
    out = []
    for m in models:
        out.extend(check_alignment_precision(m, alignment_tol_emu))
        out.extend(check_collisions(m, collision_tol_emu))
        out.extend(check_frame_adherence(m, margin_emu, frame_tol_emu))
    return out


def audit_presentation(
    pptx_path: str,
    *,
    margin_in: float = DEFAULT_MARGIN_IN,
    alignment_tol_px: int = DEFAULT_ALIGNMENT_TOL_PX,
    frame_tol_px: int = DEFAULT_FRAME_TOL_PX,
    collision_tol_px: int = DEFAULT_COLLISION_TOL_PX,
    px_dpi: int = DEFAULT_PX_DPI,
) -> tuple:
    prs = Presentation(pptx_path)
    models = build_slide_models(prs, None)
    slides = list(prs.slides)
    results = run_checks(
        models,
        margin_in=margin_in,
        alignment_tol_px=alignment_tol_px,
        frame_tol_px=frame_tol_px,
        collision_tol_px=collision_tol_px,
        px_dpi=px_dpi,
    )

    fails_by_slide = {}
    for r in results:
        if r.slide_index is None or r.status != "fail":
            continue
        slot = fails_by_slide.setdefault(r.slide_index, {
            "alignment_issues": [],
            "collisions": [],
            "frame_violations": [],
        })
        if r.check_id == "alignment_precision":
            slot["alignment_issues"].append(r.message)
        elif r.check_id == "collisions":
            slot["collisions"].append(r.message)
        elif r.check_id == "frame_adherence":
            slot["frame_violations"].append(r.message)

    deck = {
        "file": os.path.basename(pptx_path),
        "slide_count": len(models),
        "slide_size_in": [emu_to_in(prs.slide_width), emu_to_in(prs.slide_height)],
        "frame_in": {
            "margin": margin_in,
            "left": margin_in,
            "top": margin_in,
            "right_line": emu_to_in(prs.slide_width - in_to_emu(margin_in)),
            "bottom_line": emu_to_in(prs.slide_height - in_to_emu(margin_in)),
        },
        "tolerances_px": {
            "alignment": alignment_tol_px,
            "collision": collision_tol_px,
            "frame": frame_tol_px,
            "dpi_basis": px_dpi,
        },
        "slides": [],
    }

    for model, slide in zip(models, slides):
        raw = _raw_shapes_by_name(slide)
        shapes_out = []
        for rs in model.shapes:
            entry = {
                "name": rs.name,
                "kind": rs.kind,
                "source": rs.source,
                "role": rs.ph_type,
                "rotation": rs.rotation or None,
            }
            if rs.box.is_resolved():
                entry["pos_in"] = {
                    "left": emu_to_in(rs.box.left),
                    "top": emu_to_in(rs.box.top),
                    "width": emu_to_in(rs.box.width),
                    "height": emu_to_in(rs.box.height),
                    "right": emu_to_in(rs.box.right),
                    "bottom": emu_to_in(rs.box.bottom),
                }
            else:
                entry["pos_in"] = None
            sh = raw.get(rs.name)
            if sh is not None:
                entry.update(_shape_text_style(sh))
            shapes_out.append(entry)

        flags = fails_by_slide.get(model.index, {
            "alignment_issues": [],
            "collisions": [],
            "frame_violations": [],
        })
        deck["slides"].append({
            "index": model.index,
            "layout": model.layout_name,
            "shapes": shapes_out,
            "flags": flags,
        })
    return deck, results


def extract_deck_digest(pptx_path: str, **kwargs) -> dict:
    """The geometry/text digest (with per-slide layout flags). The grader's only
    entry point into this module."""
    digest, _ = audit_presentation(pptx_path, **kwargs)
    return digest
