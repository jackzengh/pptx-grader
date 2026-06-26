#!/usr/bin/env python3
"""
check_layout.py - compact geometry checker for PowerPoint decks.

This module focuses on two deterministic checks:
  1) collisions (bounding-box intersections)
  2) frame adherence (content outside the inner margin frame)

Shapes named "background" (case-insensitive) are exempt from both checks.

All geometry is stored and reported in INCHES. PowerPoint stores positions in
EMU (English Metric Units); we convert once at read time and work in inches
everywhere else.
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
DEFAULT_FRAME_TOL_IN = 0.02
DEFAULT_COLLISION_TOL_IN = 0.01


def _emu_to_in(v: Optional[int]) -> Optional[float]:
    """Convert a raw EMU measurement to inches, preserving None."""
    return None if v is None else round(v / EMU_PER_INCH, 3)


@dataclass
class Box:
    """A rectangle in inches. Any field may be None for unresolved placeholders."""
    left: Optional[float]
    top: Optional[float]
    width: Optional[float]
    height: Optional[float]

    @property
    def right(self) -> Optional[float]:
        if self.left is None or self.width is None:
            return None
        return self.left + self.width

    @property
    def bottom(self) -> Optional[float]:
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
    text_style: dict = field(default_factory=dict)


@dataclass
class SlideModel:
    index: int
    layout_name: str
    slide_w: float
    slide_h: float
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
    """Build a Box (in inches) from the shape's own stored geometry."""
    return Box(
        _emu_to_in(shape.left),
        _emu_to_in(shape.top),
        _emu_to_in(shape.width),
        _emu_to_in(shape.height),
    )


def _match_placeholder(ph_type, ph_idx, candidates, by="idx_then_type"):
    if by in ("idx_then_type",) and ph_idx is not None:
        for c in candidates:
            if c.placeholder_format.idx == ph_idx:
                return c
    for c in candidates:
        if ph_type_str(c.placeholder_format) == ph_type:
            return c
    return None


def resolve_geometry(shape, layout, master) -> Box:
    """Resolve a shape's box in inches, walking slide -> layout -> master for
    placeholders that leave geometry unset on the slide."""
    box = _local_box(shape)
    if box.is_resolved() or not getattr(shape, "is_placeholder", False):
        return box

    pf = shape.placeholder_format
    ph_type = ph_type_str(pf)
    ph_idx = pf.idx

    def fill_from(src_shape):
        for attr in ("left", "top", "width", "height"):
            if getattr(box, attr) is None:
                val = getattr(src_shape, attr)
                if val is not None:
                    setattr(box, attr, _emu_to_in(val))

    lay_ph = _match_placeholder(ph_type, ph_idx, list(layout.placeholders), by="idx_then_type")
    if lay_ph is not None and not box.is_resolved():
        fill_from(lay_ph)
    if not box.is_resolved() and master is not None:
        mas_ph = _match_placeholder(ph_type, ph_idx, list(master.placeholders), by="type")
        if mas_ph is not None:
            fill_from(mas_ph)
    return box


def to_resolved_shape(shape, source, layout, master, box=None) -> ResolvedShape:
    """Turn a raw python-pptx shape into a ResolvedShape (inches). Pass an
    explicit `box` to override geometry (used when flattening groups)."""
    kind = shape_kind(shape)
    if box is None:
        box = resolve_geometry(shape, layout, master) if source == "slide" else _local_box(shape)
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
        text_style=_shape_text_style(shape),
    )


def flatten_group(group, source, layout, master) -> list:
    """Recurse into a group, returning its leaf shapes as ResolvedShapes whose
    boxes are mapped onto the slide via the group's coordinate transform.

    A group's <a:xfrm> defines its position/extent on the slide (off/ext) and
    the child coordinate space (chOff/chExt). Each child is mapped:
        scale  = group_extent / child_extent
        on_slide = group_off + (child_coord - child_off) * scale
    """
    out = []
    try:
        g_left, g_top = group.left, group.top
        g_w, g_h = group.width, group.height
        xfrm = group._element.grpSpPr.xfrm
        ch_off, ch_ext = xfrm.chOff, xfrm.chExt
        ch_x, ch_y = ch_off.x, ch_off.y
        ch_w, ch_h = ch_ext.cx, ch_ext.cy
    except (AttributeError, TypeError):
        return out
    if not ch_w or not ch_h or g_w is None or g_h is None:
        return out

    scale_x = g_w / ch_w
    scale_y = g_h / ch_h

    def map_box(shape):
        if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
            return None
        left = g_left + (shape.left - ch_x) * scale_x
        top = g_top + (shape.top - ch_y) * scale_y
        return Box(
            _emu_to_in(left),
            _emu_to_in(top),
            _emu_to_in(shape.width * scale_x),
            _emu_to_in(shape.height * scale_y),
        )

    for child in group.shapes:
        if shape_kind(child) == "group":
            out.extend(flatten_group(child, source, layout, master))
            continue
        out.append(to_resolved_shape(child, source, layout, master, box=map_box(child)))
    return out


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
    w, h = _emu_to_in(prs.slide_width), _emu_to_in(prs.slide_height)
    for i, slide in enumerate(prs.slides, start=1):
        if slide_filter is not None and i not in slide_filter:
            continue
        layout = slide.slide_layout
        master = layout.slide_master
        shapes = []
        for sh in slide.shapes:
            if shape_kind(sh) == "group":
                shapes.extend(flatten_group(sh, "slide", layout, master))
            else:
                shapes.append(to_resolved_shape(sh, "slide", layout, master))
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


def _is_background(shape: ResolvedShape) -> bool:
    """Shapes the author named as a background are exempt from geometry checks
    (e.g. a full-card or full-slide backing rectangle that content sits on top of)."""
    return "background" in (shape.name or "").lower()


def _iter_resolved_content_shapes(model: SlideModel):
    for sh in model.shapes:
        if _is_background(sh):
            continue
        if sh.box.is_resolved():
            yield sh


def check_collisions(model: SlideModel, tol_in: float) -> list:
    cid = "collisions"
    shapes = list(_iter_resolved_content_shapes(model))
    results = []
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            ox = min(a.box.right, b.box.right) - max(a.box.left, b.box.left)
            oy = min(a.box.bottom, b.box.bottom) - max(a.box.top, b.box.top)
            if ox > tol_in and oy > tol_in:
                results.append(CheckResult(
                    cid,
                    model.index,
                    "fail",
                    f"{a.name} overlaps {b.name} by {round(ox, 3)}in x {round(oy, 3)}in",
                    [a.name, b.name],
                ))
    if not results:
        results.append(CheckResult(cid, model.index, "pass", "no collisions detected"))
    return results


def check_frame_adherence(model: SlideModel, margin_in: float, tol_in: float) -> list:
    cid = "frame_adherence"
    L = margin_in
    T = margin_in
    R = model.slide_w - margin_in
    B = model.slide_h - margin_in
    results = []
    checked = 0
    for sh in _iter_resolved_content_shapes(model):
        checked += 1
        b = sh.box
        violations = []
        if b.left < L - tol_in:
            violations.append(f"left by {round(L - b.left, 3)}in")
        if b.top < T - tol_in:
            violations.append(f"top by {round(T - b.top, 3)}in")
        if b.right > R + tol_in:
            violations.append(f"right by {round(b.right - R, 3)}in")
        if b.bottom > B + tol_in:
            violations.append(f"bottom by {round(b.bottom - B, 3)}in")
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


def run_checks(
    models: list,
    *,
    margin_in: float = DEFAULT_MARGIN_IN,
    frame_tol_in: float = DEFAULT_FRAME_TOL_IN,
    collision_tol_in: float = DEFAULT_COLLISION_TOL_IN,
) -> list:
    out = []
    for m in models:
        out.extend(check_collisions(m, collision_tol_in))
        out.extend(check_frame_adherence(m, margin_in, frame_tol_in))
    return out


def audit_presentation(
    pptx_path: str,
    *,
    margin_in: float = DEFAULT_MARGIN_IN,
    frame_tol_in: float = DEFAULT_FRAME_TOL_IN,
    collision_tol_in: float = DEFAULT_COLLISION_TOL_IN,
) -> tuple:
    prs = Presentation(pptx_path)
    models = build_slide_models(prs, None)
    results = run_checks(
        models,
        margin_in=margin_in,
        frame_tol_in=frame_tol_in,
        collision_tol_in=collision_tol_in,
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
        if r.check_id == "collisions":
            slot["collisions"].append(r.message)
        elif r.check_id == "frame_adherence":
            slot["frame_violations"].append(r.message)

    slide_w_in = _emu_to_in(prs.slide_width)
    slide_h_in = _emu_to_in(prs.slide_height)
    deck = {
        "file": os.path.basename(pptx_path),
        "slide_count": len(models),
        "slide_size_in": [slide_w_in, slide_h_in],
        "frame_in": {
            "margin": margin_in,
            "left": margin_in,
            "top": margin_in,
            "right_line": round(slide_w_in - margin_in, 3),
            "bottom_line": round(slide_h_in - margin_in, 3),
        },
        "tolerances_in": {
            "collision": collision_tol_in,
            "frame": frame_tol_in,
        },
        "slides": [],
    }

    for model in models:
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
                    "left": rs.box.left,
                    "top": rs.box.top,
                    "width": rs.box.width,
                    "height": rs.box.height,
                    "right": round(rs.box.right, 3),
                    "bottom": round(rs.box.bottom, 3),
                }
            else:
                entry["pos_in"] = None
            if rs.text_style:
                entry.update(rs.text_style)
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
