#!/usr/bin/env python3
"""
render.py - Turn a .pptx into things a vision model can judge:
  * one JPG per slide (LibreOffice -> PDF -> pdftoppm), and
  * the deck's text per slide (markitdown), plus a leftover-placeholder scan.

This is the whole visual-QA surface. Delete this file (and its call sites in
grade_pptx.py) and the grader falls back to geometry-only text grading.

System deps (NOT pip-installable):
  * LibreOffice  -> the `soffice` binary  (brew install --cask libreoffice)
  * Poppler      -> the `pdftoppm` binary (brew install poppler)
"""

import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

# Mirrors the grep in the QA spec: obvious leftover/template text.
_LEFTOVER_RE = re.compile(r"xxxx|lorem|ipsum|this.*(page|slide).*layout", re.IGNORECASE)

# markitdown delimits every slide with this HTML comment.
_SLIDE_MARKER_RE = re.compile(r"<!--\s*Slide number:\s*\d+\s*-->", re.IGNORECASE)


class RenderError(Exception):
    """A required tool is missing or a conversion step failed. Caught by the
    grader so it can warn loudly and fall back to text-only grading."""


def find_soffice() -> str:
    """Locate the LibreOffice binary, or raise RenderError with an install hint."""
    candidates = [
        os.environ.get("SOFFICE"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    raise RenderError(
        "LibreOffice (soffice) not found — needed to render slides to images.\n"
        "Install it:  brew install --cask libreoffice\n"
        "(or set $SOFFICE to the binary path)."
    )


def render_slides(pptx_path: str, out_dir: str, dpi: int = 150) -> list[str]:
    """Render every slide to a JPG. Returns sorted ['<out_dir>/slide-01.jpg', ...]."""
    if not shutil.which("pdftoppm"):
        raise RenderError(
            "pdftoppm not found — needed to turn the PDF into images.\n"
            "Install it:  brew install poppler"
        )
    soffice = find_soffice()
    os.makedirs(out_dir, exist_ok=True)

    # A throwaway profile dir lets soffice run headless without colliding with a
    # logged-in LibreOffice instance.
    with tempfile.TemporaryDirectory() as profile:
        try:
            subprocess.run(
                [soffice, "--headless",
                 f"-env:UserInstallation=file://{profile}",
                 "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                check=True, capture_output=True, text=True,
            )
        except subprocess.CalledProcessError as exc:
            raise RenderError(f"soffice pdf conversion failed:\n{exc.stderr}") from exc

    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RenderError(f"expected PDF not produced: {pdf}")

    try:
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", str(dpi), pdf, os.path.join(out_dir, "slide")],
            check=True, capture_output=True, text=True,
        )
    except subprocess.CalledProcessError as exc:
        raise RenderError(f"pdftoppm failed:\n{exc.stderr}") from exc

    jpgs = sorted(glob.glob(os.path.join(out_dir, "slide-*.jpg")))
    if not jpgs:
        raise RenderError(f"no slide images produced in {out_dir}")
    return jpgs


def extract_markdown_per_slide(pptx_path: str) -> list[str]:
    """Return markitdown text split per slide (index 0 = slide 1).

    markitdown separates slides with a `<!-- Slide number: N -->` comment; we
    split on it. If that marker is absent, return the whole deck as one element.
    """
    text = _markitdown_text(pptx_path)
    parts = _SLIDE_MARKER_RE.split(text)
    # split() yields a leading chunk before the first marker (usually empty);
    # drop it only when it's blank so slide numbering stays aligned.
    if parts and not parts[0].strip():
        parts = parts[1:]
    cleaned = [p.strip() for p in parts]
    return cleaned or [text.strip()]


def find_placeholder_leftovers(md_text: str) -> list[str]:
    """Lines containing obvious leftover/template text (xxxx, lorem ipsum, ...)."""
    return [line.strip() for line in md_text.splitlines() if _LEFTOVER_RE.search(line)]


def _markitdown_text(pptx_path: str) -> str:
    """markitdown via its Python API, falling back to the CLI."""
    try:
        from markitdown import MarkItDown
        return MarkItDown().convert(pptx_path).text_content
    except Exception:
        try:
            out = subprocess.run(
                [sys.executable, "-m", "markitdown", pptx_path],
                check=True, capture_output=True, text=True,
            )
            return out.stdout
        except Exception as exc:  # noqa: BLE001 - surface any markitdown failure
            raise RenderError(f"markitdown failed on {pptx_path}: {exc}") from exc
