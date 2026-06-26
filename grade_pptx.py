import argparse
import asyncio
import base64
import os
import shutil
import sys
import tempfile
import anthropic
from rubric import Rubric, OneShotOutput
from rubric.autograders import PerCriterionOneShotGrader

import check_layout  # local module (same directory)
import render        # local module (same directory): slides->jpg + markitdown
from dotenv import load_dotenv

load_dotenv()

MODEL = "claude-opus-4-8"
MAX_TOKENS = 8192 
DEFAULT_GRADER_MD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "source", "grader.md")
DEFAULT_RUBRIC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "source", "rubric.yaml")


# --------------------------------------------------------------------------- #
# Prompt assembly: fill grader.md's {{RUBRIC}} and {{DECK_DATA}} placeholders
# --------------------------------------------------------------------------- #
def render_rubric(rubric: Rubric) -> str:
    """The criteria as readable text for the {{RUBRIC}} placeholder."""
    return "\n".join(f"[{c.weight:+g}] {c.requirement.strip()}" for c in rubric.rubric)


def _fmt_pos(p) -> str:
    if not p:
        return "no-geometry"
    return (f"L={p['left']} T={p['top']} W={p['width']} H={p['height']} "
            f"(right={p['right']} bottom={p['bottom']})in")


def build_deck_data(digest: dict, md_slides: list) -> str:
    """The {{DECK_DATA}} block: geometry digest (with pre-computed layout flags)
    followed by the per-slide markitdown text."""
    lines = [
        f"DECK: {digest['file']}   "
        f"slides: {digest['slide_count']}   "
        f"size: {digest['slide_size_in'][0]} x {digest['slide_size_in'][1]} in",
        f"frame (in): {digest['frame_in']}",
        f"tolerances (in): {digest['tolerances_in']}",
        "",
        "All coordinates are in INCHES. frame_violations are shapes crossing the inner "
        "frame; collisions are shape pairs whose bounding boxes intersect. Shapes named "
        "'background' are exempt from these checks. Judge alignment visually from the "
        "slide images rather than from numeric edge deltas.",
        "=" * 70,
        "GEOMETRY DIGEST (per slide):",
    ]
    for s in digest["slides"]:
        f = s["flags"]
        lines.append(f"\n--- SLIDE {s['index']}  (layout: {s['layout']}) ---")
        lines.append(f"  frame violations: {f['frame_violations'] or 'none'}")
        lines.append(f"  collisions: {f['collisions'] or 'none'}")
        lines.append("  shapes:")
        for sh in s["shapes"]:
            head = f"    - {sh['name']} [{sh['kind']}"
            if sh.get("role"):
                head += f"/{sh['role']}"
            head += f", {sh['source']}]"
            if sh.get("rotation"):
                head += f" rot={sh['rotation']}deg"
            lines.append(head)
            lines.append(f"        pos: {_fmt_pos(sh.get('pos_in'))}")
            if sh.get("text"):
                txt = sh["text"].replace("\n", " ⏎ ")
                if len(txt) > 600:
                    txt = txt[:600] + " […digest-truncated for brevity; NOT clipped in the deck]"
                lines.append(f'        text: "{txt}"')
                lines.append(f"        font: {sh.get('font_family')} {sh.get('font_size_pt')}pt "
                             f"color={sh.get('font_color')} align={sh.get('alignment')}")

    lines.append("")
    lines.append("=" * 70)
    lines.append("DECK TEXT PER SLIDE (markitdown extraction):")
    for i, body in enumerate(md_slides, start=1):
        lines.append(f"\n--- slide {i} text ---\n{body}" if body else f"\n--- slide {i} text --- (empty)")
    return "\n".join(lines)


def build_prompt(grader_md: str, rubric: Rubric, deck_data: str) -> str:
    """Fill grader.md's two placeholders to form the text the model grades."""
    return (grader_md
            .replace("{{RUBRIC}}", render_rubric(rubric))
            .replace("{{DECK_DATA}}", deck_data))


# --------------------------------------------------------------------------- #
# Custom Anthropic generate_fn: the bridge that lets rubric grade images+geometry.
# rubric supplies the criteria scaffolding; we attach the slide JPGs and the
# filled prompt, call the model once, and hand back structured MET/UNMET verdicts.
# --------------------------------------------------------------------------- #
def _image_blocks(image_paths: list) -> list:
    """Slide JPGs as vision content blocks, each labelled with its slide number."""
    blocks = []
    for i, path in enumerate(image_paths, start=1):
        with open(path, "rb") as fh:
            data = base64.standard_b64encode(fh.read()).decode("ascii")
        blocks.append({"type": "text", "text": f"Slide {i}:"})
        blocks.append({"type": "image", "source": {
            "type": "base64", "media_type": "image/jpeg", "data": data}})
    return blocks


def make_generate_fn(client, prompt_text: str, image_paths: list):
    """Async generate_fn matching rubric's OneShotGenerateFn: one call judges all
    criteria. The filled grader prompt is a cached system block; the slide JPGs ride
    on the user turn (images can't be cached)."""
    images = _image_blocks(image_paths)

    async def generate_fn(system_prompt: str, user_prompt: str, **kwargs) -> OneShotOutput:
        system = [
            {"type": "text", "text": system_prompt},
            {"type": "text", "text": prompt_text, "cache_control": {"type": "ephemeral"}},
        ]
        resp = await client.messages.parse(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system,
            messages=[{"role": "user", "content": images + [{"type": "text", "text": user_prompt}]}],
            output_format=OneShotOutput,
        )
        return resp.parsed_output

    return generate_fn


# --------------------------------------------------------------------------- #
# Output: write rubric's computed result as a markdown report
# --------------------------------------------------------------------------- #
def write_prompt_log(prompt: str, digest, logs_dir="logs") -> str:
    """Save the exact text payload sent to the model to logs/<deck>-prompt-<ts>.md
    and return the path. This is the deck data + rubric + grader instructions the
    model grades on — persisted so a run's input is auditable, not discarded."""
    from datetime import datetime
    os.makedirs(logs_dir, exist_ok=True)
    stem = os.path.splitext(digest["file"])[0]
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    path = os.path.join(logs_dir, f"{stem}-prompt-{ts}.md")
    with open(path, "w") as fh:
        fh.write(prompt + "\n")
    return path


def write_report_md(report, digest, dpi, leftovers, logs_dir="logs") -> str:
    """Render rubric's EvaluationReport to logs/<deck>-<ts>.md and return the path."""
    from datetime import datetime
    os.makedirs(logs_dir, exist_ok=True)
    stem = os.path.splitext(digest["file"])[0]
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    path = os.path.join(logs_dir, f"{stem}-{ts}.md")

    rows = report.report or []
    pos = [r for r in rows if r.weight >= 0]
    pen = [r for r in rows if r.weight < 0]
    met = lambda r: str(r.verdict).upper() == "MET"
    pos_met = sum(1 for r in pos if met(r))
    pen_hit = sum(1 for r in pen if met(r))

    out = [
        f"# Rubric grade — {digest['file']}",
        "",
        f"- **Slides:** {digest['slide_count']}",
        f"- **Model:** {MODEL}",
        f"- **Graded at:** {datetime.now().isoformat(timespec='seconds')}",
        f"- **Render dpi:** {dpi}",
        f"- **Placeholder leftovers:** {len(leftovers or [])}",
        "",
        f"## Weighted score: {report.score * 100:.1f}%  "
        f"(raw weighted sum {report.raw_score:.1f})",
        "",
        f"Positive criteria met: {pos_met}/{len(pos)} · "
        f"Penalties triggered: {pen_hit}/{len(pen)}",
        "",
        "| Weight | Verdict | Criterion | Reason |",
        "| ---: | :--- | :--- | :--- |",
    ]
    for r in rows:
        if r.weight < 0:
            verdict = "⚠ PENALTY" if met(r) else "ok"
        else:
            verdict = "✓ MET" if met(r) else "✗ UNMET"
        req = " ".join(r.requirement.split())
        reason = " ".join((r.reason or "").split()).replace("|", "\\|")
        out.append(f"| {r.weight:+g} | {verdict} | {req} | {reason} |")

    with open(path, "w") as fh:
        fh.write("\n".join(out) + "\n")
    return path


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
async def run(args) -> int:
    if not os.getenv("ANTHROPIC_API_KEY"):
        sys.stderr.write(
            "ANTHROPIC_API_KEY is not set. Set it and re-run, e.g.:\n"
            "  export ANTHROPIC_API_KEY=sk-ant-...\n"
            "(or put it in a .env file in this directory)\n")
        return 2
    for path, what in [(args.pptx, "File"), (args.rubric, "Rubric file"),
                       (args.grader, "Grader markdown")]:
        if not os.path.exists(path):
            sys.stderr.write(f"{what} not found: {path}\n")
            return 2

    try:
        rubric = Rubric.from_file(args.rubric)
    except Exception as e:
        sys.stderr.write(f"Failed to load rubric: {e}\n")
        return 2

    with open(args.grader) as fh:
        grader_md = fh.read()

    try:
        digest = check_layout.extract_deck_digest(args.pptx)
    except check_layout.PackageNotFoundError:
        sys.stderr.write(f"Cannot open presentation: {args.pptx}\n")
        return 2

    # Rendering is required — a failure is fatal (no geometry-only fallback).
    tmp_dir = (os.path.join(args.logs_dir, f"{os.path.splitext(digest['file'])[0]}-images")
               if args.keep_images else tempfile.mkdtemp(prefix="pptx-qa-"))
    try:
        images = render.render_slides(args.pptx, tmp_dir, dpi=args.dpi)
        md_slides = render.extract_markdown_per_slide(args.pptx)
        leftovers = render.find_placeholder_leftovers("\n".join(md_slides))
    except render.RenderError as e:
        sys.stderr.write(f"\n*** Rendering failed (required for grading): ***\n{e}\n")
        if not args.keep_images:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return 2

    prompt = build_prompt(grader_md, rubric, build_deck_data(digest, md_slides))
    prompt_path = write_prompt_log(prompt, digest, args.logs_dir)
    print(f"Model payload saved to {prompt_path}", file=sys.stderr)
    if args.show_digest:
        print(prompt)
        if not args.keep_images:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return 0

    print(f"Grading {digest['file']} ({digest['slide_count']} slides) against "
          f"{len(rubric.rubric)} criteria using {MODEL} [dpi {args.dpi}]…", file=sys.stderr)

    client = anthropic.AsyncAnthropic()
    grader = PerCriterionOneShotGrader(generate_fn=make_generate_fn(client, prompt, images))
    report = await rubric.grade(prompt, autograder=grader)
    path = write_report_md(report, digest, args.dpi, leftovers, args.logs_dir)
    print(f"Weighted score: {report.score * 100:.1f}%  —  report saved to {path}", file=sys.stderr)

    if not args.keep_images:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return 0


def main(argv):
    p = argparse.ArgumentParser(
        description="Grade a PowerPoint deck against a weighted rubric (LLM + geometry + images).")
    p.add_argument("pptx", help="path to the .pptx file")
    p.add_argument("--rubric", default=DEFAULT_RUBRIC,
                   help="path to the rubric YAML (default: powerpoint-rubric.yaml beside this script)")
    p.add_argument("--grader", default=DEFAULT_GRADER_MD,
                   help="path to the grader prompt markdown (default: grader.md beside this script)")
    p.add_argument("--show-digest", action="store_true",
                   help="print the filled grader prompt and exit (renders, but no grading API call)")
    p.add_argument("--logs-dir", default="logs",
                   help="directory to save the markdown report in (default: logs)")
    p.add_argument("--dpi", type=int, default=150,
                   help="resolution for rendered slide images (default: 150)")
    p.add_argument("--keep-images", action="store_true",
                   help="keep the rendered slide JPGs under --logs-dir instead of a temp dir")
    args = p.parse_args(argv)
    return asyncio.run(run(args))


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
