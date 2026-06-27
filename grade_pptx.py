from datetime import datetime
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


_RULE = "=" * 70

_PREAMBLE = (
    "Judge every visual criterion — overlaps, off-slide bleed, clipping, overflow, "
    "alignment, margins — directly from the rendered slide IMAGES, which are ground "
    "truth for what is actually visible."
)
_CHECKS_PREAMBLE = (
    "PRECOMPUTED DETERMINISTIC CHECKS — these are computed directly from the .pptx "
    "data (font names, point sizes, hex colors, text content) and are AUTHORITATIVE. "
    "For any criterion listed here, adopt this PASS/FAIL verdict as fact rather than "
    "re-judging it by eye:"
)


def build_deck_data(digest: dict, md_slides: list) -> str:
    """The {{DECK_DATA}} block: the precomputed deterministic checks followed by the
    per-slide markitdown text. Visual judgments come from the slide images, so no
    per-shape geometry is emitted."""
    w, h = digest["slide_size_in"]
    lines = [
        f"DECK: {digest['file']}   slides: {digest['slide_count']}   size: {w} x {h} in",
        "", _PREAMBLE,
    ]

    cc = digest.get("computed_checks") or {}
    if cc:
        lines += ["", _RULE, _CHECKS_PREAMBLE]
        lines += [f"  - [{cid}] {res['verdict']}: {res['detail']}" for cid, res in cc.items()]

    lines += ["", _RULE, "DECK TEXT PER SLIDE (markitdown extraction):"]
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


# --------------------------------------------------------------------------- #
# Output: write rubric's computed result as a markdown report
# --------------------------------------------------------------------------- #
def write_prompt_log(prompt: str, digest, logs_dir="logs") -> str:
    os.makedirs(logs_dir, exist_ok=True)
    stem = os.path.splitext(digest["file"])[0]
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    path = os.path.join(logs_dir, f"{stem}-prompt-{ts}.md")
    with open(path, "w") as fh:
        fh.write(prompt + "\n")
    return path

# make the reports readable
def write_report_md(report, digest, dpi, leftovers, logs_dir="logs") -> str:
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
    image_blocks = _image_blocks(images)
    n_criteria = len(rubric.rubric)

    def _alignment_problem(out: OneShotOutput) -> str:
        """Describe any criterion_idx misalignment, or "" if the evaluations cover
        exactly {0..n-1} once each.

        The library maps each evaluation to a criterion purely by its returned
        `criterion_idx` (per_criterion_one_shot_grader.py). When the model skips a
        criterion or repeats/misnumbers an index, every later verdict silently shifts
        onto the WRONG criterion and the gap becomes "Evaluation not found". We catch
        that here — before the library maps it — so we can re-ask rather than score a
        scrambled report."""
        idxs = [e.criterion_idx for e in out.criteria_evaluations]
        expected = set(range(n_criteria))
        got = set(idxs)
        parts = []
        if expected - got:
            parts.append(f"missing indices {sorted(expected - got)}")
        if got - expected:
            parts.append(f"out-of-range indices {sorted(got - expected)}")
        dupes = sorted({i for i in idxs if idxs.count(i) > 1})
        if dupes:
            parts.append(f"duplicated indices {dupes}")
        if len(idxs) != n_criteria:
            parts.append(f"returned {len(idxs)} evaluations, expected {n_criteria}")
        return "; ".join(parts)

    # generate_fn for rubric's one-shot grader: one call judges every criterion.
    async def generate_fn(system_prompt: str, user_prompt: str, **kwargs) -> OneShotOutput:
        system = [
            {"type": "text", "text": system_prompt},
            {"type": "text", "text": prompt, "cache_control": {"type": "ephemeral"}},
        ]
        instruction = (
            "Above are the rendered slides of the deck under evaluation. The system "
            "block contains the full grading brief: the rubric criteria (numbered "
            "0,1,2,…) and the deck's geometry digest and per-slide text.\n\n"
            "GRADE THE DECK. For every criterion, decide MET or UNMET by inspecting "
            "the slide images (ground truth) together with the geometry/text digest, "
            "and return one evaluation per criterion using its 0-based index. Do not "
            "treat the digest as a 'response to echo' — it is the deck's own data. "
            f"There are exactly {n_criteria} criteria (indices 0..{n_criteria - 1}); "
            "return exactly one evaluation for each, with criterion_idx set to that "
            "criterion's own index, in order. Do not skip, merge, or renumber any. "
            "You must return a verdict for every criterion; never reply that there is "
            "nothing to evaluate."
        )
        fixup = ""  # appended on a re-ask after an alignment failure
        last_err = None
        out = None
        for attempt in range(1, 4):  # transient overloads/timeouts shouldn't kill a run
            try:
                resp = await client.messages.parse(
                    model=MODEL,
                    max_tokens=MAX_TOKENS,
                    system=system,
                    messages=[{"role": "user",
                               "content": image_blocks
                               + [{"type": "text", "text": instruction + fixup}]}],
                    output_format=OneShotOutput,
                )
                out = resp.parsed_output
                problem = _alignment_problem(out)
                if problem and attempt < 3:
                    # Verdicts would attach to the wrong criteria — re-ask with the
                    # specific defect named, rather than scoring a scrambled report.
                    print(f"  criterion_idx misaligned ({problem}); re-asking "
                          f"{attempt}/2…", file=sys.stderr)
                    fixup = (
                        f"\n\nYOUR PREVIOUS RESPONSE WAS MISALIGNED: {problem}. Each "
                        f"criterion's verdict is matched by its criterion_idx, so a "
                        f"wrong or missing index silently scores the wrong criterion. "
                        f"Return exactly {n_criteria} evaluations, one per criterion, "
                        f"each with criterion_idx equal to that criterion's index "
                        f"(0..{n_criteria - 1}), none skipped or repeated."
                    )
                    continue
                return out
            except Exception as e:  # noqa: BLE001 - retry any API/parse failure
                last_err = e
                if attempt < 3:
                    await asyncio.sleep(2 * attempt)
                    print(f"  grading call failed ({type(e).__name__}), "
                          f"retry {attempt}/2…", file=sys.stderr)
        if out is not None:
            return out  # last attempt, even if still imperfect — better than crashing
        raise last_err

    grader = PerCriterionOneShotGrader(generate_fn=generate_fn)
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
