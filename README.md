# pptx-grader

![Python](https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white)
![Model](https://img.shields.io/badge/judge-Claude%20Opus%204.8-d97757)
![python-pptx](https://img.shields.io/badge/built%20with-python--pptx-0a7e07)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

**A vision-led rubric grader for PowerPoint decks.** Point it at a `.pptx` and it
renders every slide to an image, extracts the per-slide text, and computes a handful
of exact data checks — then sends all of it to Claude in **one call** that scores
each weighted rubric criterion MET/UNMET and returns a single weighted score with a
per-criterion report.

The design principle: **let the eye judge what's visual, let code judge what's exact.**
Defects you can see — overlaps, off-slide bleed, text overflow, misalignment, crowded
margins — are judged by the vision model from the rendered slides (ground truth).
Only the things the eye *can't* read reliably — font-family count, exact point-size
hierarchy, and the number of distinct brand hues — are computed deterministically in
code and handed to the model as authoritative fact, so those verdicts stay stable
run-to-run.

## How it works

For a given `deck.pptx`:

1. **`check_layout.py`** reads the .pptx XML and computes four deterministic PASS/FAIL
   checks that are pure data lookups (≤2 font families, title>body>footnote size
   hierarchy, one size per level, ≤6 non-neutral hues). No geometry/collision
   heuristics — visual layout is left to the model.
2. **`render.py`** renders each slide to a JPG (LibreOffice → PDF → `pdftoppm`) and
   extracts each slide's text with markitdown (also flagging leftover placeholder
   text like `lorem ipsum`).
3. **`grade_pptx.py`** fills the `source/grader.md` prompt — `{{RUBRIC}}` (criteria
   from `source/rubric.yaml`) and `{{DECK_DATA}}` (the precomputed checks + slide
   text) — attaches the slide images, and runs one `rubric` grading call. The model
   returns per-criterion MET/UNMET verdicts; `rubric` computes the weighted score.
   Returned verdicts are validated for criterion alignment and re-asked if the model
   misnumbers, so a verdict can never silently attach to the wrong criterion.
4. The result is written to `logs/<deck>-<timestamp>.md`.

The two markdown files are the interface: **`source/grader.md`** in (the prompt) and
the report **out**. The criteria + weights live in **`source/rubric.yaml`** (the
single source of truth that drives the score).

## Setup

```bash
# system deps for rendering
brew install --cask libreoffice
brew install poppler

# python env
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -e ~/ML/rubric        # the scorer (local checkout)

# API key — put it in a .env file beside grade_pptx.py
echo 'ANTHROPIC_API_KEY=sk-ant-...' > .env
```

## Usage

```bash
python grade_pptx.py path/to/deck.pptx
```

Options:

- `--rubric PATH` — rubric YAML (default: `source/rubric.yaml`)
- `--grader PATH` — prompt markdown (default: `source/grader.md`)
- `--dpi N` — render resolution (default 150)
- `--keep-images` — keep the rendered JPGs under `--logs-dir`
- `--logs-dir DIR` — where the report goes (default `logs/`)
- `--show-digest` — render + assemble the prompt and print it; no grading API call

Rendering is **required** — if LibreOffice/poppler are missing the grader prints an
install hint and exits 2 (it never silently grades without the images).

## Files

| File | Role |
| --- | --- |
| `grade_pptx.py` | entry point: assembles the prompt, calls the model, writes the report |
| `check_layout.py` | the four deterministic data-only checks from the .pptx XML |
| `render.py` | slides → JPGs + per-slide markitdown text |
| `source/grader.md` | the prose grader prompt (`{{RUBRIC}}` + `{{DECK_DATA}}` placeholders) |
| `source/rubric.yaml` | the weighted criteria (source of truth for the score) |
