# pptx-grader

Grades a PowerPoint deck against a weighted rubric. The deck isn't judged from
text alone — every slide is rendered to an image, its text is extracted, and its
geometry is measured, then all three signals are sent to the model in **one call**
that scores every rubric criterion at once.

## How it works

For a given `deck.pptx`:

1. **`check_layout.py`** reads the .pptx XML and builds a geometry digest: every
   shape's position (in inches), fonts/colors, and pre-computed layout flags
   (`alignment_issues`, `frame_violations`, `collisions`).
2. **`render.py`** renders each slide to a JPG (LibreOffice → PDF → `pdftoppm`) and
   extracts each slide's text with markitdown (also flagging leftover placeholder
   text like `lorem ipsum`).
3. **`grade_pptx.py`** fills the `grader.md` prompt — `{{RUBRIC}}` (criteria from
   `powerpoint-rubric.yaml`) and `{{DECK_DATA}}` (digest + slide text) — attaches
   the slide images, and runs one `rubric` grading call. The model returns
   per-criterion MET/UNMET verdicts; `rubric` computes the weighted score.
4. The result is written to `logs/<deck>-<timestamp>.md`.

The two markdown files are the interface: **`grader.md`** in (the prompt) and the
report **out**. The criteria + weights live in **`powerpoint-rubric.yaml`** (the
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

- `--rubric PATH` — rubric YAML (default: `powerpoint-rubric.yaml` beside the script)
- `--grader PATH` — prompt markdown (default: `grader.md` beside the script)
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
| `check_layout.py` | geometry digest from the .pptx XML (no computer vision) |
| `render.py` | slides → JPGs + per-slide markitdown text |
| `grader.md` | the prose grader prompt (`{{RUBRIC}}` + `{{DECK_DATA}}` placeholders) |
| `powerpoint-rubric.yaml` | the weighted criteria (source of truth for the score) |
