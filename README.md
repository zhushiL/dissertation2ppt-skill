# dissertation2ppt Skill

`dissertation2ppt` is a Codex skill for turning Chinese degree theses into editable, evidence-heavy defense PowerPoint decks. It is designed for undergraduate, master's, and doctoral dissertation defenses where the output must be structured, source-grounded, visually mature, and ready for further editing in PowerPoint or WPS.

The skill is especially tuned for engineering theses in control science, automation, robotics, computer vision, AI, sensing, reconstruction, detection, segmentation, and related fields.

## What It Produces

For a normal run, the skill creates a single organized `output/` folder:

```text
output/
├── ppt_plan.md
├── dissertation_defense.pptx
├── asset_manifest.md
├── qa_report.md
└── assets/
    └── figures/
```

The default deck is:

- 16:9 widescreen.
- 18-25 slides unless the user asks for a different scope.
- Simplified Chinese by default.
- A formal `目录` slide.
- White canvas with restrained school-color structure.
- Editable PPTX objects for text, tables, diagrams, and charts.
- Original thesis figures used as evidence, not decoration.
- Consistent bottom-right page markers.
- Separate `总结与展望` and closing pages.

## Repository Layout

```text
.
├── README.md
└── dissertation2ppt/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    └── scripts/
        └── validate_dissertation_ppt.py
```

- `dissertation2ppt/SKILL.md`: canonical Codex skill definition.
- `dissertation2ppt/agents/openai.yaml`: skill display metadata.
- `dissertation2ppt/scripts/validate_dissertation_ppt.py`: standard-library PPTX structural validator.

## Install

Clone this repository into your Codex skills directory:

```bash
git clone https://github.com/zhushiL/dissertation2ppt-skill.git ~/.codex/skills/dissertation2ppt
```

If you already cloned the repository elsewhere, copy the skill folder:

```bash
mkdir -p ~/.codex/skills
cp -R dissertation2ppt ~/.codex/skills/dissertation2ppt
```

Then trigger it in Codex with `$dissertation2ppt`.

## Requirements

- Codex with local skills enabled.
- Python 3.10+.
- `python-pptx` for PPTX generation:

```bash
python -m pip install python-pptx
```

The validation script only uses the Python standard library.

## Usage

Best input is a Word `.docx` thesis because it preserves headings, captions, tables, formulas, and embedded figures.

Example:

```text
Use $dissertation2ppt to create a Chinese master thesis defense PPT from /path/to/thesis.docx.
```

Plan-first workflow:

```text
Use $dissertation2ppt to create output/ppt_plan.md from /path/to/thesis.docx, then stop for my approval.
```

Template-aware workflow:

```text
Use $dissertation2ppt with /path/to/thesis.docx and follow /path/to/template.pptx as the visual template.
```

## Workflow

The skill follows a defense-first workflow:

1. Extract thesis metadata, abstract, table of contents, chapter hierarchy, figures, tables, formulas, experiments, conclusions, and research outputs.
2. Verify school identity assets from the thesis or user-provided files before using logos.
3. Write `output/ppt_plan.md` before building the deck.
4. Build a slide-level claim spine instead of copying the thesis chapter by chapter.
5. Compose evidence-first layouts with readable original figures, recreated tables, concise interpretation rails, and varied macro layouts.
6. Generate an editable `.pptx` using `python-pptx`.
7. Record used assets in `output/asset_manifest.md`.
8. Run structural QA and write `output/qa_report.md` when issues remain.

## Validation

After generating a deck:

```bash
python dissertation2ppt/scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx
```

JSON report:

```bash
python dissertation2ppt/scripts/validate_dissertation_ppt.py --json output/dissertation_defense.pptx
```

The validator checks, among other things:

- Slide count and 16:9 dimensions.
- Full-slide screenshot risks.
- Non-white background risks.
- Native text presence.
- Cover paragraph centering.
- Page marker position.
- Duplicate header numbering.
- Persistent navigation-like number strips.
- Possible text overflow and long unwrapped lines.
- Small evidence image warnings.
- Internal/audit terms that should not appear on defense slides.

## Design Rules

Core presentation rules enforced by the skill:

- One slide, one defense point.
- `目录`, not `答辩提纲`, for the agenda page.
- No persistent chapter navigation strips.
- No visible internal planning labels, generation notes, speaker scripts, Q&A forecasts, or backup-answer pages.
- Normal slide text uses Microsoft YaHei / `微软雅黑`.
- Publications/output slides use SimSun / `宋体` for Chinese and Times New Roman for English.
- Key method diagrams, visual comparisons, and result figures must be large enough for a defense room.
- Tables should preserve important rows, baselines, metrics, and student-method highlights.
- Source figures, logos, screenshots, and external assets must be recorded in the asset manifest.
- If a verified school logo is unavailable, the workflow must pause for a user-provided logo or explicit logo-free approval.

## Scope And Limitations

- The final deck still needs human review, especially for author names, advisor names, defense dates, publication status, equations, and numeric values.
- `.docx` input is strongly preferred. PDF-only workflows usually require more cropping and manual verification.
- The skill must not fabricate logos, experimental evidence, tables, results, datasets, papers, patents, or partner marks.
- Do not use this repository to store real unpublished theses, private data, or restricted research materials.

## Development Notes

Keep the skill self-contained:

- Update `dissertation2ppt/SKILL.md` as the source of truth.
- Keep helper scripts under `dissertation2ppt/scripts/`.
- Do not commit generated `output/` folders, thesis source files, extracted assets, local previews, or personal `.codex` runtime files.
- Run the validator when changing deck QA requirements.

## License

No license file is included yet. Add a `LICENSE` file before redistributing or accepting external contributions.
