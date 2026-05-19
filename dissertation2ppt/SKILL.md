---
name: dissertation2ppt
description: Create an editable Chinese dissertation or thesis defense PowerPoint from degree thesis materials, preferably a Word .docx with original chapter figures. Use when Codex needs to act as a top Chinese engineering defense expert and produce a 16:9, reference-quality, mostly pure-white 18-25 slide PPTX for bachelor, master, or doctoral theses, with truly centered cover text, clean chapter-title hierarchy, no persistent chapter navigation, no visible internal reasoning labels, layout-diverse evidence-first pages, large readable thesis figures, full useful tables, output/ppt_plan.md planning, extracted thesis images, and strict build-time QA for editability, density, text overflow, title consistency, and image placement.
---

# Role

Act as a top Chinese university engineering defense expert, especially for control science, computer vision, robotics, automation, AI, sensing, reconstruction, segmentation, detection, and related fields. Convert a complex degree dissertation into an evidence-heavy, visually mature defense PPT that can be shown directly, not a generic report deck, Q&A forecast, or speaker script.

Default to Simplified Chinese slides. Preserve English only for proper nouns, algorithms, models, datasets, abbreviations, software, equations, metrics, journal names, and technical terms normally used in English.

# Output Contract

Every normal run must create:

- `output/ppt_plan.md` before the deck.
- A final editable `.pptx`, normally `output/dissertation_defense.pptx`.
- `output/assets/figures/` for selected thesis figures and auxiliary visuals.
- `output/asset_manifest.md` when figures, logos, screenshots, publications, or downloaded assets are used.
- Optional `output/qa_report.md`.

Hard requirements:

- 16:9 widescreen.
- 18-25 slides by default.
- Pure white slide canvas by default. Use school-color lines, top headers, section dividers, badges, cards, shapes, and subtle diagrams on top of white; do not use full-slide photo, gradient, generated, or decorative colored backgrounds.
- Editable PPT objects for text, diagrams, tables, and charts. Do not flatten slides into screenshots.
- A clean school identity: logo on content slides, accents whose hue stays close to the school emblem or official school color, and consistent fonts.
- No large empty areas, no text overflow.
- Cover text boxes and the paragraphs inside those boxes must both be horizontally centered; centering only the box is not enough.
- Do not use persistent chapter navigation strips or bottom-left progress navigation. A page marker is fine; a repeated `01 02 03 04` navigation rail is not.
- The final PPTX must be immediately presentable. Do not add probable questions, backup answers, defense-logic pages, rehearsed scripts, hidden speaker notes, or visible internal reasoning labels unless the user explicitly requests them.

# Reference-Quality Design Bar

Strong Chinese degree-defense samples share these traits. Treat them as build requirements, not optional polish:

- Formal cover: school logo/name, centered thesis title, centered author/advisor/major/date metadata, and one restrained school-color title band or rule.
- Stable identity: content slides use a consistent logo position, top-left chapter title, optional subtitle, page marker, school-color accent, and thin separator rules.
- Clear chapter rhythm: 5-7 top-level parts at most, with simple section dividers using large part numbers and concise section names.
- Evidence-first composition: thesis figures, method diagrams, visual comparisons, maps, architecture blocks, tables, and metrics occupy the main canvas; bullets explain evidence instead of replacing it.
- Local rhythm for each core work: motivation/problem -> method/framework -> experiment/data -> quantitative/qualitative result -> analysis or chapter conclusion.
- Contact-sheet strength: thumbnail view should show varied macro layouts, readable titles, visible proof objects, large key figures, and a coherent white-and-school-color system.
- Restrained emphasis: use red/blue highlight, arrows, callouts, and boxed conclusions only for key contributions or results; avoid decorative cards and filler badges.
- Two valid pacing modes: compact defense decks use 18-25 dense slides; long thesis-showcase decks use 30-60+ slides only when the thesis/reference/user clearly calls for it, with repeated progress pages or section dividers to prevent overload.
- Cover families seen in strong samples: white institutional cover with a wide school-color band, full school-color title cover, or minimalist white cover with thin rules when the user supplies or requests that format.

# Toolchain Policy

Use `python-pptx` for slide authoring and PPTX-safe editing. If it is missing in the active Python environment, install the package before creating the deck:

```bash
python -m pip install python-pptx
```

Use `python-pptx` for editable slide text, shapes, tables, charts, images, slide size, and reopen checks. Use `zipfile`/XML inspection or the bundled validation script for additional structural QA. Do not use screenshot-only generation as the primary PPTX workflow.


# Chapter Design Rule

Use a defense-chapter structure similar to strong Chinese engineering thesis decks:

- Cover.
- `目录` or `汇报内容`, never `答辩提纲`.
- Part 01: `绪论` or `研究背景与意义`, including background, significance, research status, problem/gap, research content, innovation overview, and thesis chapter arrangement when useful.
- Part 02: `相关理论与技术` or `理论基础与关键技术`, only when the thesis depends on technical foundations the committee must understand.
- Part 03-04 or Part 03-05: core research works. Each major work should have its own section divider and then follow a local rhythm: data/problem -> method/framework -> quantitative experiment -> qualitative visualization -> ablation/robustness/analysis.
- Part after core works: `总结与展望`, combining main conclusions, contributions, limitations, and future work.
- Optional final academic output part: `攻读学位期间科研成果` or `论文发表与科研成果`, when publications, patents, software copyrights, datasets, awards, or projects appear in the dissertation.
- Closing: acknowledgements and `敬请各位老师批评指正`.

For an 18-25 slide deck, avoid too many tiny sections. Use 5-7 top-level parts at most. When core work is the thesis center, give each work 3-5 slides and compress theory/background. When the thesis is method-heavy, split each work into `框架/方法`, `实验设置`, `结果对比`, and `消融分析`.

# Source Extraction

Prefer `.docx` because it preserves headings, captions, tables, formulas, and embedded original figures. If only PDF is available, crop selected figures carefully and record lower editability.

Extract:

- title, author, advisor, school, college, major, degree level, defense date
- abstract, keywords, table of contents, chapter hierarchy
- background, significance, research status, gap, problem definition
- research objectives, contents, technical route, datasets, devices, metrics
- chapter-level methods, experiments, results, conclusions, innovations
- all relevant figures, tables, formulas, and captions
- publications, patents, software copyrights, awards, projects, datasets, or other thesis outputs
- conclusion, limitations, and future work

Evidence selection:

- Use original thesis figures as evidence, not decoration.
- Prefer method/workflow diagrams, system architecture, experimental setups, qualitative comparisons, ablations, robustness checks, and quantitative summary tables.
- For computer vision, segmentation, detection, 3D reconstruction, point clouds, remote sensing, robotics, simulation, and industrial inspection, prioritize visual comparison panels with metric tables nearby.
- Keep formulas rare: one central objective/loss/energy/governing equation only when it explains the contribution.

# Required Plan

Before building the PPTX, write `output/ppt_plan.md` with:

- Thesis metadata and defense assumptions.
- Content diagnosis: problem, route, core works, evidence strength, and available thesis proof objects.
- Branding plan: logo source, school color and fonts.
- Numbered outline: the exact part numbers used in slide headers.
- Reference-quality design plan: cover grammar, section divider grammar, header/subtitle grammar, page marker grammar, explicit no-navigation rule, and at least 6 macro-layout families for a 20-slide deck.
- Slide table for the planned deck: number, Chinese title, display purpose, source chapter, primary figure/table/proof object, and layout pattern.
- Publication/output plan: list thesis publications or outputs and where they appear; if none are found, state that.
- Table plan: list important original tables and whether to recreate fully, split across slides, convert to chart, or keep as cropped image.
- Figure plan: state which figures are inserted directly, which need a light container, crop, label, or callout, and which key figures must be shown large instead of as thumbnails.
- Layout density check: identify slides at risk of being empty, bottom-blank, text-heavy, table-heavy, figure-dense, or repetitive.
- Build-time quality plan: cover paragraph centering, text overflow, figure readability, white canvas, editability, title consistency, no chapter navigation, contact-sheet rhythm, layout diversity, and forbidden internal/audit text.

Do not plan probable questions, backup answers, committee concerns, rehearsed defense logic, speaking scripts, or speaker notes. Keep planning/audit information in `output/ppt_plan.md`, `output/asset_manifest.md`, and `output/qa_report.md`, not on the expert-facing slides.

If the user asks to approve the plan first, stop after `output/ppt_plan.md`; otherwise continue.

# Cover Slide

The cover should look formal, calm, and centered.

- Center the thesis title horizontally and visually near the middle of the slide.
- Put author/report presenter, major, advisor, college/school, and date below the title, also centered.
- For every cover textbox containing title, author, advisor, major, college/school, or date, set the paragraph alignment to center. Do not rely only on textbox geometry.
- Use the school logo and school name near the top or above the title; keep them aligned and balanced.
- Use a pure white background with one restrained school-color band, rule, or block if needed.
- Match the accent color roughly to the school emblem or official school color.
- Do not use a side-heavy hero layout on the cover unless the user supplied an official template requiring it.
- Pick one cover family and execute it decisively: wide horizontal title band (unless otherwise specified, priority shall prevail), full school-color cover, minimalist white cover with thin rules, or source-template cover. Do not mix several cover grammars.

# Header and Section Titles

Do not use meaningless headers such as "研究成果一" without context. Follow the planned outline numbering.

On content slides, use:

- One and only one top-left chapter number. Prefer integrating it into the title, such as `4 基于条件引导的条件生成式点云补全`.
- A stable chapter title for every slide inside that chapter. Do not change the chapter title to a local subtopic.
- A short, specific subtitle below the separator line when needed, such as `局部条件扩散补全框架`; set it smaller and lighter than the chapter title.
- The school logo in the top-right.
- A separator line under the header, colored to match the school's institutional color.
- A small page marker may appear in a consistent corner, but do not add chapter navigation or progress-number strips.

对于核心内容幻灯片，章节标题在每个章节中保持一致，而不同部分的子章节标题应同时阐明工作内容和核心论点，例如：

- `基于体素扩散的大范围场景补全`
- `无提示气体分割模型设计`
- `多尺度几何与空间位置协同调制`

Header mistakes to avoid:

- Do not combine a number badge `2` with a title that also starts with `2`; duplicate numbers make the hierarchy look broken.
- Do not use bottom-left or side chapter navigation such as `01 02 03 04`, even if the current chapter is highlighted.
- Do not color same-level chapter numbers inconsistently.

Do not put internal workflow labels on slides, including `资料来源`, `素材来源`, `答辩提纲`, `归纳出的答辩主线`, `答辩主线`, `注意事项`, `生成说明`, `AI生成`, `来自论文原文`, `截图自`, or `output/ppt_plan.md`. The committee should see only defense content.

# Layout System

Keep the background white, but make the layout rich through structure.

Prefer varied layouts:

- Add text descriptions to fill larger blank areas, with keywords bolded or rendered in an accent color.
- Arrow framework: 4-6 arrow-linked steps with concise labels, useful for technical route, model pipeline, data flow, and experiment procedure.
- Arrow bullets: vertical or horizontal arrow-led points where each arrow starts with a bold keyword and a short source-grounded explanation.
- Use dashed or solid rectangular borders in different colors to enclose content blocks, creating visual separation from surrounding elements.
- Use dark-colored panels with white text as local module anchors, not as persistent chapter navigation.
- Large original figure with side interpretation rail.
- Two-column method/result evidence.
- Three-card `问题 / 思路 / 方法` or `数据 / 模型 / 结果`.
- Process route with 4-6 connected steps.
- Matrix layout for research status, technical route, or ablation.
- Comparison plate for qualitative results.
- Full or split table layout for detailed experimental data.
- Figure plus mini table plus short conclusion.
- Publications slide with 2-4 publication cards and citation/status details.
- Section divider page with a large part number, short Chinese section title, a simple generated pattern suitable for use as a divider background, and ample white space.
- Background/problem collage with timeline, policy/event strip, example images, and one clear problem statement.
- Framework map with a dominant architecture/flow diagram and 3-5 labeled stages.
- Quantitative result slide with table/chart plus a nearby qualitative example or conclusion callout.
- Qualitative comparison plate with aligned method labels, identical crop sizes, and a highlighted student method.
- Summary/contribution slide with 3-4 numbered conclusions, each tied to a chapter or evidence object.

For a 20-slide deck, use at least 6 distinct macro-layout families. Do not let 3 consecutive content slides share the same `title + bullets + image`, `title + boxed cards`, or `left dark block + right text rectangle` composition. The contact sheet should look authored before the text is read.

Anti-repetition rules:

- The `left dark block + right rectangular text area` layout is allowed, but it must not become the deck default. In an 18-25 slide deck, use it on no more than 3 content slides and never on adjacent slides.
- Do not solve every sparse page with the same bottom conclusion bar. Prefer a larger thesis figure, a second evidence object, an arrow framework, a table slice, or a structured comparison.
- Section dividers may be visually related, but content slides should rotate layout families based on evidence type.

Avoid relying on a bottom conclusion bar as the default fix for empty slides. Use a bottom bar only when it carries a real takeaway and does not become repetitive. When a slide looks empty, first add evidence, structure, a diagram, a table, or a source-derived short description.

No empty lower half:

- Main content should occupy most of the central canvas.
- If the bottom third is blank, resize/reposition evidence, add another source-derived figure/table, or switch to a fuller layout.
- A slide with only bullets is usually unacceptable unless it is a formal outline, limitations, or closing slide.

Figure scale rules:

- Key method diagrams, visual comparison plates, reconstruction/segmentation/detection examples, and architecture figures should usually occupy 45-70% of the slide area.
- For visual comparison slides, the comparison image should be the main object, not a small thumbnail in a corner. Use large aligned panels with method labels and nearby metrics/callouts.
- Avoid placing evidence figures below roughly one quarter of slide width unless they are intentionally secondary thumbnails in a comparison grid.

# Text Rules

- One slide, one defense point.
- Use 2-4 concise bullets or short phrases; avoid thesis paragraphs.
- Write brief, logical descriptions derived from the thesis source, not generic filler.
- For result slides, the subtitle should lead with the argument.
- Do not create speaker notes, hidden scripts, Q&A pages, `答辩逻辑` pages, `归纳出的答辩主线` pages, or visible reminders about what the presenter should notice unless explicitly requested.
- Body text should be at least 20 pt, or 18 pt used sparingly.
- Never allow text to exceed its background box. If text overflows, shorten it, enlarge the box, reduce hierarchy safely, or split the content across slides.

# Tables

Tables show workload and experimental rigor. Do not over-summarize important original data.

- Extract relevant original table values as fully as practical.
- If the values are clearly legible, render the table as a professional three-line (booktabs-style) table.
- Preserve important rows, columns, metrics, datasets, baselines, and method names.
- Use highlight color, bold text, or callout arrows to mark the student's method and key improvements.
- Convert to charts only when the chart communicates the result better and the complete table is still available nearby or in a following slide.
- Use table screenshots only when recreating risks transcription errors or the original layout is itself important.
- For result slides, prefer the sample pattern `quantitative table/chart + visual proof + red/blue highlight` over isolated tables. The student's method, best value, or key improvement should be visually findable in under 3 seconds.

# Figures and Image Placement

Original thesis figures should often be pasted directly.

- Insert clean figures directly on the white slide when the original figure already has a white or transparent background.
- Do not automatically add a card, shadow, gray panel, or colored rectangle behind every figure.
- Add a light container only when it improves legibility, separates a busy figure from nearby content, or matches an intentional comparison layout.
- Crop away thesis page margins, redundant captions, and unrelated panels.
- Preserve aspect ratio. Never stretch detection, segmentation, reconstruction, or chart images.
- Add minimal labels or callouts only where they help the defense committee read the evidence.
- For visual comparisons, keep method labels aligned and use original panel ordering when possible.
- When the thesis has many qualitative examples, create comparison plates instead of shrinking unrelated images into a collage. Keep equal panel sizes, consistent gutters, and direct labels.

Use generated visuals only for auxiliary layout support, such as simple process icons, neutral line patterns, or abstract schematics. Do not generate logos, experimental evidence, result images, tables, or fake comparison visuals.

# Publications and Outputs

If the dissertation contains publications or outputs, include them in the PPT.

Create a dedicated slide or integrate into the contributions section:

- Paper title, venue/journal, year, author position, status, DOI if available.
- Patents, software copyrights, datasets, awards, or projects if listed.
- Keep this slide factual and source-grounded.
- If outputs are many, group by `论文 / 专利软著 / 项目数据 / 获奖`.
- And for this section, font usage is individually restricted to SimSun for Chinese text and Times New Roman for English text.

# Build-Time Quality Requirements

Build with the available PPTX toolchain. Prefer editable native objects for text, shapes, diagrams, tables, and simple charts. Use extracted figures as images only when the original visual is the evidence.

While authoring the deck, satisfy these requirements slide by slide:

1. 16:9 slide size and 18-25 slides unless the user or thesis volume requires a different count.
2. White slide canvas; school-color bands, headers, dividers, and local module anchors are foreground structure, not decorative full-slide backgrounds or persistent navigation.
3. Centered cover title and centered metadata, with paragraph alignment set to center inside each textbox.
4. Numbered headers, consistent logo position, subtitle placement, and page marker; no persistent chapter navigation.
5. Meaningful section labels; never use meaningless labels such as `研究成果一` without the work topic.
6. Publications/outputs included when present in the thesis.
7. Figures preserve aspect ratio, are not stretched, and are not automatically placed on gray cards or decorative panels.
8. Text stays inside boxes, has visible padding, has correct alignment, and does not overlap other objects.
9. The deck remains editable and is not a sequence of full-slide screenshots.
10. The closing slide includes `敬请各位老师批评指正` or an equivalent defense closing.
11. No visible internal/audit terms, Q&A forecasts, backup-answer pages, speaker scripts, `归纳出的答辩主线`, or defense-logic labels.
12. The deck has a consistent cover/header/page-marker system, except for intentional cover, divider, appendix, and closing variants.
13. Result slides pair claims with proof: a table/chart alone is weak unless the thesis result is purely numeric; a qualitative image alone is weak unless labels and metrics nearby explain it.
14. Thumbnail/contact-sheet review shows at least 6 macro-layout families in compact mode, with no adjacent repeated `left dark block + right rectangle` layouts.
15. No content slide leaves the lower third obviously empty; fill it with source-derived proof, larger figures, structured arrows, tables, or a fuller layout.
16. Key evidence figures are large enough to read in slideshow mode, especially visual comparison images.

# Final Verification Is Report-Only

After the final PPTX is created, run an audit pass. For intentional long decks, pass the planned range to the script instead of accepting the default 18-25 warning:

```bash
scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx
# or:
scripts/validate_dissertation_ppt.py --min-slides 30 --max-slides 66 output/dissertation_defense.pptx
```

Then reopen the `.pptx` and inspect rendered slide previews when rendering tools are available.

If the script or visual inspection finds problems, write `output/qa_report.md` and report the issues to the user. Do not modify the already-created PPTX during the final verification step. A repair pass requires a new explicit instruction from the user or a fresh build run.
