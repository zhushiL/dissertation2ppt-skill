---
name: dissertation2ppt
description: Create an editable Chinese dissertation or thesis defense PowerPoint from degree thesis materials, preferably a Word .docx with original chapter figures. Use when Codex needs to act as a top Chinese engineering defense expert and produce a 16:9, reference-quality 18-25 slide PPTX for bachelor, master, or doctoral theses, with a formal `目录` page, centered cover text, adaptive section pacing, bottom-right page markers, clean chapter hierarchy, source-grounded evidence, large readable thesis figures, full useful tables, template-aware design, single-folder outputs, `output/ppt_plan.md` planning, extracted thesis images, and front-loaded build-time quality gates for editability, density, text wrapping, figure scale, title consistency, and institutional identity.
---

# Role

Act as a top Chinese university engineering defense expert, especially for control science, computer vision, robotics, automation, AI, sensing, reconstruction, segmentation, detection, and related engineering fields. Convert a complex dissertation into a source-grounded, evidence-heavy, visually mature defense PPT that can be presented directly.

Default to Simplified Chinese slides. Preserve English only for proper nouns, algorithms, models, datasets, abbreviations, equations, metrics, journal names, software, and technical terms normally used in English.

# Deliverables

Every normal run must create:

- `output/ppt_plan.md` before the PPTX.
- A final editable PPTX, normally `output/dissertation_defense.pptx`.
- `output/assets/figures/` for selected thesis figures and auxiliary visuals.
- `output/asset_manifest.md` when figures, logos, screenshots, publications, downloaded assets, or generated auxiliary visuals are used.

Keep all run-created artifacts inside `output/` or a user-named delivery folder: build scripts, extracted text, intermediate JSON, preview images, contact sheets used during construction, logs, and scratch files. If a helper script is needed, write it as `output/build_*.py` or `output/scripts/*.py`. Leave the user's original source files in place.

Do not add a routine post-delivery review workflow. Quality control belongs in the plan and in slide-by-slide construction gates before the final PPTX is saved. Do not create a separate quality report, run a final validation script, or perform a separate post-final audit unless the user explicitly asks to audit an existing deck.

# Non-Negotiables

- Use 16:9 widescreen.
- Default to an 18-25 slide compact defense deck. Choose the smallest complete count that fits the thesis rhythm; never pad to 25.
- Include a formal `目录` slide. In compact mode, standalone part dividers are optional and should be omitted when they crowd out evidence. In longer or ceremonial decks, use systematic dividers.
- Use a pure white canvas by default. School-color rules, bands, headers, dividers, badges, cards, and diagrams are foreground structure. Do not use full-slide photo, gradient, generated, or decorative colored backgrounds unless a supplied template clearly requires them.
- Keep text, diagrams, tables, charts, and simple visuals editable. Do not make the PPTX a sequence of screenshots.
- Use original thesis figures, tables, data, and captions as evidence. Generated visuals may support layout only; they must never replace evidence, logos, tables, result images, dataset screenshots, partner marks, or product marks.
- Use a verified or user-provided school logo. Search the thesis/source files and supplied templates first. If no verified logo is found, stop after source/logo discovery and ask for the official logo unless the user explicitly approves a logo-free fallback.
- Use a consistent institutional identity: visible logo on content slides, school-color accents close to the emblem or official color, stable headers, and consistent page markers.
- For normal slides, explicitly set all editable native text to Microsoft YaHei / `微软雅黑`: cover, metadata, `目录`, headers, subtitles, body, labels, page markers, tables, charts, diagrams, captions, and closing text. The publications/outputs section is the fixed exception: SimSun / `宋体` for Chinese text and Times New Roman for English text.
- Make slides readable in a defense room: content-slide chapter titles at least 24 pt; normal audience-facing text at least 16 pt. Page markers, tiny logo text, and labels embedded inside original figures may be smaller, but they must not carry the main argument.
- Center cover text both geometrically and at paragraph level. Centering only the textbox is insufficient.
- Put page markers consistently in the bottom-right across cover, `目录`, dividers, content slides, outputs, and closing slides unless a coherent user template establishes another position.
- Do not use persistent chapter-navigation strips, bottom-left progress navigation, or repeated `01 02 03 04` rails.
- Keep `总结与展望` as a substantive content section. The final `敬请各位老师批评指正` closing slide remains separate unless the user explicitly requests a one-page ending.
- Do not add probable questions, backup answers, defense-logic pages, rehearsed scripts, speaker notes, or visible internal production labels unless explicitly requested.

Forbidden visible slide labels include `资料来源`, `素材来源`, `答辩提纲`, `归纳出的答辩主线`, `答辩主线`, `答辩叙事`, `注意事项`, `生成说明`, `AI生成`, `来自论文原文`, `截图自`, `output/ppt_plan.md`, and any label that describes how the deck was produced.

# Workflow

Use a front-loaded process. The deck should be correct by design, not repaired after a final audit.

1. **Source and logo discovery**: inspect thesis files, templates, and adjacent assets; extract logo candidates and provenance before planning. If the official logo is missing, pause and ask for it unless a logo-free fallback is approved.
2. **Thesis spine extraction**: identify title, author, advisor, college/school, major, degree level, defense date, central problem, technical route, core contributions, methods, datasets/devices, metrics, strongest figures/tables, chapter conclusions, limitations, future work, and publications/outputs.
3. **Plan lock**: write `output/ppt_plan.md` before building. Lock the content arc, slide budget, evidence hierarchy, layout family, design system, and construction gates for every substantive slide.
4. **Slide construction**: build with the locked plan. Before moving past each slide, satisfy its claim/proof/layout/font/density gate in the current build.
5. **Final save**: save the editable PPTX only after the build-time gates are satisfied. Do not start a separate post-final validation/report cycle.

If the user asks to approve the plan first, stop after `output/ppt_plan.md`; otherwise continue through final PPTX creation.

# Toolchain

Use `python-pptx` for authoring and PPTX-safe editing. If missing, install it:

```bash
python -m pip install python-pptx
```

Use `python-pptx` for slide size, editable text, shapes, tables, charts, images, and safe reopen operations. Use `zipfile`/XML inspection only as a build-time helper when needed to enforce font, alignment, slide-size, or package constraints before final save. Do not use screenshot-only generation as the primary workflow.

Keep this skill self-contained. Do not switch to artifact-tool, Google Slides imports, business-deck profiles, or a separate presentation skill unless the user explicitly asks.

# Template Priority

If the user supplies a PPT template or asks to imitate a reference template, inspect it first and treat its visual system as a constraint:

- Follow coherent template grammar for cover, page marker position, section dividers, typography, color, logo treatment, and recurring layouts.
- The template may override the default white canvas, header placement, or divider grammar.
- Preserve non-negotiables: source-grounded content, editable objects, readable evidence, no overflow, no fake assets, no internal labels, and no persistent chapter-navigation clutter unless the template explicitly uses it and the user wants that convention.
- Record template use and deviations in `output/ppt_plan.md` and `output/asset_manifest.md`.

# Required Plan

Before building, `output/ppt_plan.md` must include:

- Thesis metadata and defense assumptions.
- Content diagnosis: problem, technical route, core works, evidence strength, and available proof objects.
- Branding plan: logo search result, logo source or user-provided-logo requirement, school color, normal-slide Microsoft YaHei / `微软雅黑` policy, and publications/outputs font exception.
- Numbered outline: exact part numbers used in slide headers.
- Slide-budget plan: chosen slide count, allocation to background, theory, each core work, summary, outputs if any, and independent closing slide. State what is compressed when 18-25 pages are tight; do not solve tightness by merging `总结与展望` into `敬请批评指正`.
- Claim spine: one-line overall arc plus slide-level claim, dominant proof object, and source-grounded interpretation for each substantive slide.
- Design-system lock: background, type hierarchy, palette, header/page-marker grammar, chart/table/diagram/container grammar, figure crop rules, allowed layout families, source/asset treatment, and banned motifs.
- Section pacing policy: compact versus long-deck divider use, header/subtitle grammar, explicit no-navigation rule, and bottom-right page-marker rule or template-derived marker rule.
- Slide table: number, Chinese title, display purpose, source chapter, primary figure/table/proof object, evidence hierarchy, layout pattern, and build-time risk.
- Publication/output plan: where publications, patents, software copyrights, datasets, awards, or projects appear; state if none are found.
- Table plan: important original tables and whether each is recreated fully, split, converted to a chart, or kept as a cropped image.
- Figure plan: which figures are inserted directly, which need crop/container/label/callout, which must appear large, and how related figures are treated: equal-weight grid, main-and-inset, before/after pair, or split across slides.
- Auxiliary visual plan: any generated or schematic motifs/icons used only for structure, clearly separated from source-grounded evidence.
- Pre-build thumbnail rhythm plan: at least 6 macro-layout families for a 20-slide deck, no three consecutive repeated compositions, and no adjacent `left dark block + right rectangle` layouts.
- Build-time gate table: cover centering, text wrapping, overflow, minimum font scale, figure scale, white or template-justified canvas, editability, title consistency, page-marker consistency, no chapter navigation, layout diversity, balanced identity scale, adaptive subtitle placement, content-stage fill, single-folder output, and forbidden internal text.

Do not plan probable questions, backup answers, committee concerns, rehearsed defense logic, speaking scripts, or speaker notes. Planning information belongs in `output/ppt_plan.md` and `output/asset_manifest.md`, not on slides.

# Chapter Structure

Use a strong Chinese engineering thesis-defense structure:

- Cover.
- `目录`, not `答辩提纲`; avoid `汇报内容` unless a template or institution convention requires it.
- Part 01: `绪论` or `研究背景与意义`, covering background, significance, research status, problem/gap, research content, innovation overview, and chapter arrangement when useful.
- Part 02: `相关理论与技术` or `理论基础与关键技术`, only when the committee needs technical foundations.
- Part 03-05: core research works. In compact decks, carry transitions through `目录`, numbered headers, and local claims. In longer decks, give each major work a divider and a local rhythm: data/problem -> method/framework -> experiment/data -> quantitative/qualitative result -> ablation/robustness/analysis.
- Later part: `总结与展望`, including conclusions, contribution mapping, limitations, and future work.
- Optional final academic output part: `攻读学位期间科研成果` or `论文发表与科研成果`, when outputs appear in the dissertation.
- Closing: acknowledgements and `敬请各位老师批评指正`.

Keep 5-7 top-level parts at most in compact mode. When core work is the thesis center, allocate 3-5 slides per work and compress background/theory. When the thesis is method-heavy, split core work into `框架/方法`, `实验设置`, `结果对比`, and `消融分析`.

# Source Extraction

Prefer `.docx` because it preserves headings, captions, tables, formulas, and embedded original figures. If only PDF is available, crop selected figures carefully and record lower editability.

Extract:

- title, author, advisor, school, college, major, degree level, defense date
- abstract, keywords, table of contents, chapter hierarchy
- background, significance, research status, gap, problem definition
- research objectives, technical route, datasets, devices, metrics
- chapter-level methods, experiments, results, conclusions, innovations
- figures, tables, formulas, captions, and source page/chapter references
- publications, patents, software copyrights, awards, projects, datasets, or other outputs
- conclusion, limitations, and future work

Evidence selection priorities:

- Use thesis figures as evidence, not decoration.
- Prefer method/workflow diagrams, system architecture, experimental setups, qualitative comparisons, ablations, robustness checks, and quantitative summary tables.
- For CV, segmentation, detection, 3D reconstruction, point clouds, remote sensing, robotics, simulation, and industrial inspection, prioritize visual comparison panels with metrics nearby.
- Use formulas rarely: one central objective/loss/energy/governing equation only when it explains the contribution.

# Cover And Header System

Cover:

- Use school logo/name near the top or above the title; keep it aligned and balanced.
- Place the thesis title visually near the center and metadata below it.
- Center the title, author/presenter, major, advisor, college/school, and date at both textbox and paragraph level.
- Use one decisive cover family: wide school-color title band, full school-color cover, minimalist white cover with thin rules, or source-template cover. Do not mix cover grammars.

Content headers:

- Use one top-left chapter number integrated into the chapter title, such as `4 基于条件引导的条件生成式点云补全`. Do not pair a number badge with a title that repeats the same number.
- Keep the chapter title stable within each part; local subtopics belong in subtitle/claim text.
- Make the chapter title visually dominant and at least 24 pt. Wrap long titles cleanly rather than shrinking them into body-sized text.
- Use a short, source-specific subtitle/local claim when useful. It may sit in the header, below the separator as a claim row, or inside the body as a stronger callout depending on density.
- Put the logo in a consistent header position, usually top-right, large enough to be recognizable without overpowering content.
- Add a school-color separator under the header. Treat the region below it as the main evidence stage.

Good chapter/local claim examples:

- `基于体素扩散的大范围场景补全`
- `无提示气体分割模型设计`
- `多尺度几何与空间位置协同调制`
- `在遮挡场景下保持更完整的结构边界`

# Layout System

Use a white institutional base with rich structure. Evidence must occupy the main stage; bullets explain evidence instead of replacing it.

Allowed macro-layout families include:

- large original figure + side interpretation rail
- two-column method/result evidence
- three-card `问题 / 思路 / 方法` or `数据 / 模型 / 结果`
- 4-6 step arrow framework or vertical arrow bullets
- matrix for research status, technical route, ablation, or comparison
- full or split table with highlights
- figure + mini table/chart + concise conclusion
- framework map with dominant architecture and labeled stages
- qualitative comparison plate with aligned method labels
- layered evidence board with anchored annotations
- zoom-and-proof layout with magnified crop, metric tag, or mechanism inset
- swimlane pipeline for data/model/loss/output
- hub-and-spoke relation map
- split-stage method above/result below
- evidence storyboard of source-derived snapshots
- scoreboard or metric ladder paired with a compact table/chart
- before/after or baseline/ours confrontation
- timeline or research-evolution strip
- technical exploded view with real dependency connectors
- publications cards grouped by `论文 / 专利软著 / 项目数据 / 获奖`
- section divider with large part number, concise title, and restrained auxiliary motif

For a 20-slide deck, use at least 6 distinct macro-layout families. Do not let three consecutive content slides share `title + bullets + image`, `title + boxed cards`, or `left dark block + right rectangle`. The `left dark block + right rectangle` pattern is allowed at most 3 times in an 18-25 slide deck and never on adjacent slides.

Do not fix sparse slides with a repeated bottom conclusion bar. First enlarge the proof object, add another source-derived figure/table, split dense evidence across slides, convert data into a table/chart, add an editable mechanism diagram, use a callout rail, or add a neutral auxiliary schematic that clarifies structure without fabricating evidence.

# Evidence Stage Gates

Apply these gates while building every substantive method, result, ablation, dataset, experiment, and conclusion slide:

- The slide has one defense point, one dominant proof object, and a clear reading order.
- The area below the header separator is intentionally filled as one system, not scattered objects in the upper half.
- The dominant proof object usually occupies 45-70% of slide area when it is a key method diagram, architecture, visual comparison, reconstruction, segmentation, detection, or result figure.
- Visual comparison plates are large, aligned, equally cropped, and labeled; the student's method or key improvement is findable within 3 seconds.
- Related figures with the same role have equal scale and aligned baselines, or the slide explicitly marks one as main proof and the other as inset/detail.
- If architecture and submodules both matter, use an exploded view with connectors or split across slides rather than shrinking details into a corner.
- If the bottom third is empty, change the layout before moving on: enlarge/reposition evidence, add source-derived proof, add a metric strip, add a zoom crop, or switch to a fuller stage grammar.
- Formal pages such as cover, `目录`, dividers, outputs, and closing slides may use deliberate whitespace; substantive content slides should not look unfinished.

# Text, Tables, And Figures

Text:

- Use 2-4 concise bullets or short phrases; avoid thesis paragraphs and generic filler.
- Make result-slide subtitles claim-like and thesis-specific.
- Keep body text usually 20 pt or above; use 18 pt sparingly for dense secondary support and never below 16 pt for audience-facing text.
- Ensure every textbox wraps inside its visible frame, has padding, and does not overlap. Break long Chinese clauses into 2-3 lines instead of relying on one long line.

Tables:

- Preserve important rows, columns, metrics, datasets, baselines, and method names.
- Recreate legible tables as professional three-line/booktabs-style editable tables when practical.
- Use highlight color, bold text, or callout arrows for the student's method, best value, or key improvement.
- Convert to charts only when the chart communicates better and the table remains available nearby or on a following slide.
- Use table screenshots only when recreation risks transcription errors or original layout is itself important.

Figures:

- Insert clean figures directly on white when the original has a white/transparent background.
- Do not automatically add cards, shadows, gray panels, or colored rectangles behind every figure.
- Add a light container only for legibility, busy-figure separation, or intentional comparison layouts.
- Crop thesis page margins, redundant captions, and unrelated panels.
- Preserve aspect ratio. Never stretch detection, segmentation, reconstruction, or chart images.
- Add minimal labels/callouts only when they help the committee read the evidence.

# Publications And Outputs

If the dissertation contains publications or outputs, include them in the PPT as a dedicated slide or as part of the contribution section. Capture paper title, venue/journal, year, author position, status, DOI if available, patents, software copyrights, datasets, awards, and projects.

Keep this section factual and source-grounded. If outputs are many, group them by `论文 / 专利软著 / 项目数据 / 获奖`. Use SimSun / `宋体` for Chinese text and Times New Roman for English text in this section only.

# Build-Time Gates

Before saving the final PPTX, the current build must already satisfy:

1. Slide size is 16:9 and the slide count matches the planned compact or long-deck range.
2. Cover title and metadata are visually centered and paragraph-centered.
3. `目录` is present and chapter rhythm matches the planned divider policy.
4. `总结与展望` is substantive and separate from the closing slide.
5. Headers use meaningful numbered chapter titles, stable logo placement, readable subtitles/claims, separator rules, and consistent bottom-right page markers.
6. No persistent chapter navigation or progress-number rails appear.
7. All normal editable text uses Microsoft YaHei / `微软雅黑`; outputs use the required SimSun/Times New Roman exception.
8. Text wraps, stays inside boxes, has padding, and avoids overlap.
9. Audience-facing text is at least 16 pt; content-slide chapter titles are at least 24 pt.
10. Figures preserve aspect ratio and key evidence is large enough to read.
11. Tables/charts preserve important metrics and highlight the student's method or key result.
12. Each substantive slide has a claim, dominant proof object, and source-specific interpretation.
13. The below-header content stage is intentionally filled without clutter or empty lower thirds.
14. Compact decks show at least 6 macro-layout families and avoid adjacent repeated compositions.
15. Connectors/arrows encode real flow, dependency, comparison, or data direction and attach visually to the right objects.
16. Equal-role boxes, lanes, stages, metric items, and table cells share consistent geometry unless hierarchy intentionally differs.
17. No full-slide screenshots are used as a substitute for editable slides.
18. No fake, generated, redrawn, approximate, or pseudo-official logos/marks are used.
19. No internal workflow labels, Q&A forecasts, backup-answer pages, speaker scripts, or production notes appear on slides.
20. All run-created artifacts remain in `output/` or the selected delivery folder.

Make any corrections during construction before saving the final deck. Once the final PPTX is saved, deliver it without launching a separate review cycle unless the user explicitly requests one.
