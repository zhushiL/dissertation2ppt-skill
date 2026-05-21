---
name: dissertation2ppt
description: Create an editable Chinese dissertation or thesis defense PowerPoint from degree thesis materials, preferably a Word .docx with original chapter figures. Use when Codex needs to act as a top Chinese engineering defense expert and produce a 16:9, reference-quality 18-25 slide PPTX for bachelor, master, or doctoral theses, with truly centered cover text, a formal `目录` page, consistent bottom-right page markers, clean chapter hierarchy, no persistent chapter navigation, no visible internal reasoning labels, layout-diverse evidence-first pages, large readable thesis figures, full useful tables, a single organized output folder, template-aware design when the user supplies a PPT template, output/ppt_plan.md planning, extracted thesis images, and strict build-time QA for editability, density, text wrapping, overflow, title consistency, and image placement.
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

Keep generated artifacts together. Any build scripts, extracted text, intermediate JSON, contact sheets, previews, temporary work files, validation logs, and repair notes created for the run should live under `output/` or a user-named delivery folder, not beside the source thesis or scattered in the working directory. If a helper script is needed, write it as `output/build_*.py` or `output/scripts/*.py`. Leave the user's original source files in place.

Hard requirements:

- 16:9 widescreen.
- 18-25 slides by default.
- Pure white slide canvas by default when no template is supplied. Use school-color lines, top headers, section dividers, badges, cards, shapes, and subtle diagrams on top of white; do not use full-slide photo, gradient, generated, or decorative colored backgrounds unless the user supplies a template whose design language clearly requires it.
- Editable PPT objects for text, diagrams, tables, and charts. Do not flatten slides into screenshots.
- A clean school identity: logo on content slides, accents whose hue stays close to the school emblem or official school color, and consistent fonts. The identity mark should be visually present and balanced with the header, not reduced to an unnoticeable corner token.
- No large empty areas, no text overflow.
- Readable scale for defense rooms: on content slides, the top-left chapter title should be at least 24 pt, and normal audience-facing text should be at least 16 pt. Page markers, tiny logo text, and unavoidable labels embedded inside original thesis figures may be smaller, but they must not carry the slide's main argument.
- Cover text boxes and the paragraphs inside those boxes must both be horizontally centered; centering only the box is not enough.
- Do not use persistent chapter navigation strips or bottom-left progress navigation. A page marker is fine; a repeated `01 02 03 04` navigation rail is not. Put page markers consistently in the bottom-right across cover, `目录`, section dividers, content slides, publication/output slides, and closing slides unless a user-supplied official template has a clearly established different location.
- The final PPTX must be immediately presentable. Do not add probable questions, backup answers, defense-logic pages, rehearsed scripts, hidden speaker notes, or visible internal reasoning labels unless the user explicitly requests them.

# Template Priority

If the user supplies a template PPT or explicitly asks to imitate a reference template, inspect the template first and treat its visual system as a design constraint:

- Follow the template's cover grammar, page marker position, section-divider style, typography, color system, logo treatment, and recurring layout logic when they are coherent and professional.
- The template may override the default white-canvas preference, header placement, or section-divider grammar. Do not force this skill's default style onto a good user-provided template.
- Keep non-negotiable defense requirements: source-grounded content, editable text/tables/diagrams, readable figures, no text overflow, no internal/audit labels, no fake logos or fake evidence, and no persistent chapter-navigation clutter unless the template explicitly uses it and the user wants that convention.
- Record template use and any deviations in `output/ppt_plan.md` and `output/asset_manifest.md`.

# Reference-Quality Design Bar

Strong Chinese degree-defense samples share these traits. Treat them as build requirements, not optional polish:

- Formal cover: school logo/name, centered thesis title, centered author/advisor/major/date metadata, and one restrained school-color title band or rule.
- Stable identity: content slides use a consistent logo position, prominent chapter title, clear subtitle or local claim, page marker, school-color accent, and thin separator rules. The header should look designed, not like plain text placed on a blank slide.
- Strong content-stage use: after the header separator, the region below it is the main presentation stage. Substantive pages should make that region feel intentionally filled with evidence, interpretation, and structure, rather than leaving the proof object small in the upper half.
- Content-stage fullness in strong samples usually comes from layout systems, not decoration: large proof objects, aligned side rails, lower-row metric strips, multi-panel evidence boards, boxed module maps, and connected flows that occupy the slide as one system. Avoid treating content fill as an after-the-fact patch with a generic bottom sentence.
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

Keep this skill self-contained. Do not switch to artifact-tool, Google Slides import workflows, business-deck profiles, or a separate presentation skill unless the user explicitly asks for those tools. Borrow only the general design discipline: claim-led slide planning, locked visual grammar, contact-sheet review, structured visual precision, and iteration before delivery.

# Editorial Quality Workflow

Before making slides, turn the thesis into a defense-quality narrative, not a chapter-by-chapter summary.

1. Extract the thesis spine: central problem, technical route, core contributions, strongest evidence objects, and final conclusion.
2. Define a slide-level claim spine in `output/ppt_plan.md`: every non-cover, non-toc, non-divider, non-closing slide should have one defense claim, one dominant proof object, and one short source-grounded interpretation.
3. Lock the design system in `output/ppt_plan.md`: fonts, school-color palette, header/page-marker grammar, chart grammar, table grammar, connector grammar, callout/container grammar, figure crop rules, allowed layout families, and banned motifs.
4. Make a pre-layout decision for every substantive slide before placing objects: identify whether the proof objects are a single dominant evidence object, a set of equal-weight related figures, a main proof with supporting insets, or evidence that should be split across two slides. Similar method diagrams, architecture diagrams, result panels, or ablation visuals should not appear accidentally unequal; if one is much larger, the plan must make that hierarchy meaningful through labels, callouts, or an explicit main-proof/supporting-proof relationship.
5. Plan the contact sheet before building: the thumbnail view should show thesis evidence, varied composition, and stable school identity before any body text is read. The plan should already indicate how each content slide uses the page below the header as one coherent stage, not a scattered set of objects.
6. Build the editable PPTX, then audit the weakest slides and iterate before final verification when the first export looks merely acceptable.

Slide claims should be specific enough that they would not still work after replacing the thesis topic with another project. Topic labels such as `实验结果分析` are acceptable as subtitles or section context, but result-slide titles or subtitles should state what the experiment proves.

# Chapter Design Rule

Use a defense-chapter structure similar to strong Chinese engineering thesis decks:

- Cover.
- `目录`, never `答辩提纲`. For Chinese degree-defense decks, the second slide title should normally be `目录`; avoid `汇报内容` unless the user supplies a template or institution convention that uses that wording.
- Part 01: `绪论` or `研究背景与意义`, including background, significance, research status, problem/gap, research content, innovation overview, and thesis chapter arrangement when useful.
- Part 02: `相关理论与技术` or `理论基础与关键技术`, only when the thesis depends on technical foundations the committee must understand.
- Part 03-04 or Part 03-05: core research works. Each major work should have its own section divider and then follow a local rhythm: data/problem -> method/framework -> quantitative experiment -> qualitative visualization -> ablation/robustness/analysis.
- Part after core works: `总结与展望`, combining main conclusions, contributions, limitations, and future work.
- Optional final academic output part: `攻读学位期间科研成果` or `论文发表与科研成果`, when publications, patents, software copyrights, datasets, awards, or projects appear in the dissertation.
- Closing: acknowledgements and `敬请各位老师批评指正`.

For an 18-25 slide deck, avoid too many tiny sections. Use 5-7 top-level parts at most. When core work is the thesis center, give each work 3-5 slides and compress theory/background. When the thesis is method-heavy, split each work into `框架/方法`, `实验设置`, `结果对比`, and `消融分析`.

Treat `总结与展望` as a substantive section, not as a decorative closing gesture. It should normally have its own slide or slides with conclusions, contribution mapping, limitations, and future work. The final `敬请各位老师批评指正` closing slide should remain separate from the summary section. If the slide count is tight, compress background, theory, or repeated experiment pages before merging summary content into the closing page.

Every standalone top-level part in the planned outline should have its own section divider or intentionally designed guide page. Do not let some numbered parts have guide pages while another numbered part begins abruptly on a normal content page unless the plan explicitly merges that part into a neighboring section for pacing reasons.

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

Asset authenticity:

- Use verified school logos, user-provided logos, or logos extracted from the thesis/source files. Do not redraw, approximate, stylize, or generate official logos.
- Record logo, screenshot, publication, dataset, and external visual provenance in `output/asset_manifest.md`.
- If a verified identity asset cannot be obtained cleanly, rely on typography, school-color accents, and source-grounded content instead of inventing a pseudo-official mark.

# Required Plan

Before building the PPTX, write `output/ppt_plan.md` with:

- Thesis metadata and defense assumptions.
- Content diagnosis: problem, route, core works, evidence strength, and available thesis proof objects.
- Branding plan: logo source, school color and fonts.
- Numbered outline: the exact part numbers used in slide headers.
- Slide-budget plan: allocate pages to background, theory, each core work, summary, outputs if any, and the independent closing slide. State what is compressed if the 18-25 page range is tight; do not solve a tight count by fusing `总结与展望` with `敬请批评指正`.
- Claim spine: one-line overall arc plus slide-level claim, proof object, and support note for each substantive slide.
- Design-system lock: background, type hierarchy, school-color usage, chart/table/diagram/container grammar, bottom-right page marker grammar or template-derived marker grammar, source/asset treatment, allowed layout families, and banned motifs.
- Reference-quality design plan: cover grammar, section divider grammar, header/subtitle grammar, page marker grammar, explicit no-navigation rule, content-stage fill strategy below the header separator, readable type scale, and at least 6 macro-layout families for a 20-slide deck.
- Slide table for the planned deck: number, Chinese title, display purpose, source chapter, primary figure/table/proof object, evidence hierarchy, and layout pattern.
- Publication/output plan: list thesis publications or outputs and where they appear; if none are found, state that.
- Table plan: list important original tables and whether to recreate fully, split across slides, convert to chart, or keep as cropped image.
- Figure plan: state which figures are inserted directly, which need a light container, crop, label, or callout, which key figures must be shown large instead of as thumbnails, and how related or visually similar figures will be treated: equal-weight grid, main-and-inset, before/after pair, or split across slides.
- Auxiliary visual plan: identify any topic-aligned generated or schematic patterns/icons used only to improve layout structure, and state how they stay separate from source-grounded evidence.
- Contact-sheet plan and layout density check: identify slides at risk of being empty, bottom-blank, text-heavy, table-heavy, figure-dense, repetitive, under-scaled, too generic, or visually unbalanced because related figures have accidental size hierarchy.
- Build-time quality plan: cover paragraph centering, text wrapping, text overflow, minimum readable font scale, figure readability, white canvas or template-justified canvas, editability, title consistency, bottom-right page-marker consistency, no chapter navigation, contact-sheet rhythm, layout diversity, balanced identity scale, section-divider coverage for every numbered part, adaptive subtitle placement, content-stage fill below the header separator, single-folder output organization, and forbidden internal/audit text.

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
- The content-slide chapter title in the top-left header should be visually dominant and normally at least 24 pt. If the title is long, wrap it cleanly into two lines or tighten the wording; do not reduce it into small body-sized text.
- A short, specific subtitle or local claim when useful, such as `局部条件扩散补全框架`; set it smaller than the chapter title but still visibly intentional. It should have enough contrast, spacing, and width to read as the slide's claim or local argument, not as a faint footnote.
- Place the subtitle adaptively according to slide density. If the header is visually crowded or the page needs more breathing room, put the subtitle below the chapter-title separator as a clean claim row. If the content area is sparse, the subtitle may sit within the header system or become a stronger body callout. Keep the choice consistent within a section unless a slide's evidence type clearly calls for a variant.
- The school logo in a consistent header position, usually top-right, sized and positioned to balance the header. Avoid tiny corner logos that disappear in thumbnail view; use a restrained but recognizable logo, usually paired with enough top/right margin and not crowded by the title.
- A separator line under the header, colored to match the school's institutional color.
- A small page marker should appear in the bottom-right by default. Do not put page markers in the top-right on some slides and bottom-right on others. Do not add chapter navigation or progress-number strips.

Header design should have a visible hierarchy. Use a school-color title rule, a slim title band, an accent number block, or a similar restrained institutional device so the chapter title does not look like unstyled body text. The subtitle should either make a source-specific claim or name the local proof object; if there is no useful subtitle, improve the title or layout rather than leaving a weak gray line. Avoid making some pages feel cramped by header text while others feel empty because the subtitle is too faint or too far from the evidence.

Treat the separator line as the boundary between identity/header and the main slide stage. On normal content slides, design the area below that line first: place the dominant proof object, table, diagram, comparison, or process system so it occupies the available field confidently, then add interpretation rails, callouts, or secondary evidence to complete the lower half.

对于核心内容幻灯片，章节标题在每个章节中保持一致，而不同部分的子章节标题应同时阐明工作内容和核心论点，例如：

- `基于体素扩散的大范围场景补全`
- `无提示气体分割模型设计`
- `多尺度几何与空间位置协同调制`

Header mistakes to avoid:

- Do not combine a number badge `2` with a title that also starts with `2`; duplicate numbers make the hierarchy look broken.
- Do not use bottom-left or side chapter navigation such as `01 02 03 04`, even if the current chapter is highlighted.
- Do not color same-level chapter numbers inconsistently.
- Do not mix page-marker positions across slide types. A deck with top-right page numbers on content slides and bottom-right page numbers on dividers feels accidental.

Do not put internal workflow labels on slides, including `资料来源`, `素材来源`, `答辩提纲`, `归纳出的答辩主线`, `答辩主线`, `答辩叙事`, `注意事项`, `生成说明`, `AI生成`, `来自论文原文`, `截图自`, `output/ppt_plan.md`, or labels that describe how the deck was produced. The committee should see only defense content.

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
- Layered evidence board: one dominant figure/table in the center, with 3-5 short anchored annotations around it.
- Zoom-and-proof layout: large full evidence object plus one magnified crop, metric tag, or mechanism inset that explains the key detail.
- Swimlane or lane-based pipeline: rows for data/model/loss/output, useful when a method has parallel branches or training/inference stages.
- Hub-and-spoke or radial relation map: central method/system/dataset with surrounding modules, constraints, or evaluation dimensions.
- Split-stage layout: upper half for method or experiment setup, lower half for result proof, connected by arrows or numbered anchors.
- Evidence storyboard: 3-5 source-derived snapshots or stages arranged left-to-right with concise claims, useful for procedures, reconstruction, segmentation, and experiments.
- Scoreboard or metric ladder: large key numbers, deltas, ranks, or ablation steps paired with a compact table/chart.
- Before/after or baseline/ours confrontation: two large balanced panels, one highlighted result zone, and a short interpretation rail.
- Timeline or research-evolution strip: use for literature status, algorithm evolution, dataset construction, experiment protocol, or thesis chapter progression.
- Technical exploded view: split a framework into modules arranged around the original architecture figure, with connector lines that show real dependencies.
- Generated auxiliary motif: when the thesis lacks enough visual structure for a divider or conceptual bridge, generate or draw a neutral topic-aligned line pattern, icon set, or abstract schematic; record it in the asset manifest and keep it clearly separate from logos, data, tables, experiments, and other evidence.

For a 20-slide deck, use at least 6 distinct macro-layout families. Do not let 3 consecutive content slides share the same `title + bullets + image`, `title + boxed cards`, or `left dark block + right text rectangle` composition. The contact sheet should look authored before the text is read.

Anti-repetition rules:

- The `left dark block + right rectangular text area` layout is allowed, but it must not become the deck default. In an 18-25 slide deck, use it on no more than 3 content slides and never on adjacent slides.
- Do not solve every sparse page with the same bottom conclusion bar. Prefer a larger thesis figure, a second evidence object, an arrow framework, a table slice, or a structured comparison.
- Section dividers may be visually related, but content slides should rotate layout families based on evidence type.

Avoid relying on a bottom conclusion bar as the default fix for empty slides. Use a bottom bar only when it carries a real takeaway and does not become repetitive. When a slide looks empty, first add evidence, structure, a diagram, a table, or a source-derived short description.

Main-stage density:

- On content slides, the area below the header separator is the primary content stage. The dominant proof object and its interpretation should fill this stage as the default, not sit as a small item near the top.
- Main content should occupy most of the below-header canvas. Aim for a full, authored stage with clear reading order, not sparse floating objects.
- Filling the content stage means composing a complete evidence field, not merely increasing object count. Strong content pages often use one of these stage grammars: a large figure/table plus an aligned interpretation rail, a full-width process/framework with bottom labels, a two-row evidence board with method above and result below, equal-weight comparison panels with consistent captions, a table/chart paired with metric cards, or a central architecture surrounded by short anchored annotations.
- Decide the spatial grammar before placing the first object. The slide should have a top/bottom or left/right system that reaches the lower part of the stage deliberately; do not place a few objects in the upper half and then leave the bottom to a weak note, page marker, or decorative bar.
- If the bottom third is blank, resize/reposition evidence, add another source-derived figure/table, switch to a fuller layout, or use an auxiliary schematic/pattern that clarifies the evidence without fabricating results.
- Treat normal content pages more strictly than cover, `目录`, section dividers, publication/output pages, and closing pages. Formal pages may keep deliberate whitespace; substantive method, result, ablation, dataset, experiment, and conclusion pages should use the lower half deliberately.
- A slide with only bullets is usually unacceptable unless it is a formal outline, limitations, or closing slide.
- Do not leave a content page with small evidence objects floating in the upper half. If a table, sampling sequence, ablation, architecture, or visual comparison is the main proof, make it dominate the page, pair it with another source-derived proof object, add source-grounded explanatory text, or convert it into a fuller matrix/comparison plate.
- A visually full slide is not the same as clutter: keep margins, but the main proof and interpretation should fill the available canvas with a clear reading order.
- When source material is sparse, enrich the page with one of: a larger crop of the proof object, a second source-derived figure/table, a compact source-grounded explanation, an editable process diagram, a callout rail, a comparison matrix, a zoomed detail, a metric ladder, or a neutral auxiliary pattern/schematic that supports reading. Generated or schematic elements may support structure, but must never fabricate evidence, logos, experimental results, or data.
- Balance density across adjacent slides. Avoid a deck where one evidence page is packed and the next content page looks unfinished; adjust subtitle position, figure scale, side rails, and explanatory text so the contact sheet has an even professional rhythm.
- When a result or method slide still has unused bottom space after placing the main proof, do not merely add a decorative conclusion bar. Prefer one of: a compact metric strip, a zoomed crop from the key figure, a small source-derived table slice, an editable mechanism diagram, a before/after comparison, or a short evidence-linked interpretation rail.
- Before finalizing a content page, ask whether a committee member can infer the slide's visual hierarchy from the thumbnail alone. If the largest element is not the main proof, if two comparable figures are inexplicably different sizes, or if the proof objects occupy disconnected islands, redesign the layout before writing more body text.

Figure scale rules:

- Key method diagrams, visual comparison plates, reconstruction/segmentation/detection examples, and architecture figures should usually occupy 45-70% of the slide area.
- For visual comparison slides, the comparison image should be the main object, not a small thumbnail in a corner. Use large aligned panels with method labels and nearby metrics/callouts.
- Avoid placing evidence figures below roughly one quarter of slide width unless they are intentionally secondary thumbnails in a comparison grid.
- When two or more figures are visually similar or conceptually parallel, either give them equal scale and aligned baselines, or explicitly turn one into the main figure and the other into a labeled inset/detail. Do not let similar figures appear with accidental scale differences, mismatched containers, or unrelated positions.
- If a method slide needs to show both a whole architecture and submodule details, prefer an exploded view: the architecture remains the dominant object, submodules are cropped/linked as insets with clear connectors, and all empty regions become labels, module notes, or dependency arrows. If the submodule is too important or too detailed, split it onto a following slide instead of shrinking it into a corner.
- Treat repeated audit signals about small content images as repair candidates, except for logos, section dividers, closing slides, or intentionally secondary thumbnails. If the slide's proof depends on a figure, enlarge or split the proof instead of accepting a tiny image.

Structured visual precision:

- Before authoring any chart, framework, pipeline, table, matrix, or connector diagram, decide what it must prove and what the primary reading order is.
- Connectors and arrows must encode real direction, dependency, comparison, or data flow. Do not use arrows as decoration.
- Connector endpoints should attach visually to the intended source and target, avoid unrelated objects, and preserve the intended direction after export.
- Equal-role boxes, lanes, stages, metric items, or table cells should share consistent size, alignment, padding, border logic, and text treatment unless the hierarchy intentionally differs.
- Text inside filled boxes, dark panels, badges, callouts, and metric rails needs visible top/bottom breathing room; if it looks pinned to an edge, enlarge the box or shorten the copy.
- Labels must sit close enough to the figure panel, mark, row, series, connector, or module they describe that the committee never has to guess the attachment.
- If a complex chart or diagram needs many exceptions to stay readable, simplify the visual rather than patching around poor geometry.

# Text Rules

- One slide, one defense point.
- Use 2-4 concise bullets or short phrases; avoid thesis paragraphs.
- Write brief, logical descriptions derived from the thesis source, not generic filler.
- For result slides, the subtitle should lead with the argument.
- Prefer claim-like titles or subtitles on evidence slides: `在遮挡场景下保持更完整的结构边界` is stronger than `定性结果展示`.
- Remove body copy that only fills space. Use source-derived evidence, labels, or a clearer diagram instead.
- Do not create speaker notes, hidden scripts, Q&A pages, `答辩逻辑` pages, `归纳出的答辩主线` pages, or visible reminders about what the presenter should notice unless explicitly requested.
- Content-slide chapter titles should be at least 24 pt. Normal audience-facing text, including body copy, callouts, cards, labels, and readable table text, should be at least 16 pt; if a table or caption cannot meet this, split, crop, chart, or enlarge it instead of shrinking the main argument. Body text should still usually be 20 pt or above, with 18 pt used sparingly for dense but secondary support. Page markers, tiny logo text, and labels already embedded inside original figures are the main exceptions.
- Never allow text to exceed its background box. If text overflows, shorten it, enlarge the box, reduce hierarchy safely, or split the content across slides.
- Text must wrap inside the visible content area. A long sentence placed in a wide box but rendered as one unbroken line that crosses its frame is a defect, even if the text box itself has enough nominal width. Set wrapping behavior, use manual line breaks for long Chinese clauses, and inspect rendered output instead of trusting geometry alone.
- Do not use a single long line for a dense claim, agenda item, table note, metric explanation, or publication status. Break it into 2-3 short lines or convert it into a structured card with explicit line breaks and enough top/bottom padding.

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

Use generated visuals only for auxiliary layout support, such as simple process icons, neutral line patterns, section-divider motifs, abstract schematics, or topic-aligned non-evidence textures. They may make the deck feel more designed, but they must never replace source evidence. Do not generate logos, experimental evidence, result images, tables, fake comparison visuals, dataset screenshots, or partner/product marks.

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
2. White slide canvas unless a user-provided template justifies another coherent style; school-color bands, headers, dividers, and local module anchors are foreground structure, not decorative full-slide backgrounds or persistent navigation.
3. Centered cover title and centered metadata, with paragraph alignment set to center inside each textbox.
4. Numbered headers, consistent logo position, subtitle placement, top-left chapter titles at least 24 pt, and bottom-right page marker by default; no persistent chapter navigation.
5. Meaningful section labels; never use meaningless labels such as `研究成果一` without the work topic. Every standalone numbered part has a matching divider/guide page unless the plan explicitly merges it with another part for pacing.
6. Publications/outputs included when present in the thesis.
7. Figures preserve aspect ratio, are not stretched, and are not automatically placed on gray cards or decorative panels.
8. Text stays inside boxes, wraps inside the intended frame, has visible padding, has correct alignment, and does not overlap other objects. Normal audience-facing text should not drop below 16 pt; if it must, redesign the layout or split the slide.
9. The deck remains editable and is not a sequence of full-slide screenshots.
10. `总结与展望` is a real content section and should not be collapsed into the final thanks/closing slide. The closing slide includes `敬请各位老师批评指正` or an equivalent defense closing as a separate final page unless the user explicitly asks for a one-page ending.
11. No visible internal/audit terms, Q&A forecasts, backup-answer pages, speaker scripts, `归纳出的答辩主线`, or defense-logic labels.
12. The deck has a consistent cover/header/page-marker system, except for intentional cover, divider, appendix, and closing variants.
13. Result slides pair claims with proof: a table/chart alone is weak unless the thesis result is purely numeric; a qualitative image alone is weak unless labels and metrics nearby explain it.
14. Thumbnail/contact-sheet review shows at least 6 macro-layout families in compact mode, with no adjacent repeated `left dark block + right rectangle` layouts.
15. No substantive content slide leaves the lower third obviously empty or lets the main proof sit as small objects in the top half. Fill the below-header content stage through an intentional stage grammar: enlarged source-derived proof, equalized comparison panels, structured arrows, tables, compact explanations, auxiliary schematics, generated non-evidence motifs, or a fuller layout. Cover, `目录`, section dividers, publication/output pages, and closing pages may use more formal whitespace.
16. Key evidence figures are large enough to read in slideshow mode, especially visual comparison images. If a key proof would become too small in a dense grid, split the proof across slides or enlarge the most important panel. Related figures with the same role should have comparable scale and alignment; if they do not, the slide must communicate why one is primary and the other is secondary.
17. Slide titles, subtitles, and callouts are source-specific; a title that could fit any thesis should be sharpened or moved to a divider.
18. Charts, pipelines, matrices, arrows, and framework diagrams preserve correct geometry, attachment, reading order, and label association after export.
19. Repeated visual systems such as step sequences, KPI/metric rails, chapter dividers, and comparison plates render every item with the full grammar: label, value or stage name, context, and adequate contrast.
20. Official identity assets are verified or user-provided; no generated, redrawn, approximate, or pseudo-official school logos, partner marks, dataset logos, or product UI are used.
21. Official or user-provided logos used in headers are large enough to support the institutional identity without overpowering the content. If the identity mark is too small to recognize in a contact sheet, revise the header system.
22. Subtitle placement is density-aware: it may sit in the header, below the separator line, or as a body claim/callout, but it must improve rhythm and readability rather than make the page cramped or sparse.
23. After the first complete export, inspect the deck as a contact sheet and revise the weakest slides before calling the deck final unless final verification is intentionally report-only. Empty-looking content pages, underfilled below-header stages, unwrapped long lines, small chapter titles, body text under 16 pt, abrupt missing section dividers, inconsistent page markers, weak subtitles, uneven subtitle placement, and undersized header logos are repair candidates.
24. All run-created artifacts should remain inside `output/` or the selected delivery folder. Do not leave `build_*.py`, extracted text dumps, validation JSON, preview images, or repair scratch files in the parent working directory unless the user explicitly asks for that layout.

# QA Scorecard

When creating or updating `output/qa_report.md`, include a compact scorecard rather than vague praise. Use 0-5 ratings and name remaining weak spots:

- `defense story`: the deck has a coherent thesis arc and each core work advances it.
- `source specificity`: claims, titles, figures, metrics, and outputs are grounded in the dissertation.
- `evidence strength`: each substantive slide has a dominant proof object and the proof supports the claim.
- `layout rhythm`: contact-sheet view shows varied macro-layouts without template-like repetition.
- `visual precision`: charts, tables, connectors, labels, and boxed systems align and read correctly.
- `readability`: title/body/table/figure sizes survive slideshow viewing, with no overflow or cramped boxes.
- `institutional coherence`: cover, headers, logo, page marker, school color, and fonts feel consistent.
- `header strength`: content-slide title, subtitle or local claim, separator, logo, and page marker form a balanced institutional header rather than plain text.
- `density`: substantive slides use the canvas effectively without empty lower halves, tiny main proof objects, or obvious density swings between adjacent content slides.
- `restraint`: no filler boxes, decorative badges, fake assets, repeated bottom bars, or overdesigned backgrounds.

Treat any dimension below 4 as a repair candidate in a normal build. If final verification is already in report-only mode, record the problem clearly and wait for a fresh repair instruction.

# Final Verification Is Report-Only

After the final PPTX is created, run an audit pass. For intentional long decks, pass the planned range to the script instead of accepting the default 18-25 warning:

```bash
scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx
# or:
scripts/validate_dissertation_ppt.py --min-slides 30 --max-slides 66 output/dissertation_defense.pptx
```

Then reopen the `.pptx` and inspect rendered slide previews when rendering tools are available.

If the script or visual inspection finds problems, write `output/qa_report.md` and report the issues to the user. Do not modify the already-created PPTX during the final verification step. A repair pass requires a new explicit instruction from the user or a fresh build run.
