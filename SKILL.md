---
name: tech-report-ppt-safe-layout
description: Use when generating or editing a technical presentation or PPT and layout reliability matters. Produces chart-led, image-text balanced slides and includes explicit rules to prevent text overlap, text clipping, and text exceeding card or shape boundaries.
---

# Tech Report PPT Safe Layout

Use this skill for paper reports, agent method presentations, benchmark reviews, and technical decks where the output is a real PPT rather than plain slide notes.

This is the default workflow for PPT generation and PPT editing tasks. When the user asks to create, refresh, restyle, or revise a PPT, you must use this skill and its bundled scripts unless the user explicitly asks for a different process.

## Goal

Produce slides that are:

1. Presentation-ready, not document-like
2. Chart-led when the paper contains measurable results
3. Safe against text clipping, overlap, and out-of-bounds wrapping
4. Easy to revise after generation

## Default Structure

For a 7-slide technical report:

1. Overview: what it is, method, problem, result
2. Trigger scenario: what problem forced the method
3. Core mechanism 1
4. Core mechanism 2
5. Core mechanism 3
6. Evaluation: tasks, baselines, metrics, headline results
7. Takeaways: technical essence, engineering implications

## Content Design Rules

- Use one visual anchor per slide: chart, flow, matrix, timeline, or comparison diagram.
- Keep each slide to one claim, one visual, and one short support block.
- Prefer charts and structured diagrams over long bullet lists.
- Use only the numbers that support the slide's main claim.
- Split dense content across slides rather than shrinking text aggressively.

## Layout Safety Rules

These rules are mandatory when generating the actual PPT:

- Reserve a safe content area. On 16:9 slides, keep at least 0.45in outer margins and 0.18in gutters between major regions.
- Never place long text in a small metric card. If the descriptor needs more than 2 lines, widen the card or shorten the copy.
- Do not rely on "it will probably wrap". Size boxes for wrapped text explicitly.
- Prefer putting text directly inside the target shape when possible. Avoid stacking a separate textbox on a small colored card unless the text is very short.
- Do not fill a text region to its geometric limit. Leave at least 15-20% vertical slack in body text containers.
- Titles and subtitles need extra headroom because Chinese + English mixed text wraps unpredictably.
- If a text block contains both Chinese and English terms, assume it is wider and taller than a same-length pure English block.

## Text Budget Rules

Use these as defaults unless the user requests otherwise:

- Title: <= 22 Chinese characters or <= 55 mixed characters
- Subtitle: <= 32 Chinese characters or <= 80 mixed characters
- Main bullet block: 3-4 bullets, each <= 26 Chinese characters or <= 70 mixed characters
- Small card label: <= 10 Chinese characters or <= 24 mixed characters
- Small card description: <= 2 lines
- Takeaway strip: one sentence only

If a line exceeds the budget, rewrite it instead of shrinking the font first.

## Visual Composition Rules

- Every slide should have a dominant visual area and a separate explanation area.
- Charts should occupy enough area to be readable; do not squeeze charts to make room for too much prose.
- Use short annotation blocks beside charts instead of full-width paragraphs below them.
- Keep decorative elements minimal. The slide should read as a technical report, not a poster.

## PPT Generation Workflow

1. Build the slide outline.
2. Assign a visual type to each slide before writing detailed copy.
3. Draft the copy to fit the text budgets above.
4. Generate the PPT. Use `scripts/build_safe_ppt.py` as the default skeleton when starting from scratch.
5. Run `scripts/check_ppt_layout.py <pptx>` on the generated file.
6. Re-open the generated PPT and inspect every text-bearing shape.
7. Fix risky boxes before finishing.

## Mandatory Application Rule

For any PPT creation or PPT modification request:

- Start from `scripts/build_safe_ppt.py` when creating a new deck or rebuilding a deck structure.
- Run `scripts/check_ppt_layout.py` on the output before final delivery.
- If the checker reports findings, fix them before returning the PPT.
- Do not skip the checker just because the deck "looks fine" from the code.
- Mention in the final response that the PPT was checked with the layout checker.

## Bundled Resources

- `scripts/build_safe_ppt.py`: creates a 7-slide technical report skeleton with safe margins, card spacing, and chart-friendly regions.
- `scripts/check_ppt_layout.py`: scans a `.pptx` and flags likely text clipping, compact-card overflow, banner wrapping, and textbox/card misalignment.
- `references/layout-rules.md`: short reference sheet for safe PPT geometry and repair order.

## Required Post-Generation Checks

After generating the PPT, explicitly check for these failure modes:

- Title box too narrow for the actual mixed-language title
- Subtitle height too short after wrapping
- Bullet box height consumed by too many lines
- Textbox sitting on top of a card but extending beyond the card's visual area
- Chart notes or metric descriptions spilling into adjacent elements
- Bottom annotations colliding with footer or page chrome

If any of these appear likely, regenerate or patch the PPT before final delivery.

Example:

```powershell
python C:\Users\27116\.codex\skills\tech-report-ppt-safe-layout\scripts\build_safe_ppt.py D:\out\deck.pptx --title "Meta-Harness"
python C:\Users\27116\.codex\skills\tech-report-ppt-safe-layout\scripts\check_ppt_layout.py D:\out\deck.pptx
```

## Heuristics For Flagging Risky Text Boxes

Treat a text box as risky when any of the following is true:

- It contains more than 4 wrapped lines in a standard body area
- It contains more than 2 lines inside a compact badge/card
- The text height visually approaches the box height
- The box width is only slightly larger than the longest mixed Chinese-English phrase
- The slide depends on precise overlay alignment between a shape and an independent textbox

When a box is risky, fix it in this order:

1. Rewrite shorter
2. Widen the box
3. Increase box height
4. Reduce font size slightly
5. Change the slide layout

## Lessons From Prior Failure

The main failure patterns to avoid are:

- Hard-coding box sizes too early
- Writing slide copy as if it were a document paragraph
- Using too many separate textboxes on top of decorative cards
- Leaving no slack below charts for labels and notes
- Skipping a final layout pass after generation

The safe default is: fewer words, larger boxes, more whitespace, and a mandatory validation pass.
