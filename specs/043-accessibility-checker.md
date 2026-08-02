---
id: SPEC-043
title: Accessibility checker
status: draft
area: editing
depends_on: [SPEC-019, SPEC-024, SPEC-022]
---
## Why
Public-sector and enterprise customers are legally required to publish
accessible documents, and Word ships this check. It is also a quiet quality
tool: most of what it flags is bad structure, which hurts everyone.

## Acceptance criteria
- [ ] "Tools → Check accessibility" opens a results pane listing issues grouped by severity (Error, Warning, Tip)
- [ ] The checker detects: images without alt text, tables without a header row, tables with merged cells, heading levels skipped (e.g. H1 → H3), empty headings, links whose text is a bare URL or "click here", text with insufficient contrast against its background, and content distinguished only by colour
- [ ] It also detects documents with no heading structure at all and paragraphs that are visually headings (large bold text) but not styled as headings
- [ ] Selecting an issue scrolls to and highlights the offending content
- [ ] Each issue explains why it matters in one sentence and offers a fix action where one is mechanical (add alt text, mark header row, fix heading level)
- [ ] Applying a fix re-runs the check for that item and updates the list
- [ ] The pane shows a summary count and a clean state ("No accessibility issues found")
- [ ] The check runs on a 200-page document in under 3 seconds without blocking the UI
- [ ] The editor UI itself passes an automated accessibility scan (axe or equivalent) on the main editing screen with no serious or critical violations

## Out of scope
Automatic alt-text generation. Reading order controls for floating objects.
Remediating the PDF export. WCAG conformance reporting.
