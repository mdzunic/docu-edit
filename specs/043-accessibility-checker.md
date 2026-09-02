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
- [ ] "Tools → Check accessibility" opens a results pane listing issues grouped by severity (Error, Warning, Tip), each issue assigned the severity given in "Severity" below
- [ ] The checker detects: images without alt text, tables without a header row, tables with merged cells, heading levels skipped (e.g. H1 → H3), empty headings, links whose text is a bare URL or "click here", text below the contrast threshold, and content distinguished only by colour
- [ ] Contrast is judged by the WCAG 2.1 AA thresholds against a resolved background, both fixed in "Contrast" below — never by eye or by an implementation's own constant
- [ ] It also detects documents with no heading structure at all and paragraphs that are visually headings but not styled as headings, "visually a heading" being the test in "Looks like a heading" below
- [ ] "Distinguished only by colour" is the test in "Colour alone" below
- [ ] Selecting an issue scrolls to and highlights the offending content
- [ ] Each issue explains why it matters in one sentence and offers a fix action where one is mechanical (add alt text, mark header row, fix heading level)
- [ ] Applying a fix re-runs the check for that item and updates the list
- [ ] The pane shows a summary count and a clean state ("No accessibility issues found")
- [ ] The check runs on a 200-page document in under 3 seconds without blocking the UI — chunked through `scheduleIdleWork`, as spell check and find already are
- [ ] The editor UI itself passes an automated accessibility scan (axe or equivalent) on the main editing screen with no serious or critical violations

## Out of scope
Automatic alt-text generation. Reading order controls for floating objects.
Remediating the PDF export. WCAG conformance reporting. Text inside images —
undetectable without OCR, and the checker never guesses at pixels.

## Reference

### Severity
The grouping is not cosmetic — it is what a user triages by, so each detection
has one assigned level. The principle, and the assignments it produces:

- **Error** — a screen-reader user cannot reach the content, or is given the
  wrong content.
- **Warning** — the content is reachable but likely to mislead.
- **Tip** — correct as it stands; better structure is available.

| Detection | Severity |
| --- | --- |
| Image without alt text | Error |
| Table without a header row | Error |
| Empty heading | Error |
| Text below the contrast threshold | Error |
| Heading level skipped (H1 → H3) | Warning |
| Link text is a bare URL or "click here" | Warning |
| Meaning carried by colour alone | Warning |
| Merged cells in a table | Warning |
| No heading structure at all | Tip |
| Paragraph that looks like a heading but is not one | Tip |

A new detection added later inherits its level from the principle above, not
from whichever row it superficially resembles.

### Contrast
**Thresholds — WCAG 2.1 AA.** Text is flagged when its contrast ratio against
its resolved background is below:

- **4.5:1** for normal text;
- **3:1** for large text, where large means an effective `fontSizePt` of **18 pt
  or more, or 14 pt or more when the run is bold**.

AA rather than AAA (7:1) because AA is the level the public-sector procurement
rules that motivate this spec actually cite, and a 7:1 default would flag
ordinary dark-grey-on-white body text as an Error.

**Ratio.** The WCAG relative-luminance formula: linearize each sRGB channel,
`L = 0.2126R + 0.7152G + 0.0722B`, ratio `(Lmax + 0.05) / (Lmin + 0.05)`. Round
the ratio to two decimals before comparing, so a value that computes as
4.4999999 is not an Error while its neighbour at 4.5000001 passes.

**Foreground**, first that is set: the run's `color` → the resolved named
style's colour (`resolveNamedStyle`) → `DefaultStyles` → `#000000`.

**Background**, first that is set: the run's `highlight` (mapped through the
same `HIGHLIGHT_HEX` palette the renderer uses, `src/lib/renderFormatting.ts:22`)
→ the run's `shadingFill` → the *paragraph's* `shadingFill`, its own or its
style's (SPEC-074 AC-27) → the enclosing table cell's `shadingFill`
(`src/lib/renderFormatting.ts:315`) → the page, `#ffffff`
(`src/components/DocPage.vue:309`).

That chain is exhaustive by construction, and it always terminates at the
page. SPEC-074 added the paragraph step: text on a shaded paragraph is read
on the shading, so that is what it is checked against — the step this spec
originally called impossible ("the model has no paragraph shading"). Run
shading beyond a highlight is still ruled out of SPEC-013. A section's
`pageBackground` (SPEC-074 AC-23) is deliberately NOT a step: it sits behind
the whole sheet and does not print by default, so judging body text against
it would report contrast the reader on paper never sees. The chain is
otherwise the renderer's own resolution order — the checker must report the
contrast of what was actually painted, and a parallel implementation would
eventually disagree with the screen.

A run whose `highlight` is `false` (an explicit "No colour", distinct from
unset) falls through to the next step like an unset value.

### Looks like a heading
Flagged (as a Tip) when **all four** hold for a paragraph:

1. its style resolves to Normal — a paragraph already carrying a heading style
   is not a candidate;
2. every run in it is bold, **or** its effective size is at least 1.2× the
   document's `DefaultStyles.fontSizePt`;
3. its text is 100 characters or fewer;
4. it is followed by at least one non-empty paragraph.

All four, because any one alone is ordinary prose. Bold text is a lead-in; short
text is a list item or a caption; large text is a pull quote. It is the
combination — short, emphasised, standing above content, in a document whose
author never reached for a heading style — that indicates a heading someone
typed by hand. Conditions 3 and 4 are what keep the check from flagging every
bold sentence in a paragraph of prose.

1.2× rather than a fixed point size because documents set their own body size:
+2 pt is a heading at 11 pt body and invisible at 24 pt.

### Colour alone
Flagged (as a Warning) when a run's `color` differs from the paragraph's
dominant colour — the colour covering the most characters — and the run carries
**no other distinguishing formatting**: not bold, not italic, not underlined,
not a different size, not a different named style.

The case this exists for is "the items in red are mandatory", which conveys
nothing to a reader who cannot see the red. The same sentence set in bold red is
not flagged: the bold carries the meaning on its own.

Hyperlinks are the one carve-out worth naming, and they are flagged: a link
rendered in the link colour with no underline is distinguished from its
surrounding text by colour alone, which is exactly the failure this check is
for.
