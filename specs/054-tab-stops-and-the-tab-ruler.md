---
id: SPEC-054
title: Tab stops and the tab ruler
status: draft
area: formatting
depends_on: [SPEC-014, SPEC-022, SPEC-032]
---
## Why
Tab stops are how a signature block lines up, how a price list gets its column
of right-aligned figures, how a contents page gets its dot leader. They are
also the last piece of paragraph geometry the product does not have.

Most of the machinery is already there and does nothing. SPEC-032 shipped tab
stops as a paragraph property because numbered lists need them to align their
markers: `ParagraphFormatting.tabStops`, a ruler that draws markers, a Tabs
dialog that edits them, and round-trips through both formats. What no part of
the product does is **advance a tab to one**. Nothing in the renderer reads
`tabStops`; `white-space: pre-wrap` hands U+0009 to the browser, which jumps to
its own default grid. So a document authored in Word arrives with its stops
intact, saves with them intact, and lays out as if they were not there.

This spec closes that, and the three smaller gaps around it: there is no
document-level default tab interval, no bar stop, and no stated precedence for
the Tab key — which today has three owners (SPEC-014, SPEC-018, SPEC-022) whose
contexts overlap in a table cell containing a list.

## Acceptance criteria
- [ ] A tab advances to the next stop and the text after it aligns per that stop's alignment — the behaviour SPEC-032's stored stops have never had; `tabStops` is read by the renderer, not left to `white-space: pre-wrap`
- [ ] Clicking the ruler adds a stop of the currently selected type at that position; dragging a marker moves it live; dragging it off the ruler removes it
- [ ] A type selector at the ruler's origin cycles through the five alignments and shows the current one, and each stop's marker on `DocumentRuler.vue` is drawn distinctly per alignment
- [ ] The existing `TabStopsDialog.vue` gains the bar alignment and the document's default interval, keeping its position/alignment/leader list, Set, Clear and Clear All
- [ ] The default tab interval is 1.27 cm, settable per document in the Tabs dialog, and applies wherever no explicit stop follows the caret
- [ ] Pressing Tab advances the caret to the next stop, and text laid out after it aligns per that stop's type; past the last explicit stop, default intervals resume
- [ ] A decimal stop aligns on the decimal separator of the document's language (SPEC-035), falling back to alignment on the first separator character in the run and to right alignment when the text has none
- [ ] A bar stop draws a vertical rule at its position through the paragraph's line height and consumes no horizontal advance
- [ ] A leader fills the space the tab advanced across with dots, hyphens or an underscore, in the run formatting of the tab itself
- [ ] Tab stops apply to every paragraph the selection touches, using SPEC-014's "touched by the selection" rule, and are inherited from the paragraph style (SPEC-019) until the paragraph sets its own — setting any stop on a paragraph replaces the inherited list rather than merging with it
- [ ] The Tab key resolves to exactly one behaviour per context, per the precedence table in "What the Tab key does" below, and no other spec's Tab binding survives ambiguously
- [ ] "Insert → Tab character" inserts a literal tab regardless of context, and is the only way to get one inside a table cell
- [ ] With formatting marks shown (SPEC-028), a tab renders as `→` spanning the distance it advanced
- [ ] In a right-to-left paragraph, stop positions are measured from the start (right) edge, mirroring SPEC-014's decision that indents are logical and only their physical mapping flips
- [ ] Tab stops round-trip through `.docx` (`w:tabs/w:tab` with `@w:val`, `@w:pos` in twips, `@w:leader`; the document default as `w:defaultTabStop` in `settings.xml`) and `.odt` (`style:tab-stops/style:tab-stop` with `style:position`, `style:type`, `style:char`, `style:leader-style`/`style:leader-text`; the default as `style:tab-stop-distance`)
- [ ] Saving to `.odt` reports a `DropRecord` for every bar stop, because ODF has no equivalent — see "Stop types"
- [ ] Opening the fixtures below lays out a right-aligned figure column, a decimal-aligned figure column, a dot-leader contents line and a bar-ruled paragraph at the same positions the source application shows

## Out of scope
Tab stops inside table cells beyond what a paragraph already carries — cell
alignment is SPEC-023's. Per-run tab overrides. Word's "tab stop after list
number" as a separate concept from the paragraph's own stops. Dragging a stop
to snap against another paragraph's stops. Hanging-indent presets. Converting
tabs to a table, or tabs to spaces.

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** What matters here: `ParagraphFormatting` stores centimetres and
points (`indentStartCm`, `indentEndCm`, `indentFirstLineCm`, `spaceBeforePt`),
and this spec's positions are centimetres like its neighbours.

The tab-stop model already exists, shipped by SPEC-032:

```ts
// src/lib/documentModel.ts — as it ships today
export interface ParagraphFormatting {
  …                             // SPEC-014's alignment, indents and spacing
  /** Custom tab stops (SPEC-032), ordered by position. Absent means the
   *  default tab grid only. */
  tabStops?: TabStop[]
}
export interface TabStop {
  positionCm: number            // from the paragraph's start edge, not the page
  alignment: TabStopAlignment
  leader?: TabLeader            // absent means none
}
export type TabStopAlignment = "left" | "right" | "center" | "decimal"
export type TabLeader = "dot" | "hyphen" | "underscore"
```

**This spec widens `TabStopAlignment` with `"bar"`** and changes nothing else
about the shape. `leader` keeps its absent-means-none convention rather than
gaining a `"none"` member — absent and `"none"` would be two spellings of one
state, and `equalStructural` treats absent and empty as the same value anyway.
The leader names stay OOXML's (`dot`/`hyphen`/`underscore`), which is what the
`.docx` writer already emits.

`tabStops` is held sorted by `positionCm` ascending, with at most one stop per
position — setting a second stop at an existing position replaces it. No new
equality rule is needed: `equalStructural` compares plain objects and arrays
already, and the sort invariant means two writers that emit stops in different
orders still produce an equal model.

The document-level default interval is genuinely new — nothing stores one
today, and `INDENT_STEP_CM` in `src/lib/commands.ts` is the indent button's
step, not a tab grid. It belongs on `DefaultStyles`:

```ts
export interface DefaultStyles {
  …
  /** Absent means 1.27 cm. `w:defaultTabStop` (.docx) /
   *  `style:tab-stop-distance` (.odt). */
  defaultTabStopCm?: number
}
```

### A tab is already a character
This needs no decision — `src/lib/documentModel.ts` states it on `Run`:

> Tabs and forced line breaks are preserved as literal `"\t"` / `"\n"`
> characters in `text` — renderers use `white-space: pre-wrap` so both display
> exactly as the source document intended.

So a tab is U+0009 inside `Run.text`, the writers convert it to `w:tab` and
`text:tab` on the way out, and there is no `Inline` union or `LineBreak` node
for it to be an alternative to. The spec records this because every rule below
depends on it: find-and-replace (SPEC-017), word count (SPEC-036) and the
plain-text clipboard flavour (SPEC-016) all read that same string, and all
three need a tab to be one character in it.

What this spec does add is one invariant, because nothing else in the model
needs it: **U+0009 and U+000A are the only control characters permitted in
`Run.text`.** Readers strip every other C0 control (U+0000–U+0008,
U+000B–U+001F) on the way in and warn when they do.

Note the consequence for rendering: `white-space: pre-wrap` makes the browser
advance a tab to its own default stop, which is not this spec's stop list. The
renderer must measure and place tab advances itself rather than leaving U+0009
to the browser.

### Stop types
| `TabStopAlignment` | Aligns | `.docx` `w:tab/@w:val` | `.odt` `style:tab-stop/@style:type` |
| --- | --- | --- | --- |
| `left` | text starts at the stop | `left` | `left` |
| `center` | text centres on the stop | `center` | `center` |
| `right` | text ends at the stop | `right` | `right` |
| `decimal` | the separator sits on the stop | `decimal` | `char` + `style:char="."` |
| `bar` *(new)* | draws a rule; no advance | `bar` | *(no equivalent)* |

The first four ship already; only `bar` is added here.

**ODF has no bar tab.** A `.docx` → `.odt` save drops every bar stop and
reports it with SPEC-008's existing machinery in `src/lib/opaque.ts`. Reusing
`kind: "opaqueBlock"` would be wrong, so `DropKind` widens by one member — as
SPEC-013 (`runShading`), SPEC-024 (`svgImage`) and SPEC-039
(`trackedPropertyChange`) have each already widened it:

```ts
| "tabStop"     // a tab stop the target format cannot express
```
with `location` naming the paragraph and `reason` naming the type. This is the
whole of the format asymmetry; every other type maps in both directions.

`TabLeader` maps as `w:tab/@w:leader` (`dot`/`hyphen`/`underscore`, with
`none` written for an absent leader) and as `style:leader-style`
(`dotted`/`dash`/`solid`, `none` when absent) with `style:leader-text` carrying
the character. Both mappings ship; this spec only makes the leader render.

### What the Tab key does
Three merged specs bind Tab and none of them states what happens when their
contexts overlap — a list inside a table cell binds two rules today. The
precedence, most specific first, is exhaustive:

| # | Context | Behaviour | Owner |
| --- | --- | --- | --- |
| 1 | Caret in a table cell | Move to the next cell; Shift+Tab to the previous; Tab in the last cell appends a row | SPEC-022 |
| 2 | Caret collapsed at offset 0 of a list item | Demote the item; Shift+Tab promotes | SPEC-018 |
| 3 | Caret collapsed at offset 0 of any other paragraph | Increase indent by 1.27 cm; Shift+Tab decreases, floored at 0 | SPEC-014 |
| 4 | Anywhere else — mid-paragraph, a non-collapsed selection, the start of a wrapped line | Insert U+0009 | SPEC-014 |

Rule 1 wins unconditionally inside a cell, including for a list in that cell,
because cell-to-cell navigation is the behaviour every word processor has and
the one users will press Tab expecting.

That leaves no key for a literal tab in a table cell. Word uses Ctrl+Tab;
browsers reserve it for switching tabs and it cannot be intercepted. So the
menu carries **"Insert → Tab character"** (SPEC-049), which works everywhere
and is the documented route inside a cell. Inventing a novel chord for a
keystroke people press once a year would be worse than a menu item.

The bindings are registered in the SPEC-021 shortcut registry and appear in the
shortcut reference, so the precedence above is discoverable rather than folklore.

### Ruler
The horizontal ruler is SPEC-014's, showing start indent, first-line indent,
end indent and margin markers. This spec adds tab markers to it and the type
selector at its origin. It does not change the indent markers, the drag
behaviour or the RTL mirroring SPEC-014 defines; a tab marker mirrors with the
rest of the ruler because `positionCm` is start-relative.

### Specs this changes
- `specs/032-multilevel-lists-and-numbering.md` — **merged, and not undone by
  this spec.** Its criterion "Custom tab stops are supported (left, right,
  centre, decimal, with leader characters), settable via the ruler and a tab
  dialog" is satisfied: the model, `DocumentRuler.vue`, `TabStopsDialog.vue`
  and both format mappings all exist. What it never claimed, and what this spec
  adds, is that a tab *advances* to a stop. Its criterion is amended only to
  point here for the alignment and leader sets, so the two specs cannot
  disagree about them.
- `specs/014-paragraph-alignment-indents-spacing.md` — its out-of-scope line
  "Tab stops (defer to SPEC-032's out of scope note; custom tab stops are a
  later spec)" is resolved: the later spec is this one. Its Tab-key criteria
  are unchanged, and become rows 3 and 4 of the precedence table.
- `specs/022-tables-insert-and-edit.md` and
  `specs/018-bulleted-and-numbered-lists.md` — their Tab criteria are unchanged
  in behaviour and become rows 1 and 2. What changes is that the overlap
  between them is now decided rather than left to whichever handler runs first.
- `specs/008-round-trip-preservation.md` — `DropRecord['kind']` gains
  `'tabStop'`, as SPEC-050 widens the same union.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the Insert menu gains
  "Tab character" and the Format menu gains "Tabs…".

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic`, **which is not implemented**; this spec does not need it.
- **`equalStructural`** — for the same-format round trip of the fixtures below,
  with SPEC-014's tolerance rule applied to `positionCm`: no converted position may
  drift by more than 0.01 cm, since `.docx` stores twips and `.odt` stores
  centimetre strings.
- For cross-format saves, assert `equalStructural` between the source model
  masked by the writer's `dropped` records and the re-opened model. Bar stops
  are the reason that masking is needed rather than a formality.

`equalBytes` is not used by this spec.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). Adds
`fixtures/docx/word-tabs.docx` and `fixtures/odt/lo-tabs.odt`, each containing
a right-aligned figure column, a decimal-aligned column with one value lacking
a separator, a contents line with a dot leader, a paragraph with a bar stop, a
paragraph relying only on the default interval, and one RTL paragraph with
stops. Manifest entries name each paragraph's stop positions, types and leaders.
