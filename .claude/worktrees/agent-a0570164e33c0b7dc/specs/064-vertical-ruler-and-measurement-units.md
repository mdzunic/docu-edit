---
id: SPEC-064
title: The vertical ruler and measurement units everywhere
status: draft
area: layout
depends_on: [SPEC-027, SPEC-029, SPEC-054, SPEC-063]
---
## Why
The product has one ruler and it runs one way. `DocumentRuler.vue` draws
the horizontal scale — page width, margins, indent markers, column
boundaries, tab stops — and there is nothing along the left edge of the
page. So the horizontal geometry of a paragraph is direct-manipulable and
the vertical geometry of the *page* is not: changing a top margin means
opening Page setup, typing a number, and looking at the result; the same
for a bottom margin, and there is no visual indication at all of where the
header and footer bands sit, which is the single most confusing thing about
headers in every word processor. Word, Writer and Pages all put a vertical
ruler there, and all three use it for exactly those three jobs.

The document model already expects it. `DocumentModel.showRuler` is
documented as "whether the **horizontal** ruler is shown", which is the
field naming its own gap.

The second half of this spec is the unit preference. `src/lib/core/units.ts`
ships a complete `MeasurementUnit` type — `"cm" | "mm" | "in" | "pt"` — with
`toCm`, `fromCm` and `roundForDisplay`, and `DocumentModel.measurementUnit`
persists a choice. Exactly one surface honours it: `PageSetupDialog.vue`.
Everywhere else in the product a measurement is centimetres whatever the
user asked for — the ruler's ticks, the Tabs dialog, the indent fields, the
table properties dialog, image size, shape position. A US user who sets
inches in Page setup gets inches in one dialog and centimetres in the other
nine. Adding a second ruler without fixing that would add a second
centimetres-only surface, so the two belong in one spec.

## Acceptance criteria
- [ ] A vertical ruler runs down the left edge of the page area in print layout, aligned to the page, scaled with the zoom (SPEC-063), and scrolling with the document
- [ ] It shows the page's top and bottom margin boundaries as draggable markers; dragging one changes the caret section's margin live and commits one undo entry on release, mirroring how the horizontal ruler's margin edges behave
- [ ] It shows the header and footer bands as distinct regions with their own draggable boundaries — header distance from the top edge, footer distance from the bottom — writing to the section's header/footer distances (SPEC-029)
- [ ] Ticks are labelled per the current measurement unit, with major/minor divisions per "Tick scheme", and the numbering origin is the top of the text area, counting outward like the horizontal ruler counts from the text area's start
- [ ] While the caret is in a table, the vertical ruler shows that table's row boundaries as markers, and dragging one sets that row's `minHeightCm` — the same edit dragging the row boundary in the page performs, with the same undo label
- [ ] While a header or footer is being edited (SPEC-029's editing mode), the vertical ruler marks the band being edited so it is obvious which region has focus
- [ ] `View → Ruler` toggles **both** rulers together, keeping one control and one persisted flag; `DocumentModel.showRuler`'s doc comment stops saying "horizontal"
- [ ] The vertical ruler is hidden in web layout and in print/PDF output, and is present in focus mode only if it was visible before entering it (SPEC-063)
- [ ] A measurement-unit preference applies **app-wide**: every numeric length field, every ruler tick label and every read-out of a length in the UI is displayed and accepted in the chosen unit, per "Where units apply"
- [ ] The unit selector moves out of the Page setup dialog into Preferences-style reach: it is set from either the Page setup dialog (where it is today, unchanged) or by right-clicking either ruler, and both write the same stored value
- [ ] Typing a value with an explicit unit suffix into any length field — `2cm`, `0.5in`, `12pt`, `10mm` — is accepted and converted, whatever the current display unit
- [ ] Values are rounded for display per `roundForDisplay(value, unit)` and never rounded in the model: switching units back and forth does not drift a stored measurement, asserted by a unit test over the four units
- [ ] Both rulers are keyboard-operable: each marker is focusable, arrow keys nudge it by one minor division, Shift+arrow by one major division, and each announces its name and current value
- [ ] Both rulers are `role="group"` with a name, their markers `role="slider"` with `aria-valuenow`/`aria-valuetext` in the display unit, so a screen reader reads "Top margin, 2.5 centimetres"
- [ ] Dragging any ruler marker under a non-100% zoom lands accurately, using SPEC-063's `clientToContentPx` rather than raw client coordinates
- [ ] An e2e test drags the top-margin marker, asserts the page's first line moved and the section's margin changed, then switches the unit to inches and asserts the ruler's labels and the Page setup dialog agree

## Out of scope
A vertical ruler in web layout, where there are no page boundaries to
measure against. Vertical tab stops (no such construct exists in either
format). Dragging column *heights*. Changing what a margin, header
distance or row height means — SPEC-027, SPEC-029 and SPEC-023 own those
and are unchanged. A per-section ruler unit. Ruler guides / draggable
alignment guides pulled from the ruler (SPEC-072's snapping territory).
Making the measurement unit a saved document property visible to other
applications — it stays an app-side preference persisted with the record,
exactly as it is today.

## Reference

### The horizontal ruler this parallels
`src/components/DocumentRuler.vue`, as shipped:

```ts
const props = defineProps({
  pageWidthCm:   { type: Number, required: true },
  marginLeftCm:  { type: Number, required: true },
  marginRightCm: { type: Number, required: true },
  /** First touched paragraph's formatting — indents plus direction. */
  formatting:    { type: Object, required: true },
  /** Caret section's column boundaries, or null for a single column. */
  columns:       { type: Object, default: null },
  /** First touched paragraph's custom tab stops. */
  tabStops:      { type: Array,  default: () => [] },
  disabled:      { type: Boolean, default: false },
});
const emit = defineEmits(["update", "open-tabs-dialog"]);
```

Its geometry comes from `usePageSetupDialogs.ts` (`rulerPageWidthCm`,
`rulerMarginLeftCm`, …) and it emits an `update` the view turns into a
model edit. The vertical ruler is a sibling component with the same shape:

```ts
// src/components/DocumentVerticalRuler.vue
const props = defineProps({
  pageHeightCm:      { type: Number, required: true },
  marginTopCm:       { type: Number, required: true },
  marginBottomCm:    { type: Number, required: true },
  /** Header band distance from the page's top edge (SPEC-029). */
  headerDistanceCm:  { type: Number, default: 0 },
  /** Footer band distance from the page's bottom edge (SPEC-029). */
  footerDistanceCm:  { type: Number, default: 0 },
  /** Row boundaries of the table containing the caret, as distances from
   *  the page's top edge; empty when the caret is not in a table. */
  tableRows:         { type: Array,  default: () => [] },
  /** "body" | "header" | "footer" — which band is being edited. */
  activeBand:        { type: String, default: "body" },
  unit:              { type: String, default: "cm" },
  disabled:          { type: Boolean, default: false },
});
const emit = defineEmits(["update"]);
```

Both rulers stay presentational: they emit a proposed value, and the view
turns it into a command with an undo entry, exactly as the horizontal one
does today. Neither reaches into the model.

Placement: the vertical ruler occupies a fixed-width strip to the left of
the pages container, top-aligned to the first page and translated by the
same scroll offset, so page and ruler cannot drift apart. It is inside
SPEC-063's zoom transform for scale, or scaled by the same factor —
whichever keeps its labels crisp — but a ruler centimetre must equal a
document centimetre at every zoom level, which is the testable statement.

### Tick scheme
Per unit, major ticks are labelled, minor ticks unlabelled:

| Unit | Major | Minor |
| --- | --- | --- |
| cm | 1 cm | 0.5 cm, quarter marks below 75% zoom suppressed |
| mm | 10 mm | 5 mm |
| in | 1 in | 1/8 in (1/4 in below 75% zoom) |
| pt | 72 pt | 36 pt |

Below 50% zoom only major ticks draw; above 150% the next subdivision
appears. Ticks are geometry, not chrome, so they derive from `PX_PER_CM`
and the zoom rather than from a token (SPEC-058's third exemption).

### Where units apply
`src/lib/core/units.ts` already provides everything needed:

```ts
export type MeasurementUnit = "cm" | "mm" | "in" | "pt";
export function asMeasurementUnit(value: string): MeasurementUnit;
export function toCm(value: number, unit: MeasurementUnit): number;
export function fromCm(cm: number, unit: MeasurementUnit): number;
export function roundForDisplay(value: number, unit: MeasurementUnit): number;
```

The model stays centimetres everywhere — this is a display and input
concern only, and no reader or writer is touched.

Surfaces that must honour the preference after this spec: both rulers;
`PageSetupDialog.vue` (already does); `TabStopsDialog.vue`;
`CustomSpacingDialog.vue`'s length fields; the paragraph indent fields;
`TablePropertiesDialog.vue` (column widths, row heights, cell padding);
`ImageOptionsDialog.vue` (size and position); the shape toolbar's size,
position and rotation read-outs; `ColumnsDialog.vue` (width and spacing);
`HeaderFooterOptionsDialog.vue` (distances); and the drawing position
dialog SPEC-072 adds.

Point-valued things stay points regardless of the unit: font size, space
before/after, line spacing, outline width, border width. The unit
preference is about **lengths on the page**, not type metrics, and a user
who picks inches does not want 0.14in fonts.

To keep that consistent, the shared parsing and formatting lives in one
place rather than in each dialog:

```ts
// src/lib/core/units.ts — added by this spec
/** Parses a user-typed length. Accepts a bare number in `unit`, or an
 *  explicit suffix ("2cm", "0.5in", "12pt", "10mm", case- and
 *  space-insensitive). Returns centimetres, or null when unparseable. */
export function parseLengthToCm(input: string, unit: MeasurementUnit): number | null;

/** Formats centimetres for a field or a label in `unit`, rounded per
 *  roundForDisplay, with no unit suffix (the field labels the unit). */
export function formatLength(cm: number, unit: MeasurementUnit): string;
```

Both are pure, live in the product-neutral core layer, and are unit-tested
directly — including the round-trip property that
`parseLengthToCm(formatLength(x, u), u)` returns `x` to within
`roundForDisplay`'s precision, for each of the four units.

### Storage
The preference stays where it already is: `DocumentModel.measurementUnit`,
absent meaning `"cm"`. It is a document-level app preference, persisted
with the record and mapped to nothing in either file format — the same
category as `showRuler` and `showFormattingMarks`. A future Preferences
spec may add an app-level default that new documents inherit; this spec
does not require one, and does not change where the value lives.

### Specs this changes
- `specs/027-page-setup.md` — its AC-6/AC-7 unit preference is widened from
  "the Page setup dialog" to the whole UI. The dialog itself is unchanged.
  Its AC-14 ruler toggle now governs both rulers.
- `specs/029-headers-footers-and-page-numbers.md` — header and footer
  distances gain a direct-manipulation affordance. Their meaning, storage
  and round-tripping are untouched.
- `specs/023-table-formatting-and-cell-operations.md` — row height gains a
  second way to be dragged (the ruler); it is the same edit, with the same
  undo label, and the in-page row boundary drag stays.
- `specs/054-tab-stops-and-the-tab-ruler.md` — the horizontal ruler's tab
  interaction is unaffected; its position fields display in the chosen
  unit.
- `specs/002-create-blank-document.md` — `showRuler`'s doc comment drops
  the word "horizontal". No field is added or removed.
