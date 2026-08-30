---
id: SPEC-075
title: Watermarks and line numbering
status: draft
area: layout
complexity: normal
story_points: 8
depends_on: [SPEC-028, SPEC-029, SPEC-051, SPEC-074]
---
## Why
Two page-level features were deferred by the specs that owned their
neighbourhood and never picked up: SPEC-029 put "Watermarks" out of scope
while building headers and footers, and SPEC-028 put "Line numbers" out of
scope while building breaks and section properties. They are unrelated
features that share one property — both draw on every page of a section
without being part of the text flow — and both are the kind of thing a
document arrives with rather than something a user goes looking for.

Watermarks are the more visible loss. DRAFT, CONFIDENTIAL and a company
logo behind the text are how organisations mark documents, and in `.docx` a
watermark is a shape anchored in the header with specific properties that
this product currently reads as an ordinary header drawing — so it renders
in the header band, in the wrong place, at the wrong size, instead of
behind the page. That is worse than not supporting it: the document does
not merely lose the watermark, it gains a visual defect.

Line numbering is narrower but non-negotiable where it is needed: legal
filings, contracts, legislative drafts and translation review all require
numbered lines, and there is no way to fake it. Both formats model it as a
section property with a handful of options, so it is small to add and
impossible to work around.

## Acceptance criteria
- [ ] A watermark is settable per section from `Insert → Watermark…`, in two kinds — text and image — per "The watermark model"
- [ ] A text watermark offers the text, a font, a size (including "auto" fitting the page width), a colour, a transparency, an orientation of horizontal or diagonal, and a semi-transparent default
- [ ] An image watermark references a media part through the existing image pipeline (SPEC-024), with a scale and a washout toggle, and does not duplicate the part when the same image is also in the body
- [ ] The dialog offers the common presets — DRAFT, CONFIDENTIAL, SAMPLE, DO NOT COPY, URGENT — and "No watermark" to remove one
- [ ] A watermark renders **behind the body text** on every page of its section, centred on the page, clipped to the page, and is not selectable, editable or hit-testable in the document canvas — it is edited only from its dialog
- [ ] A watermark appears in print, print preview and PDF export exactly as on screen, per "Printing"
- [ ] Applying a watermark to "this section" or "the whole document" is offered, matching SPEC-027's Apply-to control
- [ ] A `.docx` watermark round-trips through the header shape both source applications use, per "Format mapping — .docx", and a watermark read from Word renders behind the page rather than in the header band
- [ ] An `.odt` watermark round-trips through the page style's `style:background-image` for images and through a header-anchored drawing for text, per "Format mapping — .odt"
- [ ] A header that also contains ordinary content keeps that content: adding or removing a watermark never disturbs the header's own paragraphs, and a header drawing that is *not* a watermark is still read as an ordinary drawing
- [ ] Line numbering is settable per section from Page setup — on/off, start value, increment, distance from the text, and restart mode (continuous, each page, each section)
- [ ] Line numbers render in the margin beside each body line at the stated distance, on the start edge (mirroring for RTL, per SPEC-014's logical-direction rule), and count only body text — headers, footers, footnotes, endnotes, tables and floating drawings are not numbered
- [ ] A paragraph can be excluded from line numbering (`Format → Paragraph → Suppress line numbers`), and an excluded paragraph's lines do not advance the counter
- [ ] Line numbers appear in print, print preview and PDF export
- [ ] Both features round-trip: line numbering through `.docx` `w:lnNumType` (`@w:start`, `@w:countBy`, `@w:distance`, `@w:restart`) and `w:suppressLineNumbers`, and through `.odt` `text:linenumbering-configuration` and `text:number-lines` on the paragraph style
- [ ] Neither feature changes pagination: turning a watermark or line numbers on does not move a single line break, asserted by comparing the page-break model before and after
- [ ] Fixtures with a text watermark, an image watermark, a continuously-numbered section and a per-page-restarting section open, render, save and reopen `equalStructural`

## Out of scope
A different watermark on the first page, or on odd and even pages —
one per section. Multiple watermarks. Watermarks in front of text. Editing
a watermark by direct manipulation on the page. Word's decorative art page
borders (SPEC-074 leaves them preserved). Numbering by paragraph rather
than by line. Numbering table rows or list items. Change-bar rules in the
margin — that is SPEC-039's territory. Page background colour and page
borders, which are SPEC-074's.

## Reference

### The watermark model
A watermark is not a drawing in the block list. Both formats store it as
one, but modelling it that way would make it selectable, deletable and
reorderable in the document canvas, which is exactly the behaviour that
makes header-anchored watermarks so awkward in Word. So the model states
what it is, and the readers and writers translate:

```ts
// src/lib/documentModel.ts
export type Watermark =
  | { kind: "text"; text: string; fontFamily: string;
      /** Points, or "auto" to fit the page width. */
      fontSizePt: number | "auto";
      color: string; alpha: number;          // 0..1
      orientation: "horizontal" | "diagonal"; bold?: boolean; italic?: boolean }
  | { kind: "image"; mediaId: string; scalePercent: number; washout: boolean; alpha: number };

export interface Section {
  …
  /** Drawn behind the body on every page of this section. */
  watermark?: Watermark;
}
```

`mediaId` references the same media store `Image` blocks use (SPEC-024), so
one part serves both.

Rendering: a page-level layer between the page background (SPEC-074) and
the body text, painted per page by the existing canvas painter, clipped to
the page, centred, with the text laid out on the diagonal when the
orientation says so. It is not in the DOM's editable subtree, carries no
`drawing-<index>` proxy (SPEC-057's counter is for body drawings), and is
described once per page to assistive technology as
`aria-label="Watermark: DRAFT"` on the page container rather than as a
focusable object.

### Format mapping — .docx
Word stores a watermark as a VML shape (`v:shape` with
`type="#_x0000_t136"` for text, or a `v:shape` with an `v:imagedata` for
an image) inside a `w:pict` in the section's header part, with
`o:allowincell`, a `WordArt`-style text path for the text case, and
`style` attributes positioning it relative to the page. Modern Word also
writes a DrawingML fallback.

- **Read**: a header whose only content is a drawing bearing the watermark
  signature — a `v:shape` with the text-effect type, or a shape named
  `WordPictureWatermark*`/`PowerPlusWaterMarkObject*` — is recognised as a
  watermark, lifted out of the header content into `Section.watermark`,
  and its paragraph is dropped if empty. Anything else in the header
  stays header content.
- **Write**: the watermark is re-emitted into the section's default header
  in that same shape, ahead of the header's own paragraphs. A section with
  a watermark and no other header content writes a header part containing
  only it.

The signature test is stated explicitly because getting it wrong in either
direction is a visible bug: too loose and an ordinary header logo becomes a
watermark; too tight and a Word watermark renders in the header band, which
is today's behaviour.

Where the source's watermark shape carries properties this model does not
express, the shape is preserved verbatim per SPEC-008 and the modelled
fields drive rendering — same-format saves re-emit the preserved shape.

### Format mapping — .odt
ODF splits the two kinds:

- **Image** — `style:background-image` on the page layout's
  `style:page-layout-properties`, with `style:repeat="no-repeat"`,
  `style:position="center center"` and an opacity. This is a true page
  background image and is the natural home for an image watermark.
- **Text** — ODF has no text watermark. It is written as a drawing
  anchored to the header's first paragraph, positioned relative to the
  page, with `style:run-through="background"` so it sits behind the text —
  which is what Writer itself produces for its Format ▸ Watermark command.

Reading mirrors both. A `style:background-image` on the page layout reads
as an image watermark; a background-positioned header drawing whose text
matches Writer's watermark shape reads as a text watermark.

Cross-format, both kinds carry. A property that does not (Word's washout
filter has no ODF equivalent beyond opacity) reports a `DropRecord` of a
new `DropKind` `"watermark"`.

### Printing
A watermark prints. This is the difference between a watermark and
SPEC-074's page background, and it is why they are separate features
despite both drawing behind the text: a background colour is a screen
convenience that prints only on request, and a watermark exists *for* the
printed copy. SPEC-034's export path includes the watermark layer
unconditionally.

### Line numbering
```ts
export interface Section {
  …
  lineNumbering?: {
    startAt: number;                  // default 1
    countBy: number;                  // number every Nth line; default 1
    distanceCm?: number;              // from the text; absent = automatic
    restart: "continuous" | "newPage" | "newSection";
  };
}

export interface ParagraphFormatting {
  …
  /** Excluded from line numbering (w:suppressLineNumbers /
   *  text:number-lines="false"). */
  suppressLineNumbers?: boolean;
}
```

Numbers render in the margin outside the text area, on the paragraph
direction's start edge, in the document's default font at a size derived
from it, in `--text-secondary` (SPEC-058) — they are chrome around the
document's own content, and the formats do not store a colour for them.

The counter advances per **laid-out line** of body text, so it is a
pagination-time value; SPEC-047's incremental pagination already recomputes
lines per page and the counter rides along with it, which is why the
"pagination does not change" AC is achievable — line numbers are read from
the layout, they do not influence it.

Only body text counts. Header, footer, footnote, endnote, table-cell and
floating-drawing lines are skipped, and a paragraph with
`suppressLineNumbers` contributes nothing. This is both formats' behaviour
and worth stating because "should a table's lines count" has no obvious
answer.

### Menu placement
`Insert → Watermark…` opens the watermark dialog; Page setup gains a Line
numbers section beside the existing page and margin controls, with its own
Apply-to. Both register in the `editorCommands` Map
(`insert.watermark`, `format.line-numbers`) per SPEC-049's rule.

### Specs this changes
- `specs/029-headers-footers-and-page-numbers.md` — "Watermarks" leaves
  its out-of-scope list. Header and footer behaviour is unchanged, and
  header content is explicitly protected by the AC above.
- `specs/028-page-and-section-breaks.md` — "Line numbers" leaves its
  out-of-scope list. Sections, breaks and their round-tripping are
  unchanged.
- `specs/034-print-and-export-pdf.md` — the watermark and line-number
  layers join the printed output; page background stays opt-in per
  SPEC-074.
- `specs/008-round-trip-preservation.md` — `DropKind` gains
  `"watermark"`.
