---
id: SPEC-081
title: Slide text editing and formatting
status: draft
area: editing
depends_on: [SPEC-079, SPEC-080]
---
## Why
SPEC-080 can restructure a deck but every slide is still frozen content:
placeholders from a file can't be retyped, and a blank deck's title can't
be filled in. This spec makes text frames editable — click in, type, format
— and builds the piece the whole editing milestone hangs on: the **frame →
XML serialisers** that rebuild a dirty slide for both formats while
re-emitting untouched siblings verbatim (SPEC-080 defined *when* parts
rebuild; this spec defines *what* a rebuilt slide contains). Once text
edits round-trip, SPEC-082/083/085/086 are additional frame kinds and
properties on the same machinery. **This file is the normative source for
slide-part serialisation in both formats.**

## Acceptance criteria
- [ ] Double-clicking a text frame (or pressing Enter with it selected) starts a text-editing session in place: a caret appears at the click point, the frame shows an editing outline, and Escape ends the session returning to frame selection; clicking outside any frame deselects
- [ ] A single click selects a frame (selection handles rendered; resize/move behaviour arrives in SPEC-082 — this spec renders the selection box only); Tab/Shift+Tab cycle frame selection in z-order
- [ ] Typing, Enter (new paragraph), Backspace/Delete, and caret/selection movement (arrows, Home/End, Shift-extensions, Cmd/Ctrl+A within the frame) work in the editing session; the model's `paragraphs` reflect the frame's content after every input, and the stage re-renders from the model (the DOM never diverges: one normalisation pass per input event)
- [ ] Paste inserts plain text only (formatting stripped; newlines become paragraph breaks); cut/copy from a slide text selection put plain text on the clipboard
- [ ] An empty placeholder renders its prompt ("Click to add title" / "Click to add text" by role) outside editing sessions; the prompt never enters the model, either format's files, or the clipboard
- [ ] The view gains a presentation toolbar with: font family (the `core/fonts.ts` `STANDARD_FONT_FAMILIES` list), font size (6–96, stepper + typed entry), bold/italic/underline toggles, text colour swatch, alignment (left/centre/right/justify), bullet toggle (disc), and indent +/− (level 0–8); every control reflects the current selection's state (mixed state renders indeterminate) and applies to the selection, or to the whole frame when it is selected without an editing session
- [ ] Cmd/Ctrl+B/I/U work in an editing session; all SPEC-080 shortcuts still work outside one; typing bursts coalesce into single undo entries (a pause ≥ 1s, ending the session, or any non-typing command closes the burst), and formatting commands are one entry each
- [ ] Any edit marks the slide dirty (`dirty: true`, and for odp drops `sourceXml`); saving then rebuilds exactly the dirty slides per the serialisation rules below, while untouched slides and all unmodelled parts stay byte-verbatim (part-level `equalBytes` against the source)
- [ ] A rebuilt pptx slide preserves what this editor doesn't model: `preserved` XML (`p:bg`, `p:transition`, `p:timing`, `p:clrMapOvr`, `p:extLst`, root attributes) re-emitted in schema position, foreign/group frames re-emitted verbatim in z-order, and the slide's `.rels` keeping every original relationship (layout included) with original ids
- [ ] A rebuilt odp page does the same: original `draw:page` attributes, `presentation:notes`/`anim:*` from `preserved`, foreign frames verbatim, generated automatic styles added without touching existing ones
- [ ] Effective-value fidelity: a rebuilt slide writes explicit run/paragraph properties equal to the model's effective values, so text looks identical before and after a save→reopen cycle (`slideOutline` + formatting assertions on the reopened model), even though inheritance is baked (documented trade)
- [ ] Editing the frame the outline derives a title from updates the filmstrip thumbnail and any outline-derived label live
- [ ] Save → reopen → the edited text and formatting survive in both formats (unit round trip over the fixture decks, edited via commands, reopened via `readPresentationBytes`)
- [ ] e2e (`e2e/SPEC-081.spec.ts`): open the pptx fixture → double-click the title, select-all, type a new title, bold it via toolbar, add a bulleted body line on slide 2 → download → re-open the download → both edits visible on the stage; a second download with zero further edits re-downloads the edited deck stably

## Out of scope
Editing inside group frames (SPEC-082's ungroup unlocks them) and foreign
frames. Inserting/moving/resizing/deleting frames (SPEC-082). Rich paste,
paste of images, drag-and-drop text. Find and replace, spell check on
slides. Numbered-bullet authoring (the "number" bullet renders read-only;
the toggle writes "disc" only — numbered authoring rides with a future
lists pass). Text autofit; overflowing text keeps overflowing exactly as
SPEC-079 renders it. Line spacing, letter spacing, highlight, links.
RTL/vertical text. Notes editing (SPEC-083).

## Reference

### Dependencies restated
- SPEC-079's model (`TextFrame { role?, phType?, phIdx?, paragraphs,
  fill? }`, `SlideParagraph { runs, align?, bullet?, level? }`, `SlideRun
  { text, bold?, italic?, underline?, fontFamily?, sizePt?, color? }`),
  stage components, effective-value semantics, `preserved` capture, and
  the colour/geometry conventions (cm, pt, `#rrggbb`).
- SPEC-080: `Slide.dirty`, `structureDirty`, `odpContentEnvelope`, the
  substitution save path in both writers (this spec plugs the "serialise
  a dirty slide" hole SPEC-080 left), `usePresentationUndo.record`,
  write-through persistence, `PresentationView.size.spec.ts` ratchet.
- Equality: part-level `equalBytes`; model comparisons via SPEC-080's
  `strippedForCompare` (minted ids masked).

### Editing session architecture
`src/views/presentation/useSlideTextEditing.ts` (+
`useFrameSelection.ts`), wired by `PresentationView`; the stage's text
frame component hosts a `contenteditable="true"` div only while its frame
is the editing target. Input handling is
**normalise-after-input**: `beforeinput` is intercepted only to reroute
paste (plain text) and to tag undo bursts; after the browser applies an
edit, `normalizeTextFrameDom(el): SlideParagraph[]` walks the DOM —
block-level children and `<br>` split paragraphs, styled spans carry run
marks (the renderer only ever emits `<span>` runs with inline
`font-weight`/`font-style`/`text-decoration`/`font-family`/`font-size`/
`color`, so normalisation reads back exactly what rendering wrote) — and
the result replaces the frame's `paragraphs`, followed by a targeted
re-render that restores the caret via a `(paragraphIndex, runIndex,
offset)` selection address. Formatting commands never round-trip through
`document.execCommand`: they edit the model (splitting runs at selection
edges) and re-render. The command layer is framework-free:
`src/lib/presentation/text.ts` — `applyRunFormat(frame, range, patch)`,
`applyParagraphFormat(frame, paragraphRange, patch)`,
`toggleBullet`, `setIndentLevel` (clamped 0–8), plus the selection-address
maths, all unit-tested without a DOM.

### Dirtying
Every command from this spec: `slide.dirty = true`, `delete
slide.sourceXml` (odp), then `persist()`. `preserved` and `partName` stay
— the rebuild consumes them. `model.structureDirty` is untouched by text
edits (pptx `presentation.xml` does not rebuild for a text-only edit; the
dirty slide's part substitutes in place, same name, same order).

### pptx slide serialisation (`presentation/pptx/serialize.ts`)
`serializeSlideXml(slide, model, rels): string` emits:
```
<?xml … ?><p:sld {preserved root attrs or the standard xmlns set}>
  <p:cSld>{preserved p:bg}<p:spTree>
    (fixed p:nvGrpSpPr/p:grpSpPr scaffold, ids minted 1..)
    …frames in z-order…
  </p:spTree></p:cSld>
  {preserved p:clrMapOvr}{preserved p:transition}{preserved p:timing}{preserved p:extLst}
</p:sld>
```
- `TextFrame` → `p:sp`: `p:nvSpPr` (`p:cNvPr id/name`, `p:nvPr` with
  `p:ph type={phType} idx={phIdx}` when role present); `p:spPr` with
  `a:xfrm` (`a:off`/`a:ext` via `cmToEmu`, `rot` = `rotationDeg*60000`
  when set), `a:prstGeom prst="rect"` + empty `a:avLst`, fill
  (`a:solidFill/a:srgbClr` or `a:noFill`); `p:txBody` with `a:bodyPr`
  (empty), one `a:p` per paragraph: `a:pPr` carrying `algn` (`l/ctr/r/just`,
  omitted for left), `lvl` (omitted for 0), and bullets — `a:buNone` for
  none, `a:buChar char="•"` for disc, `a:buAutoNum type="arabicPeriod"`
  for number; each run `a:r` with `a:rPr lang="en-US" sz={pt*100} b/i/u`
  + `a:solidFill/a:srgbClr` + `a:latin typeface` (every effective value
  written explicitly, per the fidelity AC), then `a:t` with escaped text.
  An empty paragraph emits `a:p` with `a:endParaRPr`.
- `ShapeFrame` → same shell with its mapped `prst` + `a:avLst` from
  adjustments (`core/drawingMl.ts` tables), stroke as `a:ln w={pt*12700}`
  + `a:solidFill`; text body only when paragraphs exist.
- `ImageFrame` → `p:pic` with `a:blip r:embed` resolved through the
  slide's rebuilt `.rels` (this spec: images only ever *re*-serialise on
  an already-image-bearing dirty slide — the rels entry for its media
  part is carried from the original rels; SPEC-082 adds creation).
- `GroupFrame`/`ForeignFrame` → their `xml` verbatim.
The `.rels` for a rebuilt slide = the original file's relationships,
verbatim ids and targets (plus SPEC-082's additions later). From-scratch
slides (SPEC-080) get `rId1 → slideLayout`.

### odp page serialisation (`presentation/odp/serialize.ts`)
`serializePage(slide, model, styles): string` emits `<draw:page {original
attrs, or draw:name="page-N" + master ref for new slides}>` + frames +
`{preserved presentation:notes}{preserved anim}` `</draw:page>`.
- Frames → `draw:frame` (`presentation:class` for roles, `svg:x/y/width/
  height` as `Ncm` strings, `draw:transform rotate(...)` when rotated)
  containing `draw:text-box`, `draw:image xlink:href` or the mapped
  shape element; text as `text:p`/`text:span` with generated automatic
  styles; bullets via `text:list` nesting (one level per `level`,
  disc/number list styles).
- Generated automatic styles: collected by a `StyleRegistry` during
  serialisation, named in a reserved namespace — paragraph `de-P{n}`,
  text `de-T{n}`, graphic `de-G{n}`, list `de-L{n}` — and spliced into
  `content.xml` immediately before `</office:automatic-styles>` (the
  envelope's prefix always contains that element for files LibreOffice/
  Impress produce; if a source file lacks it, insert
  `<office:automatic-styles>…</office:automatic-styles>` before
  `<office:body`). Existing automatic styles are never rewritten; name
  collisions are impossible because the `de-` prefix is stripped from no
  source (asserted at read: a source file that *does* use `de-`-prefixed
  names bumps the generated counter past every collision).

### Toolbar
`src/components/presentation/PresentationToolbar.vue`, design-system
controls (same primitives as `DocumentToolbar`; the font/size/colour
widgets are extracted to shared components only if they are already
product-neutral — otherwise presentation-local twins, and a note in the
DocumentView review debt file about the duplication). Testids:
`slides-toolbar`, `slides-font-family`, `slides-font-size`, `slides-bold`,
`slides-italic`, `slides-underline`, `slides-text-color`, `slides-align`,
`slides-bullet`, `slides-indent`, `slides-outdent`. The toolbar must fit
the e2e viewport (memory: Insert-menu 720px lesson) — overflow collapses
into a ⋯ menu like SPEC-061's pattern.

### Tests
- `presentation/text.spec.ts`: run splitting, format application across
  runs/paragraphs, selection addresses, bullet/indent rules.
- `presentation/pptx/serialize.spec.ts` + `presentation/odp/serialize.spec.ts`:
  golden serialisations for each frame kind; preserved-XML placement;
  escaping (`&`, `<`, quotes, non-ASCII); empty paragraphs.
- `presentation/roundTrip.spec.ts`: edit → write → reparse fidelity per
  AC (both formats); untouched-sibling byte identity; stability (write →
  reparse → write is byte-stable for rebuilt slides).
- View specs: session lifecycle, prompt text, toolbar state reflection,
  undo bursts.
- e2e per AC. (Memory: the browser pane's synthetic typing is not a
  keyboard — e2e must use Playwright's real `keyboard.type`, and any
  manual pane-driving during development can't validate focus behaviour.)

### Specs this changes
- SPEC-080's from-scratch writers stop emitting the "empty placeholder
  only" special case — they call these serialisers for every slide.
- SPEC-079's stage becomes the host for selection/editing overlays and
  the empty-placeholder prompt (SPEC-079 rendered empty placeholders as
  empty; the prompt is this spec's addition, and SPEC-079's
  "paragraphs: []" read rule is unchanged).
