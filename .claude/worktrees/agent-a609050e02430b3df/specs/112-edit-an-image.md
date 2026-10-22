---
id: SPEC-112
title: Edit an image
status: draft
area: editing
complexity: normal
story_points: 8
depends_on: [SPEC-015, SPEC-058, SPEC-063, SPEC-064, SPEC-107, SPEC-108]
---
## Why
An image arrives in this library and can be looked at. It cannot be
straightened, cropped to the part that matters, turned the right way up,
lightened because the scan came out grey, or have the account number covered
before it goes into a PDF. Every one of those is a precondition for the
things SPEC-109–111 build, and every one of them currently means leaving.

The editing model is the interesting decision. This repo's oldest value is
that opening and saving a file must not destroy what it did not understand
(SPEC-008). The same instinct applies to a photograph: an editor that
overwrites the original with a JPEG re-encode of a crop has destroyed
information the user may want back tomorrow. So edits here are a **stack**,
stored beside the untouched source bytes and applied on demand — the source
is always recoverable, the crop is always adjustable, and flattening is an
explicit act with a confirmation, not a side effect of clicking Save.

That choice has one sharp edge, and this file names it rather than shipping
it quietly: **a redaction on a non-destructive stack is a cover, not an
erasure**, because the original pixels are still in the library. The spec
requires the UI to say so and provides the one operation that really does
destroy them.

**This file is the normative source for `ImageEdit`, for the pixel formulas,
for the order edits are applied in, and for what flattening means.**

## Acceptance criteria
- [ ] `MediaModel` gains `edits?: ImageEdit[]` and `renderMediaId?: string` per the Reference: `sourceMediaId` never changes once written, `edits` is an ordered stack applied to the source, and `renderMediaId` names a **derived** `MediaPart` — a cached flattened PNG (or JPEG when the source is one and nothing added alpha) regenerated whenever the stack changes and used by every consumer that wants "the current image": the viewer, the library thumbnail, SPEC-109's compose and SPEC-110's assembly
- [ ] `src/lib/core/imageEdits.ts` is pure and framework-free: `ImageEdit`, `EDIT_KINDS`, `normaliseStack(edits)`, `resolveEditGeometry(edits, naturalPx): EditGeometry` (output pixel size, the source rectangle, the cumulative transform) and `describeEdit(edit): string` for the history list — given only the natural size, never the bytes, so every geometric consequence of a stack is unit-testable without a canvas
- [ ] Seven edit kinds, each specified rather than named: `crop` (a source rectangle in pixels, with aspect presets Free / Original / 1:1 / 4:3 / 3:2 / 16:9 and a straighten angle in degrees applied before the rectangle), `rotate` (90° steps, exact and lossless in geometry), `flip` (`horizontal` | `vertical`), `resize` (target pixels or percent, aspect-locked by default, `nearest`/`bilinear` resampling with bilinear the default), `adjust` (the six channels below), `redact` (opaque rectangles in a chosen colour, default `#000000`) and `annotate` (arrow, line, rectangle, ellipse, freehand, highlight and text, each with colour, stroke width and opacity)
- [ ] Application order is fixed and does not depend on the order the user worked in: `normaliseStack` sorts into `straighten → crop → rotate → flip → resize → adjust → redact → annotate`, collapsing repeats of the single-valued kinds (a second `crop` replaces the first; a second `adjust` merges channel by channel) while `redact` and `annotate` accumulate in insertion order — so re-opening an edited image shows the same picture the editor showed, and the history list is a record of intent while the stack is a canonical program
- [ ] `src/lib/core/imageAdjust.ts` states the pixel maths rather than delegating to a CSS filter string, which varies between engines: `brightness` (−100..100 → an additive term of `v * 2.55`), `contrast` (−100..100 → `f = (259 * (c + 255)) / (255 * (259 - c))` applied about 128), `saturation` (−100..100, interpolating each channel towards the Rec. 709 luma `0.2126 R + 0.7152 G + 0.0722 B`), `exposure` (−100..100 → a multiplier of `2 ** (v / 50)`), `grayscale` (the same luma) and `invert`, each a pure function over a `Uint8ClampedArray` in that order, with alpha untouched and one unit test per channel against hand-computed values at 0, ±50 and ±100
- [ ] `/media/:id` gains an edit mode (`media-edit`, entered from an "Edit" action and left by Done or Escape) with a tool rail (`tool-crop`, `tool-rotate`, `tool-adjust`, `tool-redact`, `tool-annotate`), a contextual options bar per tool, a `media-edit-history` list of applied edits with per-row remove, and `edit-revert` ("Revert to original", clearing the stack after a confirm that names the edits it will discard)
- [ ] Every tool is fully keyboard-operable and screen-reader-labelled: crop handles are focusable and move by 1 px with arrows and 10 px with `Shift`, the crop rectangle is announced as `x, y, width × height`, annotation objects are a focusable list with arrow-key nudging and `Delete` to remove, sliders are native `<input type="range">` with number inputs beside them, and rotation is also reachable as `Cmd/Ctrl+[` / `Cmd/Ctrl+]`
- [ ] Undo and redo cover the whole edit session with SPEC-015's vocabulary and shortcuts: a per-view stack capped at `IMAGE_UNDO_LIMIT = 100`, `record(label)` **before** each mutation then `commit()` (the SPEC-080 order), one entry per completed gesture (a drag is one entry, not sixty), and labels that read as the history list does ("Crop", "Adjust brightness", "Add arrow")
- [ ] The preview is the renderer at a different size, as SPEC-109 requires of composition: `src/lib/media/renderEdits.ts` exports `renderEditedImage(source, edits, targetPx): Promise<ImageBitmapLike>`, called by the on-screen preview at screen scale and by every export at full resolution, so nothing can be true of one and false of the other; adjustments preview through the same pixel functions on a downscaled buffer, never through a CSS filter
- [ ] Saving is explicit about what it touches: `edit-save` ("Save") stores the stack on the record and refreshes `renderMediaId`, leaving `sourceMediaId` and its bytes untouched; `edit-save-copy` ("Save as a copy") creates a new SPEC-107 media document carrying the same source and stack; `edit-export` opens SPEC-108's format controls to download a flattened PNG or JPEG; each reports failure without mutating anything, and each is a single write-through
- [ ] **Redaction is labelled for what it is.** Whenever the stack contains a `redact`, the editor shows a persistent notice — "The original is still stored. Choose Flatten and discard original to remove the hidden pixels." — and `edit-flatten` performs exactly that: render at full resolution, replace `sourceMediaId`'s part with the rendered bytes, clear `edits`, and confirm through a dialog that names the operation as permanent and not undoable; downloading or exporting an image with redactions produces flattened pixels, so the risk is only ever in what remains stored here
- [ ] An SVG source is rasterised once, on the first edit, at its natural size (or a size the user picks when the SVG declares none), through `core/svgToPng.ts`; the original SVG part is kept as `sourceMediaId` and the raster becomes the stack's base, with a one-line notice saying editing is raster and the vector original is preserved
- [ ] Editing stays responsive on a large image: the preview works from a downscaled buffer capped at `EDIT_PREVIEW_MAX_EDGE_PX = 2048`, full-resolution work happens only on save, copy, export or flatten, that work is chunked with `yieldToMainThread()` between stages with `edit-progress` reporting `role="status"`, and it is refused up front against SPEC-109's probed canvas ceiling with a message naming the limit
- [ ] Determinism and purity: `normaliseStack` and `resolveEditGeometry` are pure and total; the same stack over the same source produces the same geometry to the pixel; `imageAdjust`'s functions are asserted against hand-computed values rather than snapshots; rendered **bytes** are not asserted byte-identical (canvas encoders differ across platforms — SPEC-109's rule, restated) but rendered pixels are, by sampling; and no edit path mutates the source `MediaPart`, asserted against a deep-frozen model
- [ ] Tests: `src/lib/core/imageEdits.spec.ts` — normalisation order and collapsing for every kind pair, geometry through crop + straighten + rotate + flip + resize in combination (exact output sizes and source rectangles), the empty stack as identity, and an over-long stack bound; `src/lib/core/imageAdjust.spec.ts` — one table-driven test per channel plus the round trip `invert(invert(x)) === x`; `src/views/media/useImageEditor.spec.ts` — undo granularity per gesture, the redaction notice's presence, flatten's confirmation and its irreversibility, no mutation of a frozen source; e2e (`e2e/SPEC-112.spec.ts`): open `spec107-photo.jpg` → Edit → crop to 16:9 → rotate right → brightness +30 → Save → the library thumbnail updates and the details panel reports the new pixel size → reopen → the history lists three edits and the crop is still adjustable → Revert to original → the pixel size returns to the source's → redact a rectangle → the notice appears → Flatten and discard original → the history is empty, the size matches the redacted render, and Undo does not bring the original back

## Out of scope
Layers, masks, selections beyond a rectangle, the healing and clone tools,
curves, levels, colour balance, white balance, sharpening, blurring, noise
reduction and every other tool an image editor has that a document product
does not need. Perspective correction and automatic deskew. Filters and
presets. Text on a path, rich text inside an annotation (annotation text is a
single font, size and colour). Vector editing of an SVG. Colour management,
ICC profiles and any non-sRGB working space. EXIF editing, and EXIF-based
auto-orientation (an image displays as the browser displays it — SPEC-107's
exclusion carries). Animated GIF frame editing (an animated GIF is editable
only as its first frame, and the editor says so before it starts). Editing a
PDF page, which needs the renderer SPEC-107 explains does not exist here.
Batch editing across several images. Content-aware anything.

## Reference

### Dependencies restated
- `specs/107-…` — `MediaModel { kind, id, title, format, contentType,
  sourceMediaId, media: MediaPart[], byteLength, image?, pdf?, properties? }`,
  the payload as a `MediaPart { mediaId, contentType, data }` split into its
  own IndexedDB record on a reference-identity latch, `/media/:id` and
  `MediaView.vue` with its viewport composable, `decodeImageDimensions`, and
  the library thumbnail cache keyed by id + `lastModified`.
- `specs/015-undo-and-redo.md` — the undo vocabulary, its shortcuts and the
  rule that one user gesture is one entry; `specs/080-…`'s `record` → mutate →
  `commit()` ordering is the pattern the per-view stack follows.
- `specs/108-…` — the export format controls, `SHEET_DPI_CHOICES`, and
  `encodeImage`'s `maxDpi`; `specs/109-…` — `probeMaxCanvasPixels()`, the
  preview-is-the-renderer rule and the pixel-sampling test convention.
- `src/lib/core/svgToPng.ts`, `src/lib/core/idleScheduler.ts`, and
  `src/lib/zoom.ts`'s clamp/step vocabulary, reused rather than restated.

### The edit stack
```ts
// src/lib/core/imageEdits.ts
export type ImageEdit =
  | { kind: "straighten"; degrees: number }                       // -45..45
  | { kind: "crop"; x: number; y: number; width: number; height: number }
  | { kind: "rotate"; quarterTurns: 0 | 1 | 2 | 3 }
  | { kind: "flip"; axis: "horizontal" | "vertical" }
  | { kind: "resize"; widthPx: number; heightPx: number; resample: "nearest" | "bilinear" }
  | { kind: "adjust"; brightness?: number; contrast?: number; saturation?: number;
      exposure?: number; grayscale?: boolean; invert?: boolean }
  | { kind: "redact"; x: number; y: number; width: number; height: number; color: string }
  | { kind: "annotate"; shape: AnnotationShape };

export interface EditGeometry {
  /** The output size after every geometric edit. */
  widthPx: number; heightPx: number;
  /** The source rectangle the output samples, before rotation/flip. */
  source: { x: number; y: number; width: number; height: number };
  quarterTurns: 0 | 1 | 2 | 3;
  flipH: boolean; flipV: boolean;
  straightenDegrees: number;
}

export const MAX_EDIT_STACK = 500;
export function normaliseStack(edits: ImageEdit[]): ImageEdit[];
export function resolveEditGeometry(edits: ImageEdit[], naturalPx: { widthPx: number; heightPx: number }): EditGeometry;
```
Crop coordinates are in **source** pixels after straightening, so adjusting a
straighten angle keeps a crop meaningful rather than re-basing it. Straighten
rotates about the image centre and the crop rectangle is clamped into the
inscribed axis-aligned area, which is why the two are one tool in the UI.

### Why non-destructive
Three reasons, in the order they matter here: a crop the user regrets is a
lost photograph and this library is the only copy; re-encoding a JPEG on every
save compounds loss across a session; and SPEC-110's passthrough rule needs to
know whether a source is still the original file, which a stack answers
exactly (`edits` empty ⇒ eligible). The cost is storage — a source plus a
render — and it is bounded by keeping exactly one render, regenerated rather
than versioned, and evictable under quota pressure like a thumbnail.

### Tests
- `src/lib/core/imageEdits.spec.ts` and `imageAdjust.spec.ts` as in the
  acceptance criteria — both pure, both fast, both hand-computed.
- `src/lib/media/renderEdits.spec.ts` — the source-rectangle and output-size
  arithmetic handed to the canvas (pure, from `resolveEditGeometry`), the
  single-decode cache, and the SVG rasterisation base.
- `src/views/media/useImageEditor.spec.ts` — the undo/redo, notice, flatten
  and no-mutation cases; `MediaView.edit.spec.ts` — tool switching, keyboard
  operation of the crop rectangle, the history list.
- `e2e/SPEC-112.spec.ts` as in the acceptance criteria.

### Specs this changes
- `specs/107-…` — `MediaModel` gains `edits?` and `renderMediaId?`; every
  consumer of "the current image" reads the render when one exists. The
  round-trip rule is unchanged and gains a clause: an **unedited** media
  document still downloads byte-identically, and an edited one downloads its
  flattened render, which the download dialog states.
- `specs/109-…`/`specs/110-…` — an edited source composes and assembles from
  its render; SPEC-110's passthrough table already has the row.
- `specs/000-roadmap.md` gains this spec's phase row.
