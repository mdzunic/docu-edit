---
id: SPEC-095
title: Slide media — audio and video frames
status: draft
area: structure
complexity: normal
story_points: 8
depends_on: [SPEC-079, SPEC-082, SPEC-084, SPEC-086, SPEC-088]
---
## Why
SPEC-079 turns every video and sound clip on a slide into a grey dashed box labelled
"Media" (`foreignFrame(pic, base, "Media")`, `pptx/reader.ts:519`; `foreignAt(frame,
base, "Media")`, `odp/reader.ts:445`), and SPEC-088's out-of-scope list keeps them
placeholders in the HTML export. A product-launch deck is mostly a video; a
language-course deck is mostly audio. Opening one loses nothing today — the bytes ride
out in `opaque` — but shows nothing, and there is no way to add a clip at all.

This spec adds a real `MediaFrame` to the frame union, reads and writes it in both
formats, and plays it on the editing stage and in SPEC-084 playback, staying deliberately
conservative everywhere the browser is: we play exactly what the browser plays through a
plain `<video>`/`<audio>` element, never transcode, add no dependency, and degrade every
clip we cannot resolve, decode or afford to a labelled poster rather than an error.
**This file is the normative source for the media frame model, the pptx/odp media wire
mappings, the media size ceiling and browser-autoplay policy.**

## Acceptance criteria
- [ ] `SlideFrame` gains `MediaFrame` per the Reference: `kind: "media"`, `media: "video" | "audio"`, optional `mediaId` (key into `model.media`), `posterMediaId`, `linkTarget`, `contentType`, `autoStart`, `loop`, `muted`, `volume` (0–100, absent = 100), `fullScreen`, `oversize`, `stroke`, and `sourceXml` — the verbatim source element, re-emitted unchanged until a command touches this frame, exactly as `Slide.sourceXml` works one level up
- [ ] pptx reader: a `p:pic` whose `p:nvPicPr/p:nvPr` carries `a:videoFile` or `a:audioFile` (`r:embed` or `r:link`), or whose `p:extLst` carries the `p14:media` extension, becomes a `MediaFrame` with its `p:blipFill/a:blip r:embed` collected as `posterMediaId`; a `p:pic` with neither still becomes an `ImageFrame` or the SPEC-079 `ForeignFrame` `"Media"`; the reader sets no dirty flag and mutates no `opaque` part
- [ ] odp reader: a `draw:frame` containing `draw:plugin` (or a media-typed `draw:object`) becomes a `MediaFrame` — `xlink:href` through `normaliseHref` into `Pictures/…` or `media/…`, `draw:mime-type` → `contentType`, `presentation:play-full-screen` → `fullScreen`, `draw:param` rows mapped per the Reference, a sibling `draw:image` → `posterMediaId`; `anim:audio` stays outside SPEC-086's subset and keeps `preserved["anim"]` — this spec does not widen `readOdpAnim`
- [ ] `src/lib/core/media.ts` gains the product-neutral media vocabulary per the Reference (`ACCEPTED_MEDIA_MIME_TYPES`, `MEDIA_ACCEPT_ATTR`, `isAcceptedMediaType`, `MAX_MEDIA_SIZE_BYTES = 25 * 1024 * 1024`, `MEDIA_TOO_LARGE_MESSAGE`, the mp4/webm/ogg/m4a/mp3/wav/webp rows both lookup tables need), plus new `core/mediaPlayback.ts`, landed as the two-phase move the Reference describes with `core/boundary.spec.ts` and `presentation/boundary.spec.ts` still green
- [ ] The size ceiling is explicit: an inserted file over `MAX_MEDIA_SIZE_BYTES` is rejected with `MEDIA_TOO_LARGE_MESSAGE` and changes nothing, while a clip already inside an opened deck is never rejected — it models with `oversize: true` and no `mediaId`, renders the `media-oversize` chip "Not loaded — over 25 MB", and rides out byte-identically through `opaque`
- [ ] A clip's store cost is bounded by an explicit commit gate: `useSlideRender.commit()` re-slices `model.media` only when a media-mutating command has called its new `bumpMedia()`, so a 20 MB clip is not rewritten whole through `documentStore`'s `MEDIA_KEY_PREFIX + docId` record on every nudge and keystroke; **every** path that mutates `model.media` calls `bumpMedia()` before `commit()` — SPEC-082's `useFrameInsertion.insertImageFile` (which today leans on the unconditional re-slice and ends in a bare `deps.commit()`) as much as this spec's media insert and "Replace poster…" — and the new `useSlideRender.spec.ts` gates both, a just-inserted image and a just-inserted clip, each still rendering immediately
- [ ] Insert gains "Media…" (`insert-media` on `SlideInsertBar.vue`, picker `slides-media-input`, `accept` = `MEDIA_ACCEPT_ATTR`): video lands at a 16 cm × 9 cm box, audio at an 8 cm × 1.6 cm strip, both clamped into the slide and selected on landing; an unsupported type, an unreadable file or an over-ceiling file reports through `reportError` and changes nothing
- [ ] The stage renders a media frame as `slide-media-frame`: the poster (`media-poster`) when one exists, else a neutral `#1f1f1f` plate with a centred play glyph and the clip's basename, plus a `media-play` badge; nothing auto-plays in the editor; activating the badge (click, or Enter on the selected frame) swaps in a native `<video controls>`/`<audio controls>` (`media-controls`) that plays inline, and changing slides or deselecting stops it
- [ ] A Media panel (toolbar toggle `slides-media-toggle`, root `slides-media-panel`, `aria-label` `` `Media for slide ${slideNumber}` ``) edits the selected media frame: `media-start-mode` (on click vs automatically), `media-loop`, `media-mute`, `media-volume` (0–100), `media-full-screen`, "Replace poster…" (`media-poster-replace`, accepting `PRESENTATION_IMAGE_MIME_TYPES`); every control keyboard-operable, each edit recording one undo entry in the SPEC-080 voice ("Media start mode", "Media loop", "Media mute", "Media volume", "Media full screen", "Replace poster"), dirtying via `markSlideEdited` and persisting write-through
- [ ] Nothing degrades into an error: a linked clip (`r:link`, or an odp `xlink:href` resolving outside the package) models `linkTarget` without `mediaId` and renders `media-link-placeholder` captioned with the target's basename and "Linked file — not embedded"; an unplayable codec — `canPlayMedia(contentType, probe)` returning `false` from the element's own `canPlayType`, or a runtime `error` event — renders the poster plus `media-unsupported` "This browser can't play that format"; both keep open, edit, playback and round-trip working, with no transcoding and no new dependency
- [ ] Autoplay policy is specified, not hoped for: entering SPEC-084 playback is a user gesture, so an `autoStart` clip on the entered slide calls `play()`; a rejected promise (`NotAllowedError` — an unmuted policy, a slide reached by `advanceAfterMs`, or a deep-linked `/slides/:id/present/:n` load) degrades to the poster plus a `media-blocked` "Click to play" affordance that plays on the next input; `muted` clips are never blocked; leaving playback, navigating away, or `back()` past the slide pauses and resets every clip on it
- [ ] SPEC-086 animation targeting extends to media: `isAnimatableKind` (`src/views/presentation/useAnimations.ts:49`) gains `"media"`, a frame hidden by a pending entrance build does not play (an `autoStart` clip starts when its entrance build finishes, after `delayMs + durationMs`), an exit effect pauses and hides its target, and a clip's own duration is never a build step — advance never waits for playback to end
- [ ] Round trip, naming its relations: a deck carrying media opens, saves and downloads with **every** part `equalBytes`-identical (`src/lib/core/equality.ts`, compared part-level after `unzip`, never container bytes); a slide dirtied elsewhere re-emits each untouched media frame's `sourceXml` as an exact substring of the rebuilt part; insert → save → reopen compares the modelled media fields with `equalStructural` (`src/lib/equality.ts`); `JSON.parse(JSON.stringify(model))` still `toEqual(model)`
- [ ] SPEC-088 export gets a stated media rule instead of a 60 MB file: media embed as `data:` URIs only while the deck's total media base64 stays within `HTML_EXPORT_MEDIA_BUDGET_BYTES = 8 * 1024 * 1024`; past it every media frame exports as its poster plate captioned "Media omitted from this export", the dialog shows the count and reason before confirming (`export-html-media-note`), and the emitted bytes contain no media payload
- [ ] Fixture coverage: `fixtures/build.ts` gains `buildPptxMediaFixtureBytes` / `buildOdpMediaFixtureBytes` + `EXPECTED_MEDIA_FRAMES`, `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec095-media.pptx` and `.odp`, and each deck carries an embedded audio clip, an embedded video clip, a linked clip and one deliberately out-of-subset media element (pinning preserve-don't-parse), per the Reference's note on which payloads are synthesisable in-repo
- [ ] `PresentationView.vue` gains one `useSlideMedia({ … })` wiring call plus the `<MediaPanel>` mount and toolbar prop pair; the `onMounted`/`onBeforeUnmount` window-listener block moves to `src/views/presentation/useWindowKeys.ts` (or an equivalent extraction where a sibling has already landed that one) and `PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 10 from whatever it stands at when this lands (550 → 540 today), with a ledger row naming this spec and its extraction appended — the budget never rises
- [ ] e2e (`e2e/SPEC-095.spec.ts`): open `spec095-media.pptx` → the audio frame shows its poster plate, not a grey "Media" box → open the Media panel, set "Start automatically" and Loop → Present → the audio frame reports playing, or shows `media-blocked` (asserted as one of the two) → advance past it → exit → download, reopen → the panel still reads Start automatically + Loop, the linked clip still shows `media-link-placeholder`, and the video frame shows `media-unsupported` (deliberate non-assertion: no real codec runs in CI)

## Out of scope
Trimming, fade in/out, start/end offsets and bookmarks (`p:cmd` `playFrom(…)` is read
only far enough to set `autoStart`, then preserved); recording, screen capture,
transcoding, extracting a poster from the video's own frames; subtitles/captions, alt
text (SPEC-079 has no `descr` field on any frame), volume envelopes, "play across
slides" and background soundtracks; authoring pptx `p:audio` or ODF `anim:audio` timing
nodes (SPEC-086 owns the timing tree and keeps both preserved); a media library;
streaming or remote URLs (a linked file is a placeholder, never fetched — SPEC-088's
self-containment rule); cross-format conversion (SPEC-080's `unconvertibleFrameCount`).

## Reference

### Dependencies restated
- **SPEC-079** (`specs/079-slide-canvas-rendering.md`) owns the slide model. From
  `src/lib/presentation/model.ts`: `SlideFrameBase` (not exported) is `{ id: string; xCm;
  yCm; widthCm; heightCm: number; rotationDeg?: number }`, `id` minted with
  `crypto.randomUUID()` at read and never serialised; `SlideFrame = TextFrame | ImageFrame
  | ShapeFrame | GroupFrame | ForeignFrame` — `MediaFrame` goes before `GroupFrame`;
  `ImageFrame` is `{ kind: "image"; mediaId: string; stroke?: FrameStroke }`, `FrameStroke`
  is `{ color: string; widthPt: number }`. Lengths cm, colours lowercase `#rrggbb`,
  payloads base64 — never `Uint8Array`. Untouchable slots: `PresentationModel.opaque`
  (source parts verbatim, in package order), `Slide.partName`, `Slide.sourceXml`,
  `Slide.preserved`.
- **SPEC-082** (`specs/082-slide-objects-images-shapes-text-boxes.md`) owns insertion:
  `frames.ts` exports `FrameRect { xCm; yCm; widthCm; heightCm }`,
  `insertImageFrame(model, slideIndex, media: MediaPart, rect): string | null`,
  `mediaDirectory(model)` (`"Pictures/"` for odp else `"ppt/media/"`),
  `mediaExtension(contentType)` and private `mintMediaId`, allocating
  `{dir}image{n}.{ext}` past the highest numeric suffix across `opaque` + `media`; its
  view-side `useFrameInsertion.insertImageFile` is the one existing `model.media` mutator,
  and the path this spec's commit gate must not regress. `MediaPart`
  (`src/lib/core/media.ts`) is `{ mediaId; contentType; data: string }`,
  `mediaDataUri(part)` → `` `data:${contentType};base64,${data}` ``; every command ends in
  `markSlideEdited(slide, format)` (`text.ts:340`) — `slide.dirty = true`, plus odp's
  `delete slide.sourceXml`.
- **SPEC-084** owns playback navigation: `usePlayback` → `{ current: Ref<PlaybackStep>,
  atEnd, reducedMotion, advance, back, goTo }`, `PlaybackStep { slideIndex; buildIndex }`
  — no step kind is added here, media never gates an advance. **SPEC-086** owns the
  animation model and the canonical `p:timing`/`anim:*` forms: `SlideAnimation { frameId;
  effect; direction?; trigger; durationMs?; delayMs? }`, `animations.ts` exporting
  `buildGroups`, `removeAnimationsForFrame` and `hasPreservedTiming`, with the
  module-private `isAnimatableKind` at `src/views/presentation/useAnimations.ts:49`
  (text/image/shape, widened here to `"media"`) and out-of-subset timing kept whole in
  `preserved["p:timing"]`/`["anim"]`. **SPEC-088** owns DOM-free
  `exportPresentationHtml(model, opts: { includeNotes: boolean }): string`
  (`htmlExport.ts`), its no-external-requests rule and `export-html-dialog`,
  `export-html-notes`, `export-html-confirm`.

### The model (delta to SPEC-079's union)
```ts
export type SlideMediaKind = "video" | "audio";

export interface MediaFrame extends SlideFrameBase {
  kind: "media";
  media: SlideMediaKind;
  /** Key into `model.media`. Absent is never "no media": it pairs with
   *  `linkTarget` or `oversize` — a payload we did not take into the model. */
  mediaId?: string;
  linkTarget?: string;    // source href/target of a clip we did not embed
  posterMediaId?: string; // poster still; absent = the neutral plate
  contentType?: string;   // from the package part or draw:mime-type
  autoStart?: true;       // absent = start on click
  loop?: true;
  muted?: true;
  volume?: number;        // 0..100; absent = 100
  fullScreen?: true;      // odp presentation:play-full-screen
  oversize?: true;        // in the package, over MAX_MEDIA_SIZE_BYTES
  stroke?: FrameStroke;
  /** The verbatim source `p:pic` / `draw:frame`. Present ⇒ the serialisers re-emit
   *  it unchanged; deleted by any command changing a modelled field
   *  (`markFrameEdited`, the frame-level analogue of `markSlideEdited`'s odp rule).
   *  Absent ⇒ rebuild from the fields above. */
  sourceXml?: string;
}
```
Under `equalStructural` (`src/lib/equality.ts`) `sourceXml` is **not** in
`XML_STRING_KEYS`, so it compares as a raw string — byte-exactness is the point; `id` is
dropped by `IGNORED_KEYS` while `mediaId`/`posterMediaId` are not. `equalSemantic`/
`resolve` are typed `(model: DocumentModel)`, so they stay unusable on a
`PresentationModel` and no criterion here claims them.

### Reuse and the two-phase core move
`src/lib/core/media.ts` already owns the product-neutral half of embedded media for the
word processor (`MediaPart`, `mediaDataUri`, `MAX_IMAGE_SIZE_BYTES`,
`ACCEPTED_IMAGE_MIME_TYPES`, `CONTENT_TYPE_BY_EXTENSION`, `extensionForContentType`,
`contentTypeForPartName`); media types belong there too. **Phase 1**, with its own tests
and no call-site changes, adds `ACCEPTED_MEDIA_MIME_TYPES = ["video/mp4", "video/webm",
"video/ogg", "audio/mpeg", "audio/mp4", "audio/ogg", "audio/wav"]`, `MEDIA_ACCEPT_ATTR`,
`isAcceptedMediaType`, `MAX_MEDIA_SIZE_BYTES`, `MEDIA_TOO_LARGE_MESSAGE`, the
mp4/webm/ogg/m4a/mp3/wav (plus webp) rows both lookup tables need —
`extensionForContentType` today falls back to `"png"`, handing a video a `.png` part path
— and new `core/mediaPlayback.ts` exporting `canPlayMedia(contentType: string, probe:
(type: string) => string): boolean`, pure with the `canPlayType` probe injected so it
stays DOM-free. **Phase 2** switches `frames.ts` and `useFrameInsertion.ts` over and
deletes `mediaExtension` and its webp case. Nothing in `core` imports `presentation/`.

### pptx wire mapping
Read: `p:pic` → media when `p:nvPicPr/p:nvPr` holds `a:videoFile` or `a:audioFile`, or
`p:extLst` holds the `p14:media` extension; `r:link`/`r:embed` resolve through the slide
rels, types `…/officeDocument/2006/relationships/video` and `…/relationships/audio`, the
p14 pointer using a Microsoft `…/2007/relationships/media` type; the poster is the
`p:blipFill/a:blip r:embed` image, collected as an `ImageFrame`'s is.
**Wire detail this spec does not pretend to know:** the `p:ext` URI GUID for `p14:media`,
the precise `…/2007/relationships/media` string, and the shape of the `p:timing` media
nodes (`p:cmd type="call" cmd="playFrom(0.0)"`, `p:audio`/`p:video` under `p:childTnLst`)
marking a clip auto-start. Capture them **golden-fixture-first**, exactly as SPEC-086 did:
hand-author `pptx/mediaGolden.ts` (`MEDIA_PIC_GOLDEN`, `MEDIA_TIMING_GOLDEN`,
`GOLDEN_MEDIA_FRAMES`) against ECMA-376 §19.3/§19.5, assert serialiser output equals the
goldens, make the reader accept them plus a PowerPoint-flavoured variant with reordered
attributes; cross-app verification is manual QA in the PR, not CI. Anything else in
`p:timing` stays whole per SPEC-086. Write: untouched → `sourceXml`; rebuilt → the
canonical `p:pic` with `xfrmXml(frame)` geometry, poster in `p:blipFill`,
`a:videoFile`/`a:audioFile` under `p:nvPr`. `referencedMedia` (`pptx/writer.ts:178`, and
its odp twin at `odp/writer.ts:31`) today walks only `frame.kind === "image"`; it **must**
also collect `MediaFrame.mediaId` and `posterMediaId`, or an inserted clip's part is
silently dropped from the saved package. `linkImageRels` (`:239`) gains a media sibling
minting video/audio/media rels; `planMedia`'s `<Default Extension=…>` splice covers the
new extensions.

### odp wire mapping
Read: `draw:frame` containing `draw:plugin` (Impress's embedded-media form) or a
media-typed `draw:object`; `xlink:href` through `normaliseHref` (strips `./` and a leading
`/`) into `Pictures/…` or `media/…`; `draw:mime-type` → `contentType`;
`presentation:play-full-screen` → `fullScreen`; a sibling `draw:image` → `posterMediaId`;
playback flags ride as `draw:param` `draw:name`/`draw:value` pairs (`Loop`, `Mute`,
`VolumeDB`, `AutoPlay`). **Wire detail this spec does not pretend to know:** the exact
`draw:param` names and value spellings Impress emits, and when a build writes
`draw:plugin` rather than `draw:object` — same golden-first process, `odp/mediaGolden.ts`
(`MEDIA_PLUGIN_GOLDEN`, re-exporting `GOLDEN_MEDIA_FRAMES`) hand-authored against ODF 1.2
§10.4.7/§14.9, reader tolerant of an Impress-flavoured variant, manual cross-app QA in the
PR. Namespaces come from `core/odfXml.ts` (`DRAW_NAMESPACE`, `PRESENTATION_NAMESPACE`,
`XLINK_NAMESPACE`); wire ids follow SPEC-086's `de-f{n}` on both `xml:id` and `draw:id`.
Write: untouched → `sourceXml`; rebuilt → the canonical `draw:frame` + `draw:plugin` with
`geomAttrs(frame)` and `xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"`, plus
the `spliceManifest` (`odp/writer.ts:50`) entry declaring each new part's
`manifest:media-type`; `writeOdp` never consults `notesDirty`.

### Store cost, and the view surface
`model.media` is base64 in a plain JSON model, persisted as **one** IndexedDB record
(`documentStore.ts`, `MEDIA_KEY_PREFIX + docId`) whose write latch skips only while the
`media` **array reference** is unchanged, and `commit()` re-slices it unconditionally
today — so a clip is rewritten whole on every gesture, and one over `openDocument`'s 50 MB
`MAX_FILE_SIZE_BYTES` could not have arrived at all: hence the ceiling and the
`bumpMedia()` gate. Undo snapshots are `structuredClone(model.slides)` and exclude
`media`, so a clip never enters the 100-entry stack; an undone insert leaves an orphan
`MediaPart` that `referencedMedia` filters out of the save.
`src/views/presentation/useSlideMedia.ts` takes `MediaDeps { model, currentSlide,
currentIndex, record, commit, selectedFrameId, selectFrame }` and returns the panel's
whole surface: `open`/`toggle`, `frame`/`isMedia`, a value+setter pair per modelled flag
(`startMode`, `loop`, `muted`, `volume`, `fullScreen`), `posterAccept`/`replacePoster`,
`insertMediaFile`, `playState`/`play`/`stop`/`blocked`/`unsupported`; `open` is the
persisted preference `presentationMediaOpen`, added to `src/lib/preferences.ts` in all
three places. `MediaPanel.vue` follows the Transitions/Animations panel contract exactly —
props `{ controller, slideNumber }`, emits nothing, `const c = props.controller;`, runtime
`defineProps({…})` only (espree rejects TS-only syntax in `.vue` scripts).
`SlideFrameView.vue` gains the media branch (`slide-media-frame`, `role="group"`,
"Video"/"Audio" via `frameNoun`); typed transforms go in `mediaRender.ts`.

### Tests
Fixture payloads are byte constants in a new `fixtures/mediaBytes.ts`, because
`fixtures/build.ts` is DOM-free: a minimal PCM **WAV** is constructible in code (44-byte
RIFF header plus samples) and every browser decodes it — the playable clip — while a real
encoded MP4 is not, so the video part is a ~512-byte synthetic `ftyp`/`isom` payload
exercising the unsupported-codec degrade path.
- `src/lib/core/media.spec.ts` (accept lists, ceiling, extension/content-type tables),
  new `core/mediaPlayback.spec.ts`, and `presentation/frames.spec.ts`
  (`insertMediaFrame` boxes and clamping, `mintMediaId`, `markFrameEdited`).
- New `mediaFrames.spec.ts` (format-neutral helpers) plus `pptx/media.spec.ts` and
  `odp/media.spec.ts` in SPEC-086's four-test golden shape: golden equality, `""` for
  nothing, golden → model, tolerant variant, out-of-subset preserved. Then
  `pptx/serialize.spec.ts` / `odp/serialize.spec.ts`: untouched-frame `sourceXml`
  substring identity, rels/manifest/Content-Types splices, `referencedMedia` collecting
  media + poster ids.
- New `mediaRoundTrip.spec.ts` per the round-trip AC (`equalBytes` per part,
  `equalStructural` on the modelled fields), nothing in `roundTrip.spec.ts` weakened;
  `htmlExport.spec.ts` gains the budget and no-payload assertions; `useSlideMedia.spec.ts`;
  new `useSlideRender.spec.ts` gating `bumpMedia()` on **both** an inserted image and an
  inserted clip; `PresentationView.size.spec.ts` at its ratcheted-down budget;
  `e2e/SPEC-095.spec.ts` per the last AC.

### Specs this changes
- SPEC-079's `ForeignFrame` `"Media"` label now covers only media we could not identify,
  closing its "no video/audio model" gap; the four mini-deck slides' expectations are
  unchanged — coverage is preserved, never weakened.
- SPEC-082's insert surface gains "Media…", `frames.ts#mediaExtension` folds into
  `core/media.ts`, and its `useFrameInsertion.insertImageFile` now calls `bumpMedia()`
  before `commit()` — behaviour unchanged, the re-render it relied on is requested
  rather than assumed. SPEC-086's `isAnimatableKind` gains `"media"` while its timing
  subset is unchanged and `anim:audio` / `p:audio` stay preserved.
- SPEC-088's out-of-scope line "embedding video/audio (foreign frames stay
  placeholders)" is superseded by the budget rule above; `specs/000-roadmap.md` gains
  this spec's phase row; nothing else changes.
