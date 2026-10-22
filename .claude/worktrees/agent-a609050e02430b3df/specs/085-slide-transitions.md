---
id: SPEC-085
title: Slide transitions
status: draft
area: formatting
depends_on: [SPEC-081, SPEC-084]
---
## Why
Decks arrive with transitions and presenters expect to set them; today
they survive only as preserved bytes and playback hard-cuts. This spec
models a deliberate subset — the six transitions that exist in both
PowerPoint and Impress with a clean mapping — with duration and advance
behaviour, edits them in a side panel, renders them in SPEC-084 playback,
and round-trips them through both formats. Everything outside the subset
stays preserved verbatim until the user actively changes that slide's
transition. SPEC-088 exports the same model to HTML, so **this file is
the normative source for the transition model and its two wire
mappings**.

## Acceptance criteria
- [ ] `Slide` gains `transition?: SlideTransition` per the Reference: type none/fade/push/wipe/cover/uncover/split, optional direction, duration 100–5000ms (default 600), advance on click (default true) and/or auto-advance after N ms
- [ ] Both readers populate it from the subset mappings; a transition outside the subset leaves the model field unset and the source XML in `preserved` ("original transition kept"), and the panel shows that state honestly
- [ ] The toolbar gains a "Transition" toggle opening a side panel for the current slide: effect list (None + the six, with the direction control enabled per effect), duration input (seconds, 0.1 steps), "Advance on click" checkbox, "Auto-advance after N s" checkbox + input, a Preview button, and "Apply to all slides"; every control keyboard-operable, testids `slides-transition-toggle`, `transition-effect`, `transition-direction`, `transition-duration`, `transition-advance-click`, `transition-advance-auto`, `transition-preview`, `transition-apply-all`
- [ ] Choosing an effect (or changing any field) sets the model, dirties the slide, records one undo entry, persists, and replaces any preserved out-of-subset transition for that slide after an inline confirmation ("Replace the original transition this editor can't show?") — cancelling changes nothing
- [ ] A slide whose transition was replaced or set serialises it in both formats per the wire mappings; a slide never touched keeps its original transition bytes verbatim (including out-of-subset ones), even when the slide dirties for text reasons — `preserved["p:transition"]`/the page's smil attributes are re-emitted unless this spec's commands ran on that slide
- [ ] Playback renders the *incoming* slide's transition on every forward step between slides: fade (cross-dissolve), push (both slides travel), wipe (reveal via clip), cover (incoming slides over), uncover (outgoing slides away), split (barn-door reveal), each honouring direction and duration; backward steps and `goTo` are instant; input during a transition completes it instantly and then applies (no queued double-advance)
- [ ] `prefers-reduced-motion` (SPEC-084 plumbing) renders every transition as an instant cut; durations are otherwise real time (a 2s fade takes ~2s)
- [ ] Auto-advance: in playback, a slide with `advanceAfterMs` schedules an advance that cleans up on manual navigation/exit; with "advance on click" false, click/keys that mean "advance" still move (presenter escape hatch — the flag is honoured only by SPEC-088's export viewer in kiosk-like reading, and this is documented in the panel's helper text: "Clicks always advance while presenting here")
- [ ] Preview plays the transition once on the editor stage from the previous slide (or from black for slide 1) without entering playback
- [ ] Round trip: set each of the six (with a non-default direction and duration and an auto-advance) → save → reopen preserves the model in both formats (pptx duration quantisation per the mapping is the one sanctioned delta and the test pins it); Impress/PowerPoint-authored fixture markup for each mapped effect parses to the expected model
- [ ] The fixture decks gain per-effect transition slides in both formats (extend `presentation/fixtures/build.ts`)
- [ ] e2e (`e2e/SPEC-085.spec.ts`): open fixture → set Push/Left/1s on slide 1 with apply-to-none → Present → advancing to slide 2 shows movement (assert an intermediate transform/opacity state exists, then final state) → exit → download → reopen → panel shows Push/Left/1s

## Out of scope
Every other transition (morph, zoom, random bars…) — preserved verbatim,
never editable beyond replacement. Per-slide transition *sound*.
Transition on the first slide's entrance. Distinct backward-navigation
effects. Editing `advanceAfterMs` behaviour differences between click
and keyboard. p14 (PowerPoint 2010+) extension writing — see the wire
notes: we read `p14:dur` when present, we write `spd` only.

## Reference

### Dependencies restated
- SPEC-079 `preserved` capture; SPEC-080/081 dirty + rebuild rules and
  the serialisers this extends; SPEC-080 undo; SPEC-084's advance
  pipeline (`usePlayback`), reduced-motion flag, and stage-in-play view.
- Panel chrome: the SPEC-065 side-panel conventions as applied to the
  presentation view (a plain right-hand panel; the unified side panel
  is a text-product surface — presentations get a sibling styled by the
  same tokens, not an import of it).

### The model
```ts
export type TransitionType =
  | "fade" | "push" | "wipe" | "cover" | "uncover" | "split";
export type TransitionDirection = "left" | "right" | "up" | "down";

export interface SlideTransition {
  type: TransitionType | "none";       // "none" = explicit hard cut
  /** Movement/reveal direction of the incoming content; defaults per
   *  type: push→left, wipe→left, cover→left, uncover→left; fade/split
   *  ignore it (split is always horizontal-out). */
  direction?: TransitionDirection;
  durationMs?: number;                 // 100..5000, absent = 600
  advanceOnClick?: false;              // absent = true
  advanceAfterMs?: number;             // absent = no auto-advance
}
```
`transition` absent + no preserved transition XML = inherit nothing:
hard cut (the formats' own default). Direction semantics: "left" means
the incoming slide travels leftward (enters from the right edge) —
matching PowerPoint's `dir="l"` for push.

### pptx wire mapping (read and write)
`p:transition` child of `p:sld` (after `p:cSld`/`p:clrMapOvr`, the slot
SPEC-081's serialiser already reserves):
| model | element | attrs |
| --- | --- | --- |
| fade | `p:fade` | — |
| push dir | `p:push` | `dir` = `l/r/u/d` |
| wipe dir | `p:wipe` | `dir` = `l/r/u/d` |
| cover dir | `p:cover` | `dir` = `l/r/u/d` |
| uncover dir | `p:pull` | `dir` = `l/r/u/d` |
| split | `p:split` | `orient="horz" dir="out"` |
Reading also accepts `p:split` with any orient/dir (→ split) and maps
`p:cut`/absent child element to "none". Duration: read `spd`
(`slow`=1000, `med`=600, `fast`=300 when no better source) and an
extension-list `p14:dur` (milliseconds) when present anywhere inside the
element (string scan is acceptable; namespace-correct DOM read
preferred). Write `spd` only: nearest of `fast` (≤450), `med` (≤800),
`slow` — and the model keeps exact ms until the file is reopened
(quantisation on reopen is the pinned sanctioned delta). `advClick`
(default `1`) and `advTm` (ms) read/write directly on `p:transition`.

### odp wire mapping (read and write)
Transitions live on the page's drawing-page automatic style:
`draw:page draw:style-name` → `style:style style:family="drawing-page"`
→ `style:drawing-page-properties`. Read `smil:type`/`smil:subtype`/
`smil:direction`/`smil:dur` (e.g. `PT0.6S`) and
`presentation:transition-speed` as the duration fallback (slow/medium/
fast → 1000/600/300); `presentation:duration` (`PT5S` style) +
`presentation:transition-type="automatic"` → `advanceAfterMs`, and
`…="manual"` → click-only.
| model | smil:type | smil:subtype | reverse |
| --- | --- | --- | --- |
| fade | `fade` | `crossfade` | — |
| push left/right/up/down | `pushWipe` | `fromRight`/`fromLeft`/`fromBottom`/`fromTop` | — |
| wipe left/right | `barWipe` | `leftToRight` | `smil:direction="reverse"` for left |
| wipe up/down | `barWipe` | `topToBottom` | reverse for up |
| cover l/r/u/d | `slideWipe` | `fromRight`/`fromLeft`/`fromBottom`/`fromTop` | — |
| uncover l/r/u/d | `slideWipe` | same subtypes | `smil:direction="reverse"` |
| split | `barnDoorWipe` | `horizontal` | — (out) |
(Note the from-edge inversion: incoming-travels-left = enters
`fromRight`.) Unrecognised combinations → out-of-subset (preserved).
Writing: the SPEC-081 `StyleRegistry` mints a `de-DP{n}` drawing-page
style carrying the smil attributes + `presentation:transition-type` +
`presentation:duration`, and the rebuilt page's `draw:style-name` points
at it; the page's original drawing-page style is otherwise carried by
the untouched-page path. Fill/background properties of a replaced
drawing-page style are copied into the minted style so setting a
transition never changes the page background.

### Playback rendering
`usePlayback` gains a `transitionState` the play view renders as **two
stacked stages** (outgoing + incoming) inside the letterbox container;
CSS classes per type/direction drive `transform`/`opacity`/`clip-path`
with `transition-duration` set inline from the model; `transitionend`
(with a duration+100ms timeout fallback) commits the step. Split =
`clip-path: inset(0 50%)` → `inset(0 0)` on the incoming stage. Cover/
uncover move exactly one stage; push moves both. Preview reuses the same
two-stage component on the editor stage host.

### Tests
- `presentation/transitions.spec.ts` (new lib module
  `presentation/transitions.ts` holding the mapping tables both
  serialisers and SPEC-088 import): every row both directions,
  quantisation, defaults, out-of-subset detection.
- Serialiser/round-trip specs per AC (both formats, incl.
  background-copy rule for the minted odp style).
- `usePlayback.spec.ts` additions: transition step interleaving,
  instant-complete on input, auto-advance timer cleanup, reduced motion.
- View/e2e per AC. e2e asserts an intermediate state via
  `page.evaluate` sampling, not screenshots (memory: screenshot
  pixel-review limits).

### Specs this changes
- SPEC-081's serialiser: the reserved `p:transition` slot now emits
  either preserved XML or this spec's mapping (never both);
  the odp `draw:page` attribute handling gains the drawing-page style
  substitution above.
- SPEC-084: the advance pipeline's "slide changes are instant" note is
  superseded — forward steps play transitions; the *logical* step model
  is unchanged.
- SPEC-088 must export this model with matching visuals.
