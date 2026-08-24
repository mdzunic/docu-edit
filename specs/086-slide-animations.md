---
id: SPEC-086
title: Slide animations — entrance and exit builds
status: draft
area: formatting
depends_on: [SPEC-082, SPEC-084, SPEC-085]
---
## Why
Transitions move between slides; animations move the things *on* a slide
— the bullet list that appears point by point, the diagram that builds
up on clicks. Both formats encode this in their hardest corners
(PresentationML's `p:timing` SMIL tree, ODF's `anim:*` tree), which is
exactly why a deliberate subset with a canonical serialisation is the
right shape: seven effects with clean equivalents in PowerPoint and
Impress, ordered builds with click/with-previous/after-previous
triggers, everything else preserved verbatim until the user opts to
replace it. SPEC-084's step machine already reserved build steps;
SPEC-088 replays these builds in exported HTML. **This file is the
normative source for the animation model, the canonical `p:timing`/
`anim:*` forms, and build playback semantics.**

## Acceptance criteria
- [ ] `Slide` gains `animations?: SlideAnimation[]` per the Reference: ordered effects targeting frames by frame id — effect appear/fadeIn/flyIn/wipeIn/zoomIn/fadeOut/flyOut, direction for fly/wipe, trigger click/withPrevious/afterPrevious, duration 100–5000ms (default 500), delay 0–10000ms (default 0)
- [ ] Both readers recognise their format's canonical form plus the tolerant subset described in the Reference (PowerPoint/Impress-authored equivalents); a slide whose timing contains anything unrecognised models **no** animations and keeps the source verbatim in `preserved` — mixed decks never half-parse
- [ ] The toolbar gains an "Animate" toggle opening an Animations side panel for the current slide: the ordered effect list (each row: order number, target frame name, effect, trigger glyph), controls for the selected row (effect, direction, trigger, duration, delay), Add (targets the selected frame; disabled with helper text when none), Remove, Move up/down, and Play (previews the slide's builds on the editor stage); all keyboard-operable, testids `slides-animate-toggle`, `animation-list`, `animation-add`, `animation-remove`, `animation-move-up`, `animation-move-down`, `animation-effect`, `animation-direction`, `animation-trigger`, `animation-duration`, `animation-delay`, `animation-play`
- [ ] Selecting an animation row highlights its target frame on the stage; a numbered badge renders beside each animated frame while the panel is open (click badge → select row)
- [ ] Adding/editing/reordering/removing dirties the slide, records one undo entry per gesture, persists write-through; adding to a slide with unrecognised preserved timing first asks "Replace the original animations this editor can't show?" (cancel = no change), and confirming drops the preserved timing for that slide
- [ ] Deleting a frame (SPEC-082) removes animations targeting it; undo restores both; duplicating a slide keeps its animations retargeted to the copied frames' fresh ids
- [ ] Playback: builds partition the slide's animations at each `click` trigger; frames targeted by a not-yet-played entrance effect are hidden from the moment the slide appears; advance plays the next build group — each effect animating per its definition, `withPrevious` simultaneous with its predecessor, `afterPrevious` chaining after (plus delays); exit effects hide their targets; back-navigation lands fully-built (SPEC-084's rule); reduced motion plays every effect as an instant show/hide; input mid-build completes the group instantly
- [ ] The editor stage (outside playback) always renders frames fully visible — animations affect playback and preview only
- [ ] Round trip: a deck given each effect (with a non-default trigger mix) → save → reopen reproduces the model exactly in both formats (canonical read-back); untouched slides' timing bytes stay verbatim; a canonically-written file re-saves byte-stable for those parts
- [ ] Fixture coverage: both mini decks gain a slide carrying the canonical forms, plus one slide with a deliberately exotic timing tree (pinning the preserve-don't-parse path)
- [ ] e2e (`e2e/SPEC-086.spec.ts`): open fixture → add "Fade in, on click" to a body frame and "Fly in from left, after previous" to a shape → Present → slide shows with both frames hidden → advance: first frame fades in, then the shape flies in without further input → advance again moves to the next slide → exit, download, reopen → panel lists both effects in order

## Out of scope
Emphasis effects, motion paths, per-paragraph text builds ("by
paragraph" — frame-level only), repeat/rewind/auto-reverse, triggers
bound to clicking a specific shape, media playback effects, timing
rehearsal. Reading arbitrary foreign timing into the model (preserve or
replace only). Animation on group/foreign frames (targets must be
text/image/shape frames; the Add control disables otherwise).

## Reference

### Dependencies restated
- SPEC-082 frame identity (`SlideFrame.id`, minted, never serialised —
  this spec adds the wire-level id minting below), frame names for the
  panel (catalogue names / "Text box" / "Image").
- SPEC-084 `usePlayback` builds (`PlaybackStep.buildIndex`, advance
  semantics, reduced-motion flag); SPEC-085's panel chrome pattern,
  preserved-vs-modelled discipline and confirm-to-replace precedent.
- SPEC-081 serialisers and the `preserved["p:timing"]` / odp `anim`
  slots.

### The model
```ts
export type AnimationEffect =
  | "appear" | "fadeIn" | "flyIn" | "wipeIn" | "zoomIn"
  | "fadeOut" | "flyOut";
export interface SlideAnimation {
  frameId: string;
  effect: AnimationEffect;
  /** flyIn/flyOut/wipeIn only; default "left". "left" means the motion
   *  originates at the LEFT edge: fly-in enters from the left, fly-out
   *  exits toward the left, wipe reveals starting at the left. */
  direction?: "left" | "right" | "up" | "down";
  trigger: "click" | "withPrevious" | "afterPrevious";
  durationMs?: number;  // default 500
  delayMs?: number;     // default 0
}
```
Order in the array is play order. The first animation of a slide always
behaves as `click` even if marked otherwise (both formats' sequencers
require an initiating click; the panel normalises row 1's trigger).

### Wire-level frame ids
Canonical timing must reference shapes. pptx: every rebuilt slide's
shapes get deterministic `p:cNvPr id` values (2, 3, … in z-order —
SPEC-081 already mints these); `p:spTgt spid` uses them, and the reader
resolves targets back through them. odp: rebuilt frames carry
`xml:id="de-f{n}"` (and matching `draw:id`), referenced by
`smil:targetElement`; the reader resolves either attribute. Reading a
file whose timing targets a shape the slide doesn't contain →
unrecognised (preserve).

### Canonical `p:timing` (write; read tolerantly)
```
<p:timing><p:tnLst><p:par>
 <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
  <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
   — one per click group —
   <p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
    <p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
      — one per effect in the group —
      <p:par><p:cTn presetID={id} presetClass={entr|exit} presetSubtype={sub}
                    fill="hold" nodeType={clickEffect|withEffect|afterEffect}
                    stCondLst delay={delayMs or accumulated for afterPrevious}>
        <p:childTnLst>{behaviours}</p:childTnLst></p:cTn></p:par>
    </p:childTnLst></p:cTn></p:par>
   </p:childTnLst></p:cTn></p:par>
  </p:childTnLst></p:cTn>
  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
 </p:seq>
 </p:childTnLst></p:cTn>
</p:par></p:tnLst></p:timing>
```
Preset table, as (presetID, presetSubtype) with presetClass `entr` for
the In effects and `exit` for Out: appear (1, 0); flyIn/flyOut (2, dir:
fromLeft 8, fromRight 2, fromTop 1, fromBottom 4); fadeIn/fadeOut
(10, 0); wipeIn (22, same direction codes as fly); zoomIn (23, 16). Behaviours per
effect: appear = `p:set` (`style.visibility` → `visible`, delay 0, dur 1);
fade = `p:animEffect transition="in|out" filter="fade"` +
`p:set` visibility (in: before; out: after, → `hidden`); fly =
`p:anim calcmode="lin"` on `ppt_x`/`ppt_y` from the off-slide edge
expression (`#ppt_x-1` style formulas are NOT used — write plain
from/to percentages: e.g. fromLeft: `p:anim attrName="ppt_x"
from="0-#ppt_w/2" …` is Office's style, but the canonical form uses the
simpler `<p:anim … by>`-free from/to pair PowerPoint accepts:
from `({-0.5},#ppt_y)`-style is over-clever — **write the standard
preset markup PowerPoint itself emits for these presets**, captured
once into golden fixtures; the goldens in
`presentation/pptx/timingGolden.ts` are the normative bytes, and this
spec's table names which preset each golden encodes); wipe =
`p:animEffect filter="wipe(fromLeft)"` etc.; zoom = `p:animEffect
filter="fade"` + `p:anim` on `ppt_w`/`ppt_h` 0→100%. Every behaviour
carries `p:cBhvr/p:cTn dur={durationMs}` and `p:tgtEl/p:spTgt spid`.
Reading: walk the mainSeq; for each effect-level `p:cTn` accept any
tree whose (presetClass, presetID) is in the table, taking subtype,
nodeType, delay, first behaviour `dur`, and spTgt — ignore behaviour
internals. Anything else in the timing (extra sequences, unknown
presets, `p:audio`…) → unrecognised → preserve whole.

### Canonical odp `anim:*` (write; read tolerantly)
Inside `draw:page` after content (the slot SPEC-081 reserves):
```
<anim:par presentation:node-type="timing-root">
 <anim:seq presentation:node-type="main-sequence">
  — per click group —
  <anim:par smil:begin="next">
   <anim:par smil:begin="0s">
    — per effect —
    <anim:par presentation:node-type="on-click|with-previous|after-previous"
              presentation:preset-class="entrance|exit"
              presentation:preset-id={ooo id} smil:begin={delay}
              [presentation:preset-sub-type={sub}]>
      {behaviours}
    </anim:par>
   </anim:par>
  </anim:par>
 </anim:seq>
</anim:par>
```
ooo preset ids: entrance `ooo-entrance-appear`, `ooo-entrance-fade-in`,
`ooo-entrance-fly-in` (sub `from-left|from-right|from-top|from-bottom`),
`ooo-entrance-wipe` (same subs), `ooo-entrance-zoom`; exit
`ooo-exit-fade-out`, `ooo-exit-fly-out` (same subs). Behaviours: appear
= `anim:set smil:attributeName="visibility" smil:to="visible"`; fades =
`anim:transitionFilter smil:type="fade"` (+ visibility set for exit);
fly = `anim:animate` on x/y with from/to expressions in the
`width`/`height`-relative form Impress emits (golden-fixture normative,
`presentation/odp/animGolden.ts`); wipe = `anim:transitionFilter
smil:type="barWipe"` with subtype/direction per the SPEC-085 table; zoom
= `anim:animateTransform smil:type="scale"`. Durations `smil:dur="0.5s"`
form. Reading mirrors pptx: accept preset-class/preset-id in the table,
else preserve whole.

### Golden-first implementation note (normative process)
The exact behaviour markup for fly/wipe/zoom is easy to get subtly wrong
from memory; the goldens are therefore checked in as the normative
canonical forms, unit tests assert serialiser output equals them, and
the readers must accept them. They are hand-authored during
implementation against ECMA-376 §19.5 / ODF 1.2 §16.29 + the mapping
above — not screen-scraped — and reviewed against the ACs' behavioural
round trip. Cross-app verification (opening in PowerPoint/Impress) is
manual QA outside CI, recorded in the PR description.

### Playback
`usePlayback` populates builds: group the slide's animations at click
boundaries → `builds: SlideAnimation[][]`; `buildIndex` advances through
them before the slide advances (SPEC-084's reserved semantics). The play
view applies effects as CSS transitions/keyframes on the frame elements
(visibility + transform + opacity + `clip-path` for wipe), honouring
duration/delay/chaining; `afterPrevious` chains via `transitionend`
with timeout fallback (SPEC-085's pattern). Editor preview ("Play")
runs the same engine on the editor stage in a non-interactive overlay.

### Tests
- `presentation/animations.spec.ts` (new lib module
  `presentation/animations.ts`: grouping into builds, retargeting on
  duplicate, removal on frame delete, trigger normalisation).
- Serialiser specs: golden equality both formats; id minting; preserve
  path untouched-byte assertions; replace-after-confirm drops preserved.
- Round trip per AC; reader tolerance (goldens + a PowerPoint-flavoured
  variant with reordered attributes).
- `usePlayback.spec.ts`: hidden-until-built, group playback order,
  instant-complete, reduced motion, back-lands-built.
- View + e2e per AC.

### Specs this changes
- SPEC-084's `PlaybackStep.buildIndex` goes live (was always 0).
- SPEC-085's panel/confirm patterns are shared; its "Animate" neighbour
  toggle joins the toolbar overflow rules from SPEC-081.
- SPEC-088 exports builds with these semantics.
