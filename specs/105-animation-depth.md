---
id: SPEC-105
title: Animation depth — paragraph builds, emphasis effects and motion paths
status: draft
area: formatting
complexity: normal
story_points: 8
depends_on: [SPEC-080, SPEC-081, SPEC-082, SPEC-084, SPEC-086, SPEC-088]
---
## Why
SPEC-086 animates a whole frame, and a bullet list is one frame. So the single
most common thing anyone wants from an animation — the points arriving one at a
time — is the one thing this editor cannot express. Its out of scope names the
gap exactly ("emphasis effects, motion paths, per-paragraph text builds …
repeat/rewind/auto-reverse, triggers bound to clicking a specific shape") and
`specs/000-roadmap.md`'s backlog repeats it. A deck arriving with a by-paragraph
build loses nothing today — the timing is preserved verbatim — but it cannot be
seen, played or edited either.

This file widens that subset along four axes and no further: text builds, a
four-effect emphasis set, a small motion-path set, and the repetition and trigger
fields those need to be honest. **SPEC-086 remains the normative source for the
animation model, the canonical `p:timing`/`anim:*` forms, build playback semantics
and the golden-first process; this file specifies only the delta and adds new
goldens by that same process**, and anything outside the *extended* subset still
models no animations and keeps its bytes verbatim. The build model is mode-based,
not index-based, so editing text under a live build needs no repair pass.

## Acceptance criteria
- [ ] `SlideAnimation` (`src/lib/presentation/model.ts`) gains `triggerFrameId?`, `buildBy?`, `buildTrigger?`, `repeatCount?` (2–10, absent = play once), `repeatUntilClick?: true`, `autoReverse?: true`, `rewindOnComplete?: true`, `spinDegrees?` (±3600, absent 360), `scalePercent?` (10–400, absent 150), `emphasisColor?` (`#rrggbb`, absent `#ff0000`) and `path?: MotionPath` per "The model delta"; `AnimationEffect` gains `pulse`/`spin`/`growShrink`/`colorChange`/`motionPath`, `AnimationTrigger` gains `onClickOf`, and the model stays plain JSON so `roundTrip.spec.ts`'s `JSON.parse(JSON.stringify(model))` `toEqual(model)` still holds
- [ ] `src/lib/presentation/animations.ts` gains `effectClass(effect): "entrance"|"exit"|"emphasis"|"path"`, `EMPHASIS_EFFECTS`, `PATH_PRESETS`, `MIN/MAX_REPEAT_COUNT`, `effectiveRepeatCount(a): number` beside `effectiveAnimationDurationMs` (`repeatCount ?? 1`, with `repeatUntilClick` counting as exactly one pass), `DEFAULT_SPIN_DEGREES`, `DEFAULT_SCALE_PERCENT`, `DEFAULT_EMPHASIS_COLOR`, `DEFAULT_PATH_DISTANCE = 0.25`, `mainSequence(animations)` and `interactiveSequences(animations)`; `isEntranceEffect` is `false` for every emphasis and path effect, and `effectiveTrigger` forces `click` only on the first animation **of the main sequence** — an `onClickOf` row keeps its trigger and never opens a main-sequence group
- [ ] `src/lib/presentation/buildPlayback.ts` gains `BuildStep { animation; frameId; paragraphIndex?; key }` (`key` = `frameId` or `` `${frameId}#${paragraphIndex}` ``), `expandAnimation(a, frame)`, `slideBuildSteps(slide): BuildStep[][]` (main sequence only), `slideInteractiveBuilds(slide)` and `stepTimeline(steps)`; paragraph addressing is derived per "Per-paragraph builds" and never stored, so adding, deleting or reordering paragraphs changes the step count and repairs nothing, while a frame with fewer than two paragraphs (or an `ImageFrame`) expands to one whole-frame step and disables `animation-build-by` with helper text
- [ ] `isFrameHiddenAt` stops treating a played emphasis or path effect as an exit — it updates visibility only for `effectClass` `"entrance"`/`"exit"`, its SPEC-086 expectations unchanged; the companion `isBuildTargetHiddenAt(groups, key, groupsPlayed)` drives per-paragraph visibility, and `playbackDurationMs` returns `single × effectiveRepeatCount(a) × (autoReverse ? 2 : 1)` with `repeatUntilClick` contributing exactly one pass so an `afterPrevious` chain never stalls
- [ ] `slideBuildGroups`/`slideBuildCount` count main-sequence groups only, so SPEC-084's `buildCountAt` never waits on an interactive group; during a show a pointer-down on a frame that is some animation's `triggerFrameId` plays that group and consumes the input instead of advancing, mid-build input still completes the group instantly, and reduced motion settles every emphasis, repeat and path to its end state with no intermediate frame
- [ ] `AnimationsPanel.vue` groups `animation-effect` into `<optgroup>`s (Entrance / Emphasis / Exit / Motion path — existing `selectOption("flyIn")` paths keep working) and gains, each shown only for the effects that use it and all keyboard-operable: `animation-build-by` (All at once / By paragraph / By first-level paragraph), `animation-build-trigger`, `animation-repeat`, `animation-repeat-until-click`, `animation-auto-reverse`, `animation-rewind`, `animation-trigger-frame` (enabled only for "On click of", listing the slide's other animatable frames by `frameAccessibleName`), `animation-spin-degrees`, `animation-scale-percent`, `animation-emphasis-color`
- [ ] The timeline is a second view of the same list, not a replacement: `animation-timeline-toggle` (`aria-pressed`, state persisted as `presentationAnimationTimelineOpen` in `src/lib/preferences.ts` — declaration, default `false` and `isBoolean` validator beside the other three presentation keys, at declarations 82/85/88, defaults 146–148 and validators 209–211) reveals `animation-timeline` inside the panel — an `animation-timeline-ruler` at 60px/second with 1s ticks, one `animation-timeline-group-${g}` band per click group and one `animation-timeline-bar-${index}` per row offset by its `stepTimeline` `startMs` and sized by its playback duration (minimum 8px) so overlaps are visible; a bar drives the same `selectedIndex`, its left edge sets delay and its right edge duration by pointer drag or Left/Right arrow in 100ms steps, clamped to SPEC-086's 0–10000 / 100–5000 bounds
- [ ] Motion paths: `animation-path-preset` (line / arc / turn / loop / custom), `animation-path-direction`, `animation-path-distance` (5–100% of the slide dimension, default 25%), and `animation-path-edit` arming an overlay in `AnimationStageLayer.vue` (`animation-path-overlay`) drawing the resolved path as dashed SVG with a draggable `animation-path-endpoint` and a ghost of the frame at the path end; coordinates are slide fractions with the origin at the frame's centre, and an imported `preset: "custom"` path renders read-only with `animation-path-custom-note` — dragging is refused, only choosing another preset replaces it
- [ ] pptx write/read gains the delta in "`p:timing` delta": `p:bldLst`/`p:bldP` plus `p:spTgt/p:txEl/p:pRg` for builds, `presetClass="emph"` presets, `p:animMotion` under `presetClass="path"`, `repeatCount`/`autoRev` and `fill="remove"` on the effect `p:cTn`, and `nodeType="interactiveSeq"` sequences for `onClickOf`; `readPptxTiming` accepts more than one `p:seq` **only** as exactly one `mainSeq` plus zero or more `interactiveSeq`s with resolvable trigger spids — every other multi-sequence tree, `p:audio`/`p:video`, unknown preset or unresolvable target still aborts to `undefined`
- [ ] odp write/read gains the delta in "odp `anim:*` delta": per-paragraph `text:id="de-f{n}-p{k}"` minted by `odp/serialize.ts` on the animated frame's `text:p` elements and targeted by `smil:targetElement`, the `ooo-emphasis-*`/`ooo-motionpath-*` preset ids, `anim:animateMotion` with `svg:path`, `smil:repeatCount`/`smil:autoReverse`/`smil:fill`, and `presentation:node-type="interactive-sequence"` with `smil:begin="de-f{n}.click"`; the reader additionally tolerates Impress's `<anim:iterate anim:iterate-type="by-paragraph">` as `buildBy: "byParagraph"` with `buildTrigger: "afterPrevious"` and `delayMs` from `anim:iterate-interval`
- [ ] Preserve-or-replace stays all-or-nothing per slide: an unmapped preset, a partial `p:pRg` cover, a `p:bldP` disagreeing with the effect pars, or a path this spec cannot resolve models no animations and keeps `preserved["p:timing"]`/`preserved["anim"]` verbatim; a no-edit download stays byte-identical part-for-part and SPEC-086's `animation-replace-confirm` still gates the first modelled edit on such a slide
- [ ] Because the extended subset now models spin, SPEC-086's odp exotic fixture moves further out: `ODP_EXOTIC_ANIM` in `src/lib/presentation/fixtures/build.ts` swaps `ooo-emphasis-spin` for `ooo-emphasis-flash-bulb` and `animationsFixtures.spec.ts`'s single `toContain("ooo-emphasis-spin")` follows it — the same assertion about the same preserve path, coverage preserved and never weakened; the pptx exotic slide is untouched, which the emphasis preset table guarantees by never claiming `emph:8`
- [ ] `useBuildPreview` returns `paragraphStyles: Ref<Record<string, Record<number, FrameStyle>>>` alongside `frameStyles`, threaded through a new `paragraphStyles` prop on `SlideStage.vue` → `SlideFrameView.vue` and merged into each `slide-paragraph` div's style; every new control dirties the slide through `markSlideEdited`, records exactly one undo entry per gesture in SPEC-080's voice ("Change animation build", "Change animation repeat", "Change animation auto-reverse", "Change animation rewind", "Change animation trigger frame", "Change motion path", "Edit motion path" — a timeline or endpoint drag records once on pointer-up) and persists write-through
- [ ] Round trip in both formats: a deck given every new effect, a by-paragraph build, a motion path, a repeat + auto-reverse + rewind row and an `onClickOf` row → save → reopen reproduces `slide.animations` exactly, asserted with `equalStructural` (`src/lib/equality.ts`) after mapping every `frameId`/`triggerFrameId` to its frame's z-order index (neither key is in `IGNORED_KEYS`, and frame ids are fresh UUIDs per read) plus field-level `toEqual`; `equalSemantic` is typed `(model: DocumentModel)` and is **not** usable on a `PresentationModel`, so that index-normalised structural comparison is the presentation-side relation this spec requires; every part the edit did not touch is `equalBytes` (`src/lib/core/equality.ts`) after `unzip`, and writing the reopened model twice yields identical slide-part bytes
- [ ] Fixtures: `fixtures/build.ts` gains `buildPptxDepthFixtureBytes`/`buildOdpDepthFixtureBytes` + `EXPECTED_DEPTH_ANIMATIONS` under a new banner (the four mini-deck slides and SPEC-085/086's builders are never mutated), each deck carrying a by-paragraph bullet slide, an emphasis + motion-path slide, an interactive-trigger slide and one deliberately out-of-subset slide; `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec105-animation-depth.pptx` and `.odp` before its `process.exit(0)`; `PresentationView.vue` gains **no** wiring call and no template hookup — every addition lands inside the already-mounted `<AnimationsPanel>`/`<AnimationStageLayer>`, the existing `useAnimations` controller and a new `src/views/presentation/useMotionPathEdit.ts` it calls itself — so `PresentationView.size.spec.ts`'s `LINE_BUDGET` never rises
- [ ] e2e (`e2e/SPEC-105.spec.ts`): open `spec105-animation-depth.pptx` → select the bullet body, Animate → set "Fade in" + "By paragraph" → the timeline shows four bars in four groups → Present → the slide shows with no bullets → advance four times, one bullet each → back in the editor add an emphasis "Spin" on the title triggered "On click of" the shape, and a "Line" motion path on the shape → Present → clicking the shape spins the title without advancing the slide → exit, download, reopen → the panel lists the build, the spin and the path in order with their fields intact

## Out of scope
Word-level and letter-level builds, by-series/by-category chart builds (SPEC-092's
`ChartFrame` animates as a whole frame) and per-cell table builds (SPEC-091 widens
the target kinds only). Freeform path drawing and bezier point editing (SPEC-071
owns that vocabulary text-side; here an imported custom path is rendered and
preserved, never authored). Easing and SMIL accelerate/decelerate/`p:tavLst`
shaping, sound cues and media action verbs (SPEC-095 owns audio and video),
hyperlink and bookmark triggers, animation on group/foreign frames (SPEC-086's
rule stands), effect options beyond the fields listed, an animation painter,
rehearsal timings, and reading arbitrary foreign timing into the model — preserve
or replace only.

## Reference

### Dependencies restated
- **SPEC-086** (`specs/086-slide-animations.md`) owns everything this file extends:
  `SlideAnimation { frameId; effect; direction?; trigger; durationMs?; delayMs? }`,
  its seven entrance/exit effects, `AnimationTrigger`
  `click|withPrevious|afterPrevious`, duration 100–5000 (default 500), delay 0–10000
  (default 0); order in `Slide.animations` is play order and a present-even-empty
  array is the modelled state. Its modules are `animations.ts`, `slideAnimations.ts`
  (whose private `ensureModelled` deletes `preserved["p:timing"]`/`["anim"]`),
  `buildPlayback.ts` and the wire pairs `writePptxTiming`/`readPptxTiming` and
  `writeOdpAnim`/`readOdpAnim`.
- **SPEC-081** owns slide-part serialisation and `markSlideEdited(slide, format)`,
  which every command here ends with: `pptx/serialize.ts` mints `spidFor` =
  `p:cNvPr id` 2,3,… in z-order over text/shape/image frames; `odp/serialize.ts`
  mints `wireIdFor` = `de-f1`, `de-f2`, … on `xml:id` **and** `draw:id`, only when
  `slide.animations !== undefined`, so animation-free pages stay byte-verbatim; and
  `serializeSlideXml`'s slot order is `…{clrMapOvr}{transition}{timing}{extLst}`,
  `modelled ?? preserved ?? ""` at every slot, never both.
- **SPEC-082** frame identity plus `frameAccessibleName(frame)`
  (`useFrameManipulation.ts`); **SPEC-084** the normative advance pipeline
  (`PlaybackStep { slideIndex; buildIndex }`; `advanceStep()` completes an in-flight
  build and holds); **SPEC-080** undo — `record(label?)` **before** the mutation,
  mutate, `commit()`, one entry per gesture; **SPEC-088** the exported `data-anim-*`
  attribute set. `AnimationsPanel.vue`/`AnimationStageLayer.vue` declare props at
  runtime (espree lints `.vue` scripts — no `defineProps<{…}>()`), emit nothing and
  alias `const c = props.controller`; `useAnimations(deps)` keeps its shape.

### The model delta
```ts
export type AnimationEffectClass = "entrance" | "exit" | "emphasis" | "path";
export type TextBuildMode = "byParagraph" | "byFirstLevelParagraph";

export interface MotionPath {
  preset: "line" | "arc" | "turn" | "loop" | "custom";
  direction?: TransitionDirection;  // line/arc/turn; absent = "right"
  distanceFraction?: number;        // 0.05..1 of slide width (x)/height (y); absent = 0.25
  d?: string;                       // `custom` only, and then required
}
```
`SlideAnimation` keeps SPEC-086's six fields and gains the eleven optional ones
AC-1 lists, in that order and with those bounds and defaults; `buildBy` absent
means the whole frame animates at once, `buildTrigger` absent means `"click"`, and
only `true` is ever stored for the three flags. Three constraints live only here:
`triggerFrameId` is required iff `trigger === "onClickOf"` and names another frame
on the slide, never the animation's own target; `repeatCount` and
`repeatUntilClick` exclude each other; `path` is present iff the effect is
`motionPath`. Identifiers keep the codebase's American spelling (`color`, as
`SlideRun.color` does) while prose says colour, a `SlideAnimation` still carries no
id, and `equalStructural`'s `IGNORED_KEYS` drops `"id"` but neither `frameId` nor
`triggerFrameId`.

### Per-paragraph builds — expansion, never stored indices
`expandAnimation(a, frame)` is the single source of truth, used by both players,
both serialisers and SPEC-088's exporter:
- no `buildBy`, a target with no `paragraphs`, or fewer than two paragraphs → one
  step, `key = frameId`, no `paragraphIndex`;
- `"byParagraph"` → one step per paragraph in order, `` key = `${frameId}#${i}` ``;
- `"byFirstLevelParagraph"` → one step per paragraph whose `level` is 0 or absent,
  covering it and every deeper paragraph up to the next level-0 one; a leading
  deeper paragraph joins the first step.

`slideBuildSteps(slide)` then partitions the expanded main sequence into groups:
the row's own `trigger` opens its first step's group exactly as SPEC-086 defines,
and each later step opens a new group when `buildTrigger` is `"click"` (the
default) or joins the current one otherwise.

### Emphasis, motion paths, repetition and click-of triggers
Four emphasis effects, class `"emphasis"`: `pulse` (one scale-up-and-back throb),
`spin` (by `spinDegrees`), `growShrink` (to `scalePercent` and back) and
`colorChange` (to `emphasisColor` and back); one path effect, class `"path"`:
`motionPath`. Path coordinates are normative in both formats — **fractions of the
slide**, 1.0 being the full `slideWidthCm` on x and `slideHeightCm` on y, with the
origin at the frame's current centre, so a path is deck-size independent and
survives SPEC-090's rescale. `line` is a straight `direction` × `distanceFraction`
move, `arc` bows perpendicular by half that, `turn` is an L, `loop` is a closed
circle of radius `distanceFraction / 2`, and `custom.d` accepts only
`M`/`L`/`C`/`Z` with numeric arguments, else out of subset. `rewindOnComplete`
restores the frame's settled style when the effect ends — an entrance with rewind
therefore hides its target again, which both formats match and
`isBuildTargetHiddenAt` reports. `trigger: "onClickOf"` takes the row out of the
main sequence entirely: interactive rows group by `triggerFrameId` in array order,
play on a click of that frame and never count towards `slideBuildCount`.

### `p:timing` delta (write; read tolerantly)
Extends SPEC-086's canonical tree; anything not named here is unchanged.
- **Builds.** Each expanded step's `p:spTgt` gains a `p:txEl/p:pRg` carrying the
  step's first and last paragraph index, and `p:timing` gains a `p:bldLst` of
  `p:bldP spid grpId="0" build="p"` siblings after `p:tnLst` (`bldLvl="1"` for
  `byFirstLevelParagraph`). Reading collapses back to one row only when the
  `p:pRg` ranges partition `[0, paragraphCount)` in order with no gap or overlap,
  share effect/direction/duration/delay and agree with the `p:bldP`.
- **Effects.** `presetClass="emph"`, one presetID per effect, over `p:animRot
  by="{spinDegrees × 60000}"`, `p:anim` on `ppt_w`/`ppt_h` and `p:animClr`;
  `presetClass="path"` over `p:animMotion` (`path="M 0 0 L 0.25 0"`,
  `pathEditMode="relative"`) whose `p:cBhvr/p:attrNameLst` lists `ppt_x`/`ppt_y`.
- **Repetition and triggers.** On the effect's own `p:cTn`, `repeatCount="{n ×
  1000}"` or `"indefinite"`, `autoRev="1"` and `fill="remove"` in place of
  `fill="hold"`; one extra `p:seq` per trigger frame, sibling to the mainSeq
  inside the tmRoot's `p:childTnLst`, `nodeType="interactiveSeq"` begun by an
  `evt="onClick"` condition on the trigger's `p:spTgt`.

### odp `anim:*` delta (write; read tolerantly)
- **Builds.** `odp/serialize.ts` mints `text:id="de-f{n}-p{k}"` on the animated
  frame's `text:p` elements — only for frames a `buildBy` row targets, so every
  other page byte is unchanged — and each expanded step's behaviours carry
  `smil:targetElement="de-f{n}-p{k}"`, which `resolveTarget` maps back to
  `(frameId, paragraphIndex)`. Tolerant read also accepts `<anim:iterate
  anim:iterate-type="by-paragraph">` wrapping the behaviours.
- **Presets.** `presentation:preset-class="emphasis"` with `ooo-emphasis-pulse`,
  `ooo-emphasis-spin`, `ooo-emphasis-grow-and-shrink`, `ooo-emphasis-fill-color`;
  `presentation:preset-class="motion-path"` with
  `ooo-motionpath-{line|arc|turn|loop}` and `ooo-motionpath-custom`.
  **Behaviours.** `anim:animateTransform smil:type="rotate" smil:by="{deg}"` and
  `="scale"`; `anim:animateColor smil:attributeName="fill-color"
  smil:to="#rrggbb"`; `<anim:animateMotion svg:path="m 0 0 l 0.25 0" smil:dur="…"
  smil:targetElement="…"/>` in the same slide-fraction space.
- **Repetition and interactive.** `smil:repeatCount="{n}"` or `"indefinite"`,
  `smil:autoReverse="true"`, `smil:fill="remove"` vs `"hold"`; a sibling
  `<anim:seq presentation:node-type="interactive-sequence"
  smil:begin="de-f{n}.click">`, `parseAnim`'s single-`main-sequence` rule standing
  and extra sequences parsing only when every one is an interactive sequence with
  a resolvable trigger id.

### Golden-first, restated for both formats (normative process)
Exact presetIDs, subtypes, behaviour internals and uncertain ODF spellings —
`svg:path` vs `smil:path` on `anim:animateMotion`, Impress's exact
`ooo-motionpath-*` suffixes — are **not** named from memory, and neither is any
element nesting the delta sections above imply. SPEC-086's process applies
verbatim: `pptx/timingGolden.ts` gains `DEPTH_ANIMATIONS` + `TIMING_DEPTH_GOLDEN`
and `odp/animGolden.ts` gains `ANIM_DEPTH_GOLDEN` over the same model,
hand-authored against ECMA-376 §19.5/§19.7 and ODF 1.2 §16.29 and reviewed
against the behavioural round trip (cross-app verification is manual QA recorded
in the PR, not CI); the serialisers are asserted equal to them, the readers proven
to accept them with both spellings tolerated, and the modules' preset tables are
the normative artefact. `TIMING_GOLDEN`, `ANIM_GOLDEN` and `GOLDEN_ANIMATIONS`
stay byte-identical.

### Playback, the timeline and the path editor
`useBuildPreview` keeps its shape, drives `BuildStep[]` instead of
`SlideAnimation[]` and adds `paragraphStyles`. `animationStyles.ts` gains
`fromStyle`/`toStyle`/`transitionFor` cases for the five new effects
(pulse/growShrink `transform: scale(…)`, spin `rotate(…)`, colorChange
`color`/`background-color`, motionPath a `translate(…)` pair resolved from the
path at the deck's cm size) plus `animationIterationCount`/`animationDirection`.
`AnimationTimeline.vue` is presentation-only, its geometry (`barRect`,
`rulerTicks`, `msFromDx`) living in a sibling `animationTimeline.ts`, and
`useMotionPathEdit.ts` owns the endpoint drag: pointer capture, `stageMath.ts`'s
`clientToCm`, live `refresh()` while dragging, one `record` + `commit` on
pointer-up.

### Tests
- `presentation/animations.spec.ts` — `effectClass`, the main/interactive partition,
  `effectiveTrigger`'s first-of-main-sequence rule, `effectiveRepeatCount`, the new
  bounds; `slideAnimations.spec.ts` — the widened `AnimationPatch`;
  `buildPlayback.spec.ts` — all three build modes (level-0 grouping and the
  shrinking-paragraph case included), `isFrameHiddenAt` no longer hiding an
  emphasised frame, `isBuildTargetHiddenAt`, repeat/auto-reverse durations.
- `pptx/timing.spec.ts`, `odp/anim.spec.ts` — golden equality both directions,
  per-effect round trips, PowerPoint- and Impress-flavoured tolerant variants,
  out-of-subset cases (partial `p:pRg` cover, `p:bldP` mismatch, unresolvable
  trigger, a non-interactive extra sequence, a malformed path).
- `presentation/animationsRoundTrip.spec.ts` and `animationsFixtures.spec.ts` per
  their criteria; `views/presentation/useAnimations.spec.ts`, new
  `useMotionPathEdit.spec.ts`, `useBuildPreview.spec.ts`,
  `animationTimeline.spec.ts` and `e2e/SPEC-105.spec.ts`.

### Specs this changes
- SPEC-086 stays the owner; its model section, preset tables and goldens are
  extended, never replaced, and its out-of-scope list loses the five items this
  file takes (emphasis effects, motion paths, per-paragraph text builds,
  repeat/rewind/auto-reverse and click-a-shape triggers). Media playback effects
  and timing rehearsal remain out of scope there; easing is this file's own
  exclusion, never SPEC-086's.
- SPEC-084's `buildCountAt` now excludes interactive groups — a clarification of
  `usePlayback`'s contract, not a change to its step machine. SPEC-088's exported
  `data-anim-*` set grows by the build, repeat, reverse, rewind, trigger-frame and
  path fields, and its exporter consumes `expandAnimation`. SPEC-091's and
  SPEC-095's widening of `isAnimatableKind` composes unchanged: a new frame kind
  gains builds and paths once it mints a wire id.
- `specs/000-roadmap.md` gains this spec's phase row and drops the backlog line
  "Per-paragraph text builds, emphasis effects, motion paths | SPEC-086's out of
  scope"; nothing else changes.
