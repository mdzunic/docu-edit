---
id: SPEC-084
title: Slideshow playback
status: draft
area: shell
depends_on: [SPEC-079, SPEC-080]
---
## Why
A presentation editor that cannot present is a diagram tool. This spec
adds the playback surface: a full-screen, chrome-free mode that walks the
deck slide by slide with the keyboard, mouse or touch, from any starting
slide, ending on an explicit end screen. It is deliberately transition-
and animation-free — SPEC-085/086 plug their effects into the advance
pipeline this spec defines, and SPEC-088's HTML export replicates this
surface's semantics outside the app, so **this file is the normative
source for playback navigation semantics** (what "advance" means, key
map, end screen, exit behaviour).

## Acceptance criteria
- [ ] The deck header gains a "Present" button (testid `slides-present`, keyboard shortcut Cmd/Ctrl+Shift+Return) that starts playback at the current slide; a deck with zero slides disables it with a tooltip
- [ ] Playback is a route — `/slides/:id/present/:n` (1-based) — rendered by a dedicated `PresentationPlayView` with no app chrome (no AppShell/toolbar/filmstrip); visiting it directly works (deep link), an out-of-range `n` clamps, a text-document id redirects to `/doc/:id` like the editor view does, and an unknown id shows the not-found state with a "Back to library" action
- [ ] Entering playback requests browser fullscreen on the stage container (from the button's user gesture); denial or unavailability is tolerated — playback then fills the viewport — and the user leaving OS fullscreen (Esc) exits playback cleanly rather than leaving a stranded black page
- [ ] The current slide renders through the SPEC-079 stage, scaled to the largest size that fits the viewport at the deck's aspect ratio, letterboxed on black; images/shapes/text render exactly as the editor stage does (one renderer, two hosts)
- [ ] Advance — click anywhere, ArrowRight, ArrowDown, Space, PageDown, Enter, N — goes to the next step (this spec: next slide; SPEC-085/086 insert transition/build steps into the same pipeline); Back — ArrowLeft, ArrowUp, PageUp, P, Backspace — goes to the previous slide; Home/End jump to first/last; typing a number then Enter jumps to that slide; on touch, tap advances and horizontal swipe navigates (SPEC-068 pointer conventions)
- [ ] Advancing past the last slide shows the end screen — black, "End of slide show — click or press any key to exit" in `--text-inverse`-on-black styling, testid `play-end-screen`; one more advance (or Esc) exits
- [ ] Esc (when not consumed by OS fullscreen), and the end-screen exit, leave playback: the route returns to `/slides/:id`, the editor selects the slide that was showing (or the last slide when exiting from the end screen), and browser fullscreen is released if still held
- [ ] The URL tracks the current slide (`router.replace`, no history spam: one history entry total for the whole playback session — Back in the browser exits playback rather than stepping slides)
- [ ] A slide counter ("3 / 12") fades in on pointer movement and fades out after 2s idle, along with a minimal on-screen prev/next affordance for touch; both hidden otherwise and during the idle state (testids `play-slide-counter`, `play-prev`, `play-next`)
- [ ] Playback is inert content: no selection handles, no editing shortcuts (Cmd/Ctrl+Z does nothing), no text cursor; frames render without their editor affordances, and playback never dirties the model
- [ ] Screen-reader semantics: the stage container is `role="region"` `aria-roledescription="slide show"` labelled with the deck title; slide changes announce "Slide N of M — <title or 'untitled'>" via a polite live region; the end screen announces itself
- [ ] Reduced motion is respected from birth: the view reads `prefers-reduced-motion` and exposes it to the advance pipeline (SPEC-085/086 honour it; this spec has no motion to reduce, but the plumbing and a unit test exist so they cannot forget)
- [ ] e2e (`e2e/SPEC-084.spec.ts`): open the pptx fixture → Present → slide 1 content visible without app chrome → ArrowRight to slide 2 (URL updated, counter correct after mouse move) → End → advance to the end screen → advance again → back on `/slides/:id` with the last slide selected. Fullscreen assertions are conditional (headless grants are flaky): assert playback works regardless of the fullscreen promise's outcome

## Out of scope
Presenter view (notes, next-slide preview, timer) — future work over
SPEC-083's notes. Transitions/animations (SPEC-085/086). Auto-advance
timers (SPEC-085 introduces per-slide auto-advance with its transition
model). Laser pointer/annotation tools, black/white blanking keys (B/W),
looping ("kiosk") mode, external-display awareness, printing. Remembering
playback position across sessions (exit selection only).

## Reference

### Dependencies restated
- SPEC-079: `SlideStage.vue` (props: slide, model, scale — the play view
  passes a computed scale from viewport size; the stage must not depend
  on editor state to render), `slideOutline` for announcements,
  `slideWidthCm`/`slideHeightCm` for aspect.
- SPEC-080: current-slide selection state in `PresentationView` (the
  Present button starts from it; the exit path writes it back — pass via
  the route's `n` and a `?return` handled in the editor's mounted sync,
  or the shared store selection SPEC-080 introduced; either way the
  observable behaviour is the AC's).
- Router conventions: routes declared in `src/router.ts`; the kind guard
  mirrors `/slides/:id`'s (`documentKindOf`).

### The advance pipeline (normative for SPEC-085/086/088)
`src/views/presentation/usePlayback.ts` owns a step machine:
```ts
interface PlaybackStep { slideIndex: number; buildIndex: number }
advance(): void   // next build on the slide, else next slide (builds
                  // are [] until SPEC-086), else the end screen
back(): void      // previous slide (builds collapse on back — going
                  // back lands on the previous slide fully built)
goTo(slideIndex: number): void
```
plus `{ current, atEnd, reducedMotion }` reactive state. SPEC-085 wraps
slide changes in transitions (delaying `current` visually, never
logically); SPEC-086 populates builds; SPEC-088 reimplements the same
machine in the exported file's script — behavioural parity is asserted
by sharing the key-map table as data
(`src/lib/presentation/playbackKeys.ts`: `advanceKeys`, `backKeys`,
exported for both the view and the HTML export template).

### View
`src/views/PresentationPlayView.vue` (+ route
`/slides/:id/present/:n`). Black page (`background: #000` explicit —
theme-independent), stage centred via flex, scale =
`min(vw / stagePxWidth, vh / stagePxHeight)`, recomputed on resize/
fullscreenchange. Keyboard handling on `window` while mounted, ignoring
events with modifier combos it doesn't own and anything from
inputs. Fullscreen: `container.requestFullscreen().catch(() => {})`;
`fullscreenchange` → if exited and still on the route, `router.replace`
back to the editor. Idle-fade via a 2s timer reset on `pointermove`.

### Tests
- `usePlayback.spec.ts`: step machine (advance/back/goTo/end), clamp,
  key-map table coverage, reduced-motion plumbing.
- View spec: route guards (text id redirect, unknown id, clamp),
  history behaviour (`replace` not `push`), announce strings, no-chrome
  render, exit selection handoff.
- e2e per AC. (Memory: never assert absolute pixel geometry — assert
  visibility and content; fullscreen conditionally.)

### Specs this changes
- SPEC-079's PageUp/PageDown slide navigation in the editor is
  restated: those keys remain editor-side; playback owns its richer map
  via `playbackKeys.ts`, and the editor imports the same table for its
  two keys so the maps cannot drift.
