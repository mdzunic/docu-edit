---
id: SPEC-100
title: Presenter view — notes, next slide and the second window
status: draft
area: shell
complexity: normal
story_points: 8
depends_on: [SPEC-066, SPEC-083, SPEC-084, SPEC-085, SPEC-086]
---
## Why
SPEC-083 models speaker notes and shows them under the stage; SPEC-084 presents
full screen from any slide; SPEC-086 builds slides step by step. Both of the
first two explicitly defer the thing a presenter actually stands behind —
SPEC-083's out of scope says "A presenter/rehearsal view", SPEC-084's says
"Presenter view (notes, next-slide preview, timer) — future work over
SPEC-083's notes" — and `specs/000-roadmap.md`'s backlog carries the row
"Presenter view — notes, next slide, timer | SPEC-083/084 both name it". It is
the most-asked-for presentation feature this product does not have, and every
input it needs now exists.

The shape is a second same-origin window and one owner. The editor tab becomes
the presenter window; it opens the audience window with `window.open` and talks
to it over a same-origin channel — `BroadcastChannel` where available, a
`localStorage` `storage`-event fallback where it is not, so no dependency is
added (`package.json` is protected). The step machine is not forked: the
presenter window owns the single `usePlayback` instance and the audience window
becomes a follower whose machine is written only by what arrives. Everything the
browser cannot promise — popups, fullscreen without a gesture, placing a window
on a projector — degrades to a stated, visible fallback rather than a hard
requirement.

**This file is the normative source for the presenter session protocol (the
owner/follower contract, its message vocabulary and its failure modes), the
presenter surface, and blank-screen (B/W) semantics, which SPEC-084 deferred.**
It changes no file format, no model field and no writer. SPEC-088's HTML export
stays deliberately single-window.

## Acceptance criteria
- [ ] `DeckActions.vue`'s Present control gains a menu affordance beside it (testids `presentation-present-menu`, `presentation-present-menu-popup`, `presentation-present-presenter`, copying `PresentationDownloadButton.vue`'s existing menu pattern, keyboard-operable) and `present` is emitted with `"show" | "presenter"`; `useDeckActions.present(mode = "show")` pushes `/slides/:id/present/:n` or `/slides/:id/presenter/:n`, `onPresentShortcut` keeps Cmd/Ctrl+Shift+Return for the show and adds Cmd/Ctrl+Alt+Shift+Return for presenter view, and `src/views/PresentationView.vue` changes exactly one line (`@present="deckActions.present()"` → `@present="deckActions.present"`) so it gains no net lines and `LINE_BUDGET` in `src/views/PresentationView.size.spec.ts` never rises above its present 550
- [ ] `src/router.ts` gains `/slides/:id/presenter/:n` (name `slides-presenter`, 1-based `n`) with the same `documentKindOf(id) === "text" → /doc/:id` guard as the other two slide routes, rendered by a new `src/views/PresenterView.vue` (root testid `presenter-view`) that loads the deck exactly as `PresentationPlayView.vue` does (`initLibrary().then(load)`, legacy upgrade, out-of-range `n` clamped, `presenter-not-found` + "Back to library" for an unknown id)
- [ ] `src/lib/core/windowChannel.ts` is a **new** product-neutral module — nothing is moved, because the repo has no general cross-window transport to hoist (`grep -rn "BroadcastChannel" src/` finds only `core/editLock.ts:2`'s header explaining why the *lock* chose Web Locks; the one cross-window mechanism that does exist, `src/lib/preferences.ts:393`'s `storage` listener, is hard-keyed to `PREFERENCES_STORAGE_KEY` and carries no payload) — exporting `openWindowChannel(name, onMessage): WindowChannel { post(payload: unknown): void; close(): void }` and `windowChannelTransport(): "broadcast" | "storage" | "none"` per the Reference, with `core/boundary.spec.ts` still green (no Vue, no `../`) and an injectable transport for unit tests
- [ ] `src/lib/presentation/presenterSession.ts` holds the framework- and DOM-free protocol per the Reference — `PresenterState`, the `PresenterMessage` union, `presenterChannelName(deckId)`, `PRESENTER_HEARTBEAT_MS = 1000`, `PRESENTER_LOST_MS = 3000`, `isStaleState(prev, next)`, and `parsePresenterMessage(raw: unknown): PresenterMessage | null` which treats every payload from another window as untrusted (wrong shape, wrong `session`, or a non-finite number is dropped, never thrown); it imports only `core/*`, keeping `presentation/boundary.spec.ts` green
- [ ] `usePlayback` gains `setStep(step: PlaybackStep, atEnd = false): void` — indices clamped exactly as `goTo` clamps — as the **only** way a follower's machine is written; the presenter window owns the machine, broadcasts `state` on every change plus a `PRESENTER_HEARTBEAT_MS` heartbeat, and the audience window in `?follow=` mode sends `intent` messages for every local key/click/swipe and renders only what comes back, so its existing `watch(playback.current)` drives builds (SPEC-086) and transitions (SPEC-085) identically whichever window took the input
- [ ] Failure modes per the Reference: audience closed or reloaded → the presenter keeps its state and shows `presenter-audience-lost` with `presenter-audience-open`, and a returning window re-syncs from its `hello` with one full `state` snapshot; presenter closed or navigated away → it posts `bye` and the audience promotes itself to owner (a role flip, never a second machine) and continues as plain SPEC-084 playback; `PRESENTER_LOST_MS` without any message promotes the same way, so a stale `?follow=` URL opened alone still works
- [ ] `window.open` returning `null` (popup blocked) or `windowChannelTransport() === "none"` falls back to single-window presenter mode: a banner `presenter-audience-blocked` offering `presenter-audience-open` (a fresh gesture usually succeeds) and "Present here", `presenter-solo-note` stating which mode is live, and the presenter's own current-slide pane becoming the show — no toast, no dead end
- [ ] The presenter surface shows, per the Reference: `presenter-current-slide` (driven by its own `useBuildPreview()` so it mirrors the audience's build state exactly), `presenter-next-slide` rendered fully built through `SlideStage` with no `frameStyles` (`presenter-next-end` "End of show" at the last slide), `presenter-slide-counter` ("Slide 3 of 12"), `presenter-build-indicator` ("Build 2 of 4" from `slideBuildCount`, hidden when the slide has none), `presenter-prev-button`, `presenter-next-button`, `presenter-exit`, and a polite `presenter-live`
- [ ] `presenter-notes` renders the current slide's SPEC-083 `notes?: string[]` read-only, one paragraph per entry, "No notes for this slide" when absent, at a scale stepped by `presenter-notes-smaller` / `presenter-notes-larger` (and `-` / `+`) over `PRESENTER_NOTES_SCALES = [1, 1.25, 1.5, 2, 2.5]`, persisted per device as `presentationPresenterNotesScale` (default `1.25`) added in all three places in `src/lib/preferences.ts` — the interface, `PREFERENCE_DEFAULTS`, and the validator table, where the entry is `oneOf<"presentationPresenterNotesScale">(PRESENTER_NOTES_SCALES)` (`preferences.ts:164`, generic over the `Preferences` value type) or `autosaveDelayMs`'s inline numeric form at line 204, **never `inSet`** (line 172), which is `readonly string[]`-typed and returns `undefined` for every number
- [ ] `src/lib/presentation/presenterTimer.ts` is pure and clock-injected — `TimerState { startedAtMs: number | null; accumulatedMs: number }`, `startTimer`, `pauseTimer`, `resetTimer`, `elapsedMs(state, nowMs)`, `formatElapsed(ms)` → `M:SS` below one hour and `H:MM:SS` at or above it, with no `Date.now()` inside so no test couples to the clock; the timer starts with the show, `presenter-timer-toggle` pauses/resumes (`aria-pressed`), `presenter-timer-reset` zeroes it, `presenter-timer` ticks once a second
- [ ] `presenter-grid-toggle` (and `g`; Escape closes) opens `presenter-grid`, a `role="listbox"` of numbered thumbnails `presenter-grid-slide` built from `SlideStage` plus `filmstripThumbScale`/`filmstripLabel` in `src/components/presentation/slideRender.ts` — no second thumbnail maths — with arrow-key roving focus; Enter or click jumps both windows through the owner's `goTo`
- [ ] Blank screen lands here (SPEC-084 deferred it): `src/lib/presentation/playbackKeys.ts` gains `blackoutKeys = ["b", "."]` and `whiteoutKeys = ["w", ","]`, and `PlaybackKeyAction` gains `"blackout" | "whiteout"` resolved **after** advance/back so no existing binding changes; toggling paints the audience surface solid `#000`/`#fff` over the stage (`play-blank`, `aria-hidden` art with the live region announcing "Screen blanked"), the presenter surface stays fully usable and shows `presenter-blank-badge`, and any navigation action clears it
- [ ] Displays and fullscreen stay honest per the Reference: the audience window requests fullscreen on load exactly as `PresentationPlayView.enterFullscreen()` already does, and because a browser may refuse a request that window did not itself gesture for, a persistent `play-fullscreen-prompt` button and the `f` key grant it from a real gesture; when `window.getScreenDetails` exists **and** its permission resolves granted the audience window opens on the first non-primary screen via `left/top/width/height` features, otherwise it opens as an ordinary popup with a one-line hint — asked at most once per session, silent on refusal, never a hard requirement, and `navigator.presentation` is not used at all
- [ ] Presenting is inert and parity-safe: neither window sets `dirty`/`notesDirty`/`structureDirty` or calls `markSlideEdited`, the audience window never calls `acquireEditLock` (it cannot steal the editor's write lock) and a follower never runs `upgradePresentationModel`/`storePresentation` (it waits for the owner's first `state`, then re-reads the store once); a deck presented and exited still downloads byte-identical part-for-part (`equalBytes` from `src/lib/core/equality.ts`, part-level after `unzip`, as `roundTrip.spec.ts` asserts); both windows read `prefers-reduced-motion` from their own display and honour it locally (no motion flag travels), and both classify keys through `playbackActionForKey` so the map cannot drift
- [ ] View sizes ratchet: new `src/views/PresenterView.size.spec.ts` (`LINE_BUDGET = 400`, ledger header copied from `PresentationView.size.spec.ts`) with the logic in `src/views/presentation/usePresenterView.ts` and `usePresenterSession.ts`; `PresentationPlayView.vue` (589 lines, no budget today) gains only its `usePresenterSession({ … })` follower call and the blank/fullscreen-prompt markup, made room for by moving its keyboard, pointer and idle-HUD handling into `src/views/presentation/usePlayInput.ts` with behaviour unchanged, and gains `src/views/PresentationPlayView.size.spec.ts` at `LINE_BUDGET = 600`
- [ ] e2e (`e2e/SPEC-100.spec.ts`, two windows via `page.waitForEvent("popup")`): open `mini-deck.pptx` → Present menu → "Presenter view" → the audience window opens on `/present/1?follow=…` → presenter shows slide 1, the next-slide pane shows slide 2, and the notes pane shows the fixture's first note → ArrowRight in the presenter moves both windows to slide 2 → ArrowRight in the **audience** window moves both to slide 3 → `b` blanks the audience (`play-blank`) and badges the presenter, ArrowRight clears it → grid → slide 4 jumps both → Escape in the presenter exits to `/slides/:id` with slide 4 selected and the audience window closed; a second test opens `spec086-animations.pptx` and asserts `presenter-build-indicator` counts builds as the show advances

## Out of scope
Pen, laser-pointer and ink annotation during a show (no ink model exists on
either side — `docs/presentations-architecture.md`'s "The presentation model
(deliberately minimal)"; it needs a stroke model and a stroke channel, and is
the named follow-up). Editing anything from presenter view — notes included
(SPEC-083 owns notes editing, in the editor). Writing rehearsal timings back
to the file: `p:transition@advTm` and ODF `presentation:duration` belong to
SPEC-085's `advanceAfterMs` and this spec never sets them. Per-slide "time
spent" reports and rehearse-and-record. Kiosk
/ loop mode. Remote control from a phone or another device (the channel is
same-origin, one browser, by construction) and casting via
`navigator.presentation`. A presenter view for SPEC-088's HTML export (that
file stays single-window). Formatted notes (SPEC-083 flattens them). A third
window, and presenting two decks into one audience window.

## Reference

### Dependencies restated
- `specs/084-slideshow-playback.md` owns playback navigation. `src/views/presentation/usePlayback.ts`:
  `PlaybackStep { slideIndex: number; buildIndex: number }`;
  `UsePlaybackDeps { slideCount: Ref<number>; reducedMotion?: Ref<boolean>; buildCountAt?: (slideIndex: number) => number; start?: number }`;
  `UsePlayback { current, atEnd, reducedMotion, advance, back, goTo }` — `advance()` walks builds then slides then sets `atEnd`, `back()` lands the previous slide **fully built**, `goTo(i)` clamps and clears `atEnd`. `src/lib/presentation/playbackKeys.ts`:
  `advanceKeys = ["ArrowRight","ArrowDown"," ","Spacebar","PageDown","Enter","n"]`,
  `backKeys = ["ArrowLeft","ArrowUp","PageUp","p","Backspace"]`, `firstKeys = ["Home"]`,
  `lastKeys = ["End"]`, `normaliseKey`, `playbackActionForKey(key): PlaybackKeyAction | null`.
  `PresentationPlayView.vue` owns the route `/slides/:id/present/:n`, `enterFullscreen()`
  (denial tolerated), `SWIPE_THRESHOLD_PX = 48`, the 2000 ms idle HUD, `router.replace`-only
  URL sync and `exit()` → `/slides/:id?return=<n>`.
- `specs/086-slide-animations.md` owns builds. `slideBuildCount(slide)` /
  `slideBuildGroups(slide)` (`src/lib/presentation/buildPlayback.ts`) and
  `useBuildPreview()` → `{ frameStyles, playing, settle, playGroup, complete, autoplay, stop }`,
  where `frameStyles: Ref<Record<string, Record<string,string>>>` feeds `<SlideStage :frame-styles>`.
- `specs/085-slide-transitions.md` owns `Slide.transition` and the two-stage
  `SlideTransitionLayer`; nothing here changes either.
- `specs/083-speaker-notes.md` owns `Slide.notes?: string[]` — plain paragraphs, empty array
  never stored, absent means none — and the `presentationNotesOpen` preference precedent.
- `specs/066-preferences.md` owns `src/lib/preferences.ts`: a new key is declared in the
  `Preferences` interface, `PREFERENCE_DEFAULTS`, and the validator table (three places).
  `Validator<K> = (raw: unknown) => Preferences[K] | undefined`; `oneOf<K>(values:
  readonly Preferences[K][])` is the value-typed helper and `inSet(raw, values: readonly
  string[])` the string-only one, so a numeric key takes `oneOf` or an inline predicate.
- `SlideStage.vue` props `slide*, widthCm*, heightCm*, media, scale, liveLabel,
  selectedFrameId, editingFrameId, guides, drawPreview, armed, interactive (default true),
  frameStyles`; `filmstripThumbScale(widthCm, thumbWidthPx)` and `filmstripLabel(slide, i, n)`
  in `src/components/presentation/slideRender.ts`. Store access is
  `initLibrary()`, `getPresentation(id)`, `documentKindOf(id)` from `src/lib/documentStore.ts`.
- Round-trip relation: `equalBytes(a: Uint8Array, b: Uint8Array)` from
  `src/lib/core/equality.ts` (the relation every presentation suite uses, always part-level
  after `unzip`). `equalSemantic` in `src/lib/equality.ts` is typed against the text-side
  `DocumentModel` and is **not** usable on a `PresentationModel`; this spec needs no semantic
  relation, because it writes nothing.

### The presenter route and the window pair
The editor tab navigates to `/slides/:id/presenter/:n` and becomes the presenter
window. On load it mints `sessionId = crypto.randomUUID()`, opens the channel
`docu-edit:present:<deckId>`, and calls once:

```
window.open(`/slides/${id}/present/${n}?follow=${sessionId}`,
            `docu-edit-audience-${id}`,               // named: re-opens the same window
            "popup=yes,width=W,height=H,left=L,top=T")
```

The audience window is the existing `PresentationPlayView`; `?follow=` is the only
change to that route. The deck itself never travels over the channel — both windows
read it from IndexedDB exactly as `/present/:n` already does (a model carries base64
media and can be megabytes). Exiting the presenter window posts `bye`, closes the
window it opened, and lands `/slides/:id?return=<n>` like SPEC-084.

### The session contract (normative)
One owner at a time; state flows owner → follower, intents flow follower → owner.

```ts
export interface PresenterState {
  seq: number;               // monotonic per owner; a lower/equal seq is stale
  slideIndex: number;
  buildIndex: number;
  atEnd: boolean;
  blank: "none" | "black" | "white";
}
export type PresenterMessage =
  | { kind: "state";  session: string; state: PresenterState }
  | { kind: "hello";  session: string; from: "presenter" | "audience" }
  | { kind: "intent"; session: string;
      action: "advance" | "back" | "first" | "last" | "goTo" | "blackout" | "whiteout" | "exit";
      slideIndex?: number }
  | { kind: "bye";    session: string; from: "presenter" | "audience" };
```

Rules. The owner posts `state` on every change and every `PRESENTER_HEARTBEAT_MS`.
A follower applies `state` through `usePlayback.setStep(...)` unless
`isStaleState(prev, next)`, and posts `intent` for every local input — it never
advances its own machine. `hello` from either side makes the owner post one
immediate `state`. `bye`, or `PRESENTER_LOST_MS` of silence, promotes a follower to
owner (it stops applying remote state and drives itself); a promoted window keeps
the same `usePlayback` instance. Every message is validated by
`parsePresenterMessage` before use.

`src/lib/core/windowChannel.ts` is transport only, payload-agnostic:
`BroadcastChannel(name)` where available; otherwise `localStorage` key
`docu-edit:channel:${name}` written as `{ seq, at, payload }` — the monotonic `seq`
guarantees the written value always differs, so the `storage` event (which fires
only in *other* same-origin windows, exactly the semantics needed) always fires;
otherwise `"none"`. It does not duplicate `core/editLock.ts`, whose Web Lock
provides exclusion, not messages.

### The presenter surface
Two columns. Left: `presenter-current-slide` (a `SlideStage` at the pane's scale,
`interactive: false`, `frameStyles` from its own `useBuildPreview()`), under it
`presenter-slide-counter` and `presenter-build-indicator`, `presenter-prev-button`
/ `presenter-next-button`. Right, top to bottom: `presenter-next-slide` (fully
built — no `frameStyles`), the timer row, `presenter-notes`. Keys, all classified
through `playbackActionForKey` first so navigation is identical to the audience
window: `g` grid, `b`/`.` blackout, `w`/`,` whiteout, `+`/`-` notes scale, `t`
timer pause/resume, `r` timer reset, Escape exits (Escape closes the grid first).
Modifier combos (Cmd/Ctrl/Alt) and events from `INPUT`/`TEXTAREA`/contenteditable
are ignored, as in `PresentationPlayView.onKeydown`.

### Displays, fullscreen and blanking
The Window Management API is not universally available and is permission-gated, so
it is used only as an enhancement: feature-detect `window.getScreenDetails`, ask
once, and on a granted permission place the audience window on the first
`isPrimary === false` screen before requesting fullscreen there. On absence,
refusal, a single screen, or any thrown error, the audience window opens as a plain
popup and shows the hint "Drag this window to your projector, then press F for full
screen". A browser that refuses the automatic `requestFullscreen()` leaves
`play-fullscreen-prompt` visible until a gesture grants it. Blanking is view state
only: it travels in `PresenterState.blank`, is never persisted, and clears on any
navigation action.

### What this never writes (both formats)
No model field, no reader, no writer, no serialiser and no preserved slot changes —
in `.pptx` (PresentationML) or `.odp` (ODF) alike. Specifically: rehearsed timings
are **not** recorded into `p:transition@advTm` or `presentation:duration` (SPEC-085
owns both attributes through `SlideTransition.advanceAfterMs`), no `p:timing` /
`anim:*` tree is touched (SPEC-086 owns those), and no notes part is rebuilt
(SPEC-083 owns notes writing). Because nothing is written, the six ways to break the
byte-identical no-edit download are untouched and the existing per-part `equalBytes`
assertions in `src/lib/presentation/roundTrip.spec.ts` keep passing unchanged.

### Tests
- `src/lib/core/windowChannel.spec.ts` — transport selection, `storage`-fallback
  posting/receiving with an injected storage, payload passthrough, `close()` idempotence.
- `src/lib/presentation/presenterSession.spec.ts` — `parsePresenterMessage` rejecting
  malformed/foreign-session/non-finite payloads, `isStaleState`, channel naming, the
  heartbeat/lost constants.
- `src/lib/presentation/presenterTimer.spec.ts` — start/pause/resume/reset accumulation
  with an injected `nowMs`, `formatElapsed` boundaries (`0:00`, `59:59`, `1:00:00`).
- `src/lib/presentation/playbackKeys.spec.ts` — extended for the new blackout/whiteout
  rows and to re-assert that every existing key still resolves to its current action.
- `src/views/presentation/usePlayback.spec.ts` — `setStep` clamping, `atEnd` handling,
  and that no existing advance/back/goTo behaviour changed.
- New `usePresenterSession.spec.ts` / `usePresenterView.spec.ts` — owner/follower roles,
  promotion on `bye` and on timeout, stale-`seq` rejection, intent forwarding, timer and
  notes-scale wiring, popup-blocked fallback.
- `src/views/PresenterView.size.spec.ts` (new) and `src/views/PresentationPlayView.size.spec.ts`
  (new); `PresentationPlayView.spec.ts` and `PresentationPlayView.transitions.spec.ts`
  re-run unchanged after the `usePlayInput` extraction — coverage is preserved, never weakened.
- `e2e/SPEC-100.spec.ts` per the last AC. No new fixture: `e2e/fixtures/mini-deck.pptx`
  already carries notes on slide 1 and `e2e/fixtures/spec086-animations.pptx` already
  carries builds, so `scripts/build-presentation-fixtures.ts` is unchanged.

### Specs this changes
- SPEC-084: `playbackKeys.ts` gains blackout/whiteout (its out of scope deferred "black/white
  blanking keys (B/W)" and "external-display awareness"; both land here), `usePlayback` gains
  `setStep`, and `/present/:n` gains the `?follow=` follower mode. Its navigation semantics are
  unchanged and its specs still pass as written.
- SPEC-088 must replay the new `PlaybackKeyAction` members or state that its export ignores
  them; the export stays single-window either way.
- `specs/000-roadmap.md` gains this spec's phase row and drops the backlog row "Presenter view
  — notes, next slide, timer | SPEC-083/084 both name it"; nothing else changes.
