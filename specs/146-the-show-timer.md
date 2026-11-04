---
id: SPEC-146
title: The show timer — stopwatch and countdown; start, pause, resume and stop
status: draft
area: shell
complexity: normal
story_points: 8
depends_on: [SPEC-066, SPEC-084, SPEC-100]
---
## Why
A presenter needs to know how long they have been talking, and how long they
have left. Most of them present from one screen. SPEC-100 put an elapsed-time
timer in presenter view (`src/lib/presentation/presenterTimer.ts`, wired by
`src/views/presentation/usePresenterTools.ts`). The full-screen show
(SPEC-084, `/slides/:id/present/:n`), which is what a single-screen presenter
actually uses, has no timer at all. The timer that does exist also has gaps:

- **It only counts up.** A speaker with a 20-minute slot has to do the
  subtraction in their head while talking. Keynote and PowerPoint both offer
  a countdown, and this product does not.
- **It can be paused and zeroed, but not stopped or started.** It starts from
  zero when the presenter route loads (`usePresenterView.ts:283`). After that
  there is one toggle, `presenter-timer-toggle`, whose label reads "Pause (T)"
  whether it is running or not and whose state is only `aria-pressed`. There
  is also `presenter-timer-reset` (`r`), which zeroes the time but leaves the
  timer running (`resetTimer`). You cannot hold it at 0:00 and start it when
  you actually begin speaking.
- **A reload zeroes it.** The state lives only in memory. An accidental
  Cmd+R ten minutes in, or a crashed tab, makes the next load a "show start"
  and the count begins again.
- **The audience window cannot control it, and is cut off from it.** The
  audience window forwards every other key it receives as an `intent`, but
  `t` and `r` there do nothing. If the presenter window closes, the audience
  promotes itself to a plain show (SPEC-100's role flip) and the elapsed time
  is lost with the window that held it.

This spec makes the timer a **show** timer instead of a presenter-view
feature. Every window that owns a show runs one, and a following window never
does. It has two modes:

- **Stopwatch** counts up from 0:00. It is the default, and it is what
  SPEC-100 shipped.
- **Countdown** counts down from a talk length. Past zero it keeps going as
  negative overtime, so the presenter can see how far over they are.

The timer starts automatically when the show starts, unless the new
preference turns that off. A key or a button starts, pauses, resumes or stops
it in either mode. One reload or crash does not lose it.

**This file is the normative source for the show timer:** its modes and
phases, what it displays, its controls and keys, when it starts, where it is
persisted, and the two additions it makes to SPEC-100's presenter session
protocol. SPEC-100 remains the normative source for the rest of that
protocol. Nothing here touches a file format, a model field, a reader or a
writer.

## Acceptance criteria
- [ ] `src/lib/presentation/presenterTimer.ts` stays pure and clock-injected. It gains the phase functions (`TimerPhase`, `timerPhase`, `toggleTimer`, `stopTimer`), the mode and length vocabulary (`TimerMode`, `TIMER_MODES`, `COUNTDOWN_MINUTES_DEFAULT = 20`, `COUNTDOWN_MINUTES_MIN = 1`, `COUNTDOWN_MINUTES_MAX = 240`, `isCountdownMinutes`, `parseCountdownMinutes`), the display function `timerReadout`, and `parseTimerState` and `showTimerRecordKey`, all per the Reference. The existing exports (`TimerState`, `IDLE_TIMER`, `isTimerRunning`, `startTimer`, `pauseTimer`, `resetTimer`, `elapsedMs`, `formatElapsed`) keep their behaviour, and `presenterTimer.spec.ts` keeps passing unchanged. `resetTimer` stays exported and tested even though no surface calls it any more.
- [ ] **One state, two readings.** `TimerState` always measures elapsed time, and the mode only changes what is displayed. Stopwatch shows `formatElapsed(elapsed)`. Countdown shows the time remaining of the talk length, rounded **up** to whole seconds, so it reads the full length at the start and `0:00` exactly when time is up. After that it shows overtime as `−M:SS`, with a U+2212 minus sign (`timerReadout`, which has a boundary table in the Reference). Because of this, switching mode or changing the length mid-show never resets or loses the count: the readout just recomputes. A length shorter than the time already elapsed goes straight into overtime.
- [ ] The timer half of `usePresenterTools.ts` moves into a new `src/views/presentation/useShowTimer.ts` (signature in the Reference), which both windows use. `TIMER_TICK_MS = 1000` moves with it. `usePresenterTools.ts` keeps only the notes scale. The display ticks once a second from the injected clock, and every operation and every mode or length change re-stamps `nowMs()`, so the result shows immediately, as it does today.
- [ ] **The show starts the timer.** `begin()` runs once per mounted view, at the moment the deck lands in a window that **owns** the show: `usePresenterView`'s first `loadPresentation` (it replaces `tools.startTimer()` at `usePresenterView.ts:283`), and `PresentationPlayView`'s first `loadPresentation` when it is not following, next to its `input.enterFullscreen()` call. `begin()` does not wait for a `fullscreenchange` event: SPEC-084 tolerates a refused fullscreen request and SPEC-100's audience prompt waits for a gesture, but the show is running either way. The view's second load pass (`initLibrary().then(loadPresentation)`) never restarts the timer.
- [ ] **What `begin()` does:** if this tab holds a record for this deck that parses (see "Persistence" in the Reference), the timer restores it exactly. A running record keeps counting, including the time the reload took. A paused or stopped record stays paused or stopped, and auto-start never overrides it. With no usable record, the timer starts at the stopped phase, and then begins running at `nowMs()` when `presentationTimerAutoStart` is on.
- [ ] **Manual control:** `toggle()` takes stopped to running from the start, running to paused (the elapsed time is banked), and paused to running (the paused stretch is not counted). `stop()` returns the timer to the stopped phase from either other phase, and does nothing when it is already stopped. A stopped timer reads `0:00` in Stopwatch mode and the full talk length in Countdown mode (for example `20:00`). Stop replaces SPEC-100's Reset in the UI: going back to the start now also halts the timer, and Start begins a fresh count.
- [ ] A new `src/components/presentation/ShowTimer.vue` renders the timer for both windows. A `testid-prefix` prop (`"presenter"` or `"play"`) produces `<prefix>-timer`, `<prefix>-timer-toggle`, `<prefix>-timer-stop` and `<prefix>-timer-status`.
  - **The readout** has `role="timer"`, `data-state="<phase>"`, `data-mode="<mode>"`, `data-overtime="true|false"` and tabular numerals. Its `aria-label` is "Elapsed time" in Stopwatch mode, "Time remaining" in Countdown mode, and "Time over" in overtime.
  - **Overtime** is shown by the minus sign, the label and colour together, never by colour alone (WCAG 1.4.1). It uses the `--danger` token family: `--text-danger` on the presenter surface, and on the HUD a `--danger` pill behind white text. Either way the text must keep a contrast of at least 4.5:1.
  - **The toggle** shows "Start (T)", "Pause (T)" or "Resume (T)", with accessible names "Start timer", "Pause timer" and "Resume timer". It has **no `aria-pressed`**: a button whose label changes with its state must not also be a toggle button (WAI-ARIA APG), so the label names the action the button will take next.
  - **Stop** shows "Stop (S)", with accessible name "Stop timer", and is `disabled` while the timer is stopped.
  - **`<prefix>-timer-status`** is a visually hidden polite live region. It announces "Timer started", "Timer paused", "Timer resumed" or "Timer stopped" after a key or click. It announces "Time is up" once, at the moment a running countdown reaches zero, whether a tick or a length change caused it. It never announces on `begin()`, a restore or an ordinary tick.
  - The component emits `toggle` and `stop` and holds no timer state.
- [ ] **Changing mode and length during a show:** with `editable` set (presenter view only), `ShowTimer` also shows the mode and length controls.
  - **Mode:** `role="group"` with `aria-label="Timer mode"`, holding two buttons, `presenter-timer-mode-stopwatch` ("Stopwatch") and `presenter-timer-mode-countdown` ("Countdown"). Each has `aria-pressed`, which is correct here because these labels never change.
  - **Length:** while in Countdown mode, `presenter-timer-length`, a minutes field (`type="text"`, `inputmode="numeric"`, labelled "Talk length (minutes)", copying the `preferences-grid-spacing` markup). It commits on `change`. `parseCountdownMinutes` rounds and clamps what is typed into 1–240, and anything that is not a number restores the current value.
  - Both controls write the **same per-device preferences** the dialog writes. There is no per-show copy.
  - Buttons were chosen for the mode switch, rather than radios or a `<select>`, because the show's key handler classifies arrow keys as navigation. A focused button only takes Enter and Space (the presenter's existing rule at `usePresenterView.ts:236`), and `playKeyHandler` already ignores text inputs. So none of these controls can also move the slide.
  - The full-screen show's HUD is read-only. A single-screen presenter sets the mode and length in Preferences.
- [ ] **Presenter view:** `PresenterView.vue`'s timer row (`PresenterView.vue:197–219`, `presenter-timer`, `presenter-timer-toggle`, `presenter-timer-reset`) becomes one `<ShowTimer testid-prefix="presenter" editable>`. Its `.presenter-timer*` styles move into the component. `presenter-timer-reset` is removed and `presenter-timer-stop` replaces it. The presenter's extra keys (`usePresenterView.ts:238–245`) become `t` for toggle and `s` and `r` for stop. `r` keeps working for SPEC-100 users and now means stop. `PresenterView.vue`'s `LINE_BUDGET` (400) never rises.
- [ ] **Full-screen show:** when `PresentationPlayView` owns the show (no `?follow=`, or after promotion), the HUD (`PresentationPlayView.vue:419–447`) shows `<ShowTimer testid-prefix="play">` (not editable) after `play-next`. It fades with the HUD on the existing 2000 ms idle timer, so the audience sees the timer only while the pointer moves. A following window renders no timer at all (`play-timer` count 0). A click on `play-timer-toggle` or `play-timer-stop` never reaches the stage's click-to-advance, just like `play-prev` and `play-next`, which use `@click.stop`.
- [ ] **Keys in the full-screen show:** `t` maps to `timerToggle`, and `s` and `r` map to `timerStop`, through `usePlayInput`'s `extraKey` hook. The timer keys are **not** added to `src/lib/presentation/playbackKeys.ts`, so SPEC-088's export key map is unchanged. As always, modifier combos and text inputs are ignored. A timer key also reveals the HUD for its usual 2000 ms so the presenter sees the new state: `usePlayInput` exports its existing `revealControls`. On the end screen the HUD is hidden, so the timer keys are not consumed there and fall through to SPEC-084's any-key exit, which stays as it is. A pointer click on a HUD timer button does not move focus to it, so the show's Space and Enter keep advancing. While a HUD timer button holds keyboard focus (Tab), Enter and Space activate only that button.
- [ ] **Protocol, intents:** `PresenterIntentAction` (and the `ACTIONS` list `parsePresenterMessage` validates against) gains `"timerToggle" | "timerStop"`. Both windows send their timer keys and buttons through the session's `perform`, so a follower forwards them as intents for free, with no new forwarding path. An **owner's** `perform` sends a timer action to a new `onTimerIntent(action)` dependency and **posts no `state`**. The reason: today `perform` always posts one fresh `state` (`usePresenterSession.ts:165`), and a follower that receives a fresh state with an unchanged step reads it as "the owner used up an input to complete a build" (`applyState` → `onRemoteAck` → `buildPreview.complete()`, `usePresenterSession.ts:202`). A timer key routed that way would cut short the build the audience is watching.
- [ ] **Protocol, timer state:** `PresenterMessage` gains `{ kind: "timer"; session: string; timer: TimerState }`, validated by `parseTimerState`. Anything with a wrong shape, a non-finite number or a negative number is dropped, never thrown. The owner posts `timer` after every change (`useShowTimer`'s `onChange`) and once more right after the `state` it posts to answer an audience `hello` (`usePresenterSession.ts:248`). It is never part of `state` and never takes a `seq`. `TimerState` is anchored to `lib/core/clock`'s wall clock, which every window on the machine shares, so it does not go stale and the heartbeat does not need to carry it. A follower keeps the last `timer` it accepted, exposed as `lastTimer`, and a `timer` message re-arms `PRESENTER_LOST_MS` the same way a `state` does. **Mode and length never cross the channel.** Both windows read the same per-device preferences, and `preferences.ts:587`'s `storage` listener carries a change made in one window to the other.
- [ ] **Promotion carries the timer:** when an audience window promotes itself (on `bye`, or after `PRESENTER_LOST_MS` of silence), its HUD timer appears and `adopt(lastTimer)` takes over the presenter's timer as it was: a running timer keeps running and a paused one stays at its value. It shows in the mode and length the preferences hold. A promoted window that never received a `timer` message calls `begin()`, like a fresh show.
- [ ] **Persistence:** an owning window writes its timer to `sessionStorage` under `showTimerRecordKey(deckId)` on every change. A following window never reads or writes the record. The record is removed when the view **unmounts**, which every in-app departure causes: Escape, the end screen, leaving fullscreen, browser Back, or a link. A document reload or a crashed tab does not unmount the app (`src/main.ts` installs no unload hook), so the record survives exactly those two events. A deck opened in a new tab starts with empty `sessionStorage`. Every storage access is wrapped in try/catch, so a browser with storage blocked or disabled just runs without persistence. The record holds elapsed time only, and mode and length come from preferences.
- [ ] **Preferences:** `src/lib/preferences.ts` gains three keys, each added in all three places (the `Preferences` interface, `PREFERENCE_DEFAULTS` and the validator table). The constants are imported from `./presentation/presenterTimer`, following `preferences.ts:33`'s import from `./spreadsheet/numberFormat/locale`.
  - `presentationTimerAutoStart: boolean`, default `true`, validated by `isBoolean` (`preferences.ts:302`).
  - `presentationTimerMode: TimerMode`, default `"stopwatch"`, validated by `oneOf<"presentationTimerMode">(TIMER_MODES)` (`preferences.ts:306`).
  - `presentationTimerCountdownMinutes: number`, default `COUNTDOWN_MINUTES_DEFAULT`, validated by `isCountdownMinutes` in the inline numeric form `gridSpacingCm` uses (`preferences.ts:381`). It is **never `inSet`** (`preferences.ts:314`), which is string-only.
- [ ] **Preferences dialog:** `PreferencesDialog.vue` gains a "Presentations" section (`aria-labelledby="preferences-presentations-heading"`, an existing `ICON_PATHS` glyph, no new icon), placed before "Spreadsheets". It holds three fields:
  - `preferences-presentation-timer-autostart`: a checkbox labelled "Start the timer when a show starts", copying the `snapToGuides` markup, with the hint "Presenter view and the full-screen show. Off: the timer waits until you press Start (T). Default: on".
  - `preferences-presentation-timer-mode`: a `UiSelect` labelled "Timer", with "Stopwatch — counts up from 0:00" and "Countdown — counts down from the talk length", and the hint "Default: Stopwatch".
  - `preferences-presentation-timer-length`: labelled "Talk length (minutes)", copying the `preferences-grid-spacing` markup and committing through `parseCountdownMinutes`, with the hint "1–240. Used by Countdown. Default: 20".

  "Reset to defaults" covers all three keys through `PREFERENCE_DEFAULTS`, with no extra code.
- [ ] **Sizes ratchet:** `PresentationPlayView.vue` is 592 lines against `LINE_BUDGET = 600` (`PresentationPlayView.size.spec.ts`), which never rises. Room is made by moving its auto-advance block (`PresentationPlayView.vue:120–140`) and the identical one in `usePresenterView.ts:192–208` into a new shared `src/views/presentation/useAutoAdvance.ts` (`useAutoAdvance({ currentSlide, atEnd, enabled: () => boolean, advance })` → `{ schedule, clear }`, because both views call these directly on load, promotion, exit and unmount). That block has the same behaviour in both views, apart from the play view's `!presenter.following.value` guard, which becomes `enabled`. After that move the view gains only its `useShowTimer` wiring and the HUD markup. Every touched `use*.ts` stays within `src/composableSize.spec.ts`'s `DEFAULT_BUDGET` of 400.
- [ ] Presenting stays inert: no timer change sets a dirty flag, calls `markSlideEdited`, takes the edit lock or writes the deck. Timer state, mode and length are never written to `.pptx` or `.odp` (see "What this never writes"), and a deck that has been presented and exited still downloads byte-identical part for part (`equalBytes`). That is the SPEC-100 assertion, and it keeps passing unchanged.
- [ ] e2e (`e2e/SPEC-146.spec.ts`; time is driven by `page.clock.setFixedTime`, whose precedent is `e2e/SPEC-120.spec.ts:50`, and no timer assertion waits on the real clock), per "Tests" in the Reference:
  - **Stopwatch in the full-screen show:** the timer starts running at 0:00 → reads 1:15 → `t` pauses it and the paused time is not counted, while ArrowRight still advances → `play-timer-stop` zeroes it without advancing → `t` starts it again → a reload restores the timer → Escape and Present again starts a fresh 0:00.
  - **Countdown in the full-screen show:** set in Preferences, the timer reads 2:00 → 0:45 → −0:42 in overtime, and the status announces "Time is up" → Stop returns it to 2:00.
  - **Presenter view:** `t` pressed in the audience window pauses the presenter's timer, and the audience has no timer of its own. Switching to Countdown mid-show shows the remaining time without losing the count. When the presenter window closes, the promoted audience shows the inherited paused value.
  - **Auto-start off:** Present leaves the timer stopped, with "Start (T)".

## Out of scope
- **Warnings before zero** (for example a colour change at five minutes
  left), sounds or vibration at zero, and a per-slide time readout. These are
  the named follow-ups, and every one of them reads `timerReadout` without
  changing `TimerState`.
- **A talk length saved per deck.** The length is per device. Saving it per
  deck needs a library-record field, and saving it in the file is ruled out
  below.
- **A key to switch modes.** Modes are switched with the presenter-view
  buttons or in Preferences.
- **Rehearse timings, or writing the timer anywhere in the file.**
  `p:transition@advTm` and `presentation:duration` belong to SPEC-085's
  `advanceAfterMs`. SPEC-100 already ruled this out.
- **A timer that stays visible on the full-screen show.** It appears only
  with the HUD, so the audience does not watch the clock.
- **Editing mode or length from the HUD.** The HUD is in front of the
  audience.
- **A timer on SPEC-100's "Present here" full-screen pane.** The pane
  covers the presenter surface by design.
- **A timer in SPEC-088's HTML export.** The export stays without one, and
  its key map does not change.
- **Sharing one timer across tabs or devices beyond the presenter/audience
  pair.** The pair already shares one through the channel.
- **Keeping a timer through closing and reopening the tab.**
  `sessionStorage` ends with the tab, by design.

## Reference

### Dependencies restated
- `specs/100-presenter-view.md` owns the presenter session protocol in
  `src/lib/presentation/presenterSession.ts`:
  ```ts
  interface PresenterState { seq: number; slideIndex: number; buildIndex: number; atEnd: boolean; blank: "none" | "black" | "white" }
  type PresenterIntentAction = "advance" | "back" | "first" | "last" | "goTo" | "blackout" | "whiteout" | "exit";
  type PresenterMessage =
    | { kind: "state";  session: string; state: PresenterState }
    | { kind: "hello";  session: string; from: "presenter" | "audience" }
    | { kind: "intent"; session: string; action: PresenterIntentAction; slideIndex?: number }
    | { kind: "bye";    session: string; from: "presenter" | "audience" };
  PRESENTER_HEARTBEAT_MS = 1000; PRESENTER_LOST_MS = 3000;
  parsePresenterMessage(raw: unknown, session?: string): PresenterMessage | null  // never throws
  ```
  The owner posts a fresh `state` on every change. The heartbeat re-posts the
  last one under the same `seq`. A follower rejects any `seq` that is not
  newer (`isStaleState`). A fresh state whose step and blank are unchanged
  means "complete your in-flight build" (`onRemoteAck`). `bye`, or
  `PRESENTER_LOST_MS` of silence, promotes a follower to owner on the same
  `usePlayback` machine (`onPromote`).
- `src/views/presentation/usePresenterSession.ts` (SPEC-100): `perform(action, slideIndex?)` is
  every window's single input entry point. A follower sends it as an
  `intent`. An owner runs it and then posts one fresh `state`. The function
  also returns `role`, `following`, `isAudience`, `blank`, `reapply` and
  `leave`.
- `specs/084-slideshow-playback.md` owns the full-screen show:
  `PresentationPlayView.vue`, `usePlayInput.ts` (`playKeyHandler` with its
  `extraKey` hook, which runs before Escape, before the end screen's any-key
  exit and before the shared map; `isTextInput`, which skips INPUT, TEXTAREA
  and contenteditable targets; `f` for fullscreen; the 2000 ms idle HUD
  through `revealControls`/`controlsVisible`), and `playbackKeys.ts`, the key
  map that SPEC-088's export also uses.
- `specs/066-preferences.md` owns `src/lib/preferences.ts`. A key is declared
  in three places: the interface, `PREFERENCE_DEFAULTS` and the validator
  table. `Validator<K> = (raw: unknown) => Preferences[K] | undefined`.
  `isBoolean` is the boolean validator. `oneOf<K>(values)` is the
  value-typed helper. `inSet(raw, values: readonly string[])` is string-only
  and returns `string`. The `storage` listener at `preferences.ts:587`
  re-reads the preferences and notifies listeners when another window writes
  them. `PreferencesDialog.vue` renders the dialog, and
  `usePreferences()` → `{ preferences, setPreference }` is the reactive
  access.
- `src/lib/core/clock.ts`: `now(): Date` is the single wall-clock seam. Unit
  tests pin it with `setClock`. Both views default `nowMs` to
  `() => now().getTime()`. It reads the OS wall clock, which every window on
  the machine shares. `performance.now()` is the wrong clock here, because
  each document has its own time origin.
- Round-trip relation: `equalBytes` from `src/lib/core/equality.ts`, compared
  part by part after `unzip`, as in `src/lib/presentation/roundTrip.spec.ts`.
  This spec writes nothing, so it needs no new relation.

### The timer (normative)
```ts
// src/lib/presentation/presenterTimer.ts — additions
export type TimerPhase = "stopped" | "running" | "paused";
/** running ⇔ startedAtMs !== null; stopped ⇔ not running and accumulatedMs === 0; else paused. */
export function timerPhase(state: TimerState): TimerPhase;
/** stopped → running from 0 at nowMs; running → paused (banked); paused → running. */
export function toggleTimer(state: TimerState, nowMs: number): TimerState;
/** Any phase → IDLE_TIMER (a fresh object). */
export function stopTimer(): TimerState;

export type TimerMode = "stopwatch" | "countdown";
export const TIMER_MODES: readonly TimerMode[] = ["stopwatch", "countdown"];
export const COUNTDOWN_MINUTES_DEFAULT = 20;
export const COUNTDOWN_MINUTES_MIN = 1;
export const COUNTDOWN_MINUTES_MAX = 240;
/** A whole number of minutes in [MIN, MAX] (the preference validator). */
export function isCountdownMinutes(raw: unknown): raw is number;
/** What a person typed → whole minutes, rounded and clamped into [MIN, MAX];
 * "", whitespace, or anything Number() rejects → null (the field restores). */
export function parseCountdownMinutes(text: string): number | null;

export interface TimerReadout { text: string; overtime: boolean }
/** What the timer shows at nowMs. stopwatch: formatElapsed(elapsed), never
 * overtime. countdown: r = lengthMs − elapsed; s = Math.ceil(r / 1000);
 * s > 0 → formatElapsed(s·1000); s ≤ 0 → "0:00" when s is 0 (or −0), else
 * "−" (U+2212) + formatElapsed(−s·1000); overtime ⇔ r ≤ 0. */
export function timerReadout(state: TimerState, nowMs: number, mode: TimerMode, lengthMs: number): TimerReadout;

/** Validates a TimerState from another window or from storage: startedAtMs is
 * null or a finite number ≥ 0; accumulatedMs a finite number ≥ 0; extra fields
 * dropped; anything else → null. Never throws. */
export function parseTimerState(raw: unknown): TimerState | null;
/** "docu-edit:show-timer:<deckId>" */
export function showTimerRecordKey(deckId: string): string;
```

`timerReadout` boundaries for countdown, length 2 minutes (`lengthMs = 120_000`):

| Elapsed (ms) | Text | Overtime |
| --- | --- | --- |
| 0 | `2:00` | no |
| 1 | `2:00` | no |
| 1_000 | `1:59` | no |
| 75_000 | `0:45` | no |
| 119_999 | `0:01` | no |
| 120_000 | `0:00` | yes |
| 120_999 | `0:00` | yes |
| 121_000 | `−0:01` | yes |
| 162_000 | `−0:42` | yes |
| 120_000 + 3_661_000 | `−1:01:01` | yes |

With a 240-minute length and nothing elapsed, the readout is `4:00:00`.

Phase table (`t` is the toggle, `s`/`r` is stop; "start" means 0:00 in
Stopwatch mode and the full length in Countdown mode):

| Phase | Readout | Toggle shows | `t` → | `s`/`r` → | Stop button |
| --- | --- | --- | --- | --- | --- |
| stopped | the start | Start (T) | running from the start | (nothing) | disabled |
| running | ticking | Pause (T) | paused | stopped | enabled |
| paused | frozen | Resume (T) | running, gap not counted | stopped | enabled |

Pausing and starting again within the same millisecond banks 0 ms and reads
as stopped. That is correct, because both phases show the start.

```ts
// src/views/presentation/useShowTimer.ts
export const TIMER_TICK_MS = 1000;
export interface ShowTimerDeps {
  deckId: Ref<string>;
  nowMs?: () => number;                       // default () => now().getTime()
  autoStart?: () => boolean;                  // default: presentationTimerAutoStart
  mode?: () => TimerMode;                     // default: presentationTimerMode
  lengthMinutes?: () => number;               // default: presentationTimerCountdownMinutes
  storage?: Pick<Storage, "getItem" | "setItem" | "removeItem"> | null; // default window.sessionStorage (guarded); null = none
  onChange?: (state: TimerState) => void;     // after begin/toggle/stop/adopt — the session posts `timer`
}
export function useShowTimer(deps: ShowTimerDeps): {
  timer: Readonly<Ref<TimerState>>;
  phase: ComputedRef<TimerPhase>;
  mode: ComputedRef<TimerMode>;
  lengthMinutes: ComputedRef<number>;
  text: ComputedRef<string>;                  // timerReadout(...).text at the last tick
  overtime: ComputedRef<boolean>;
  status: Readonly<Ref<string>>;              // the ShowTimer live-region text
  begin(): void;                              // once per mount; see "When it starts"
  adopt(state: TimerState): void;             // a promoted window takes the presenter's timer
  toggle(): void;
  stop(): void;
  setMode(mode: TimerMode): void;             // writes the preference; re-stamps the tick
  setLengthMinutes(minutes: number): void;    // clamps, writes the preference; re-stamps the tick
};
```
Owning windows write the record on every change. On unmount, the composable
clears its tick interval and removes the record. "Time is up" is set in
`status` when `overtime` turns from false to true while the timer is
running. It is never set by `begin()`, `adopt()` or a restore that lands
already in overtime.

### When it starts
"The show starts" means the deck has loaded in a window that owns the show
and its first slide is on screen. That is the point SPEC-100 starts the timer
today, and the point the play view asks for fullscreen. Owning windows: the
presenter window always; the play view when it has no `?follow=`, and a
follower from the moment it is promoted. `begin()` handles a reload by
restoring the record, and the auto-start preference applies only when there
is no record. A follower never calls `begin()` while following. The mode does
not affect when the timer starts: a countdown starts with the show just as a
stopwatch does.

### Protocol additions (to SPEC-100's contract)
```ts
type PresenterIntentAction = /* SPEC-100's eight */ | "timerToggle" | "timerStop";
type PresenterMessage = /* SPEC-100's four */ | { kind: "timer"; session: string; timer: TimerState };
```

| Who | Event | Does |
| --- | --- | --- |
| Owner | `perform("timerToggle" \| "timerStop")`, from a local key or button or from a remote `intent` | calls `deps.onTimerIntent(action)`; posts **no** `state` and bumps no `seq` |
| Owner | its timer changed (`onChange`) | posts `{ kind: "timer", session, timer }` (does nothing with no channel) |
| Owner | audience `hello` | posts the last `state` (as today), then one `timer` |
| Follower | `perform("timerToggle" \| "timerStop")` | sends the `intent` (the existing follower path) |
| Follower | `timer` received | stores it as `lastTimer`; re-arms `PRESENTER_LOST_MS` |
| Follower | promoted | the view calls `adopt(lastTimer)` if it has one, else `begin()` |

`PresenterSessionDeps` gains `onTimerIntent?: (action: "timerToggle" | "timerStop") => void`.
The session's return value gains `postTimer(timer: TimerState): void` and
`lastTimer: Readonly<Ref<TimerState | null>>`. `run()` never sees a timer
action. The blank and navigation rules stay as SPEC-100 wrote them. In
particular, a timer key does not clear a blanked screen, because it is not
navigation. Mode and length are not protocol: a mode switch in the presenter
window writes the preference, and the audience window picks it up through
the `storage` listener. It needs that only once it has been promoted, since
a follower displays no timer.

### Keys
| Key | Presenter window | Full-screen show (owner) | Audience window (follower) |
| --- | --- | --- | --- |
| `t` | toggle | toggle, reveals the HUD | forwarded as `timerToggle` |
| `s`, `r` | stop | stop, reveals the HUD | forwarded as `timerStop` |

None of these keys belongs to `playbackKeys.ts` or collides with it. The
shared map holds `n`, `p`, `b`, `.`, `w`, `,`, Home, End and the arrows. The
play view's `f`, the digit jumps and the presenter's `g`, `+`, `=` and `-`
are unaffected. In the presenter window the timer keys also work at the end
of the show, because its surface stays visible, as today. In the full-screen
show the timer keys fall through on the end screen to its any-key exit.
Typing in `presenter-timer-length` never reaches the show:
`playKeyHandler` skips text inputs, so digits, `t`, `s` and `r` typed there
stay in the field.

### Persistence
Key `showTimerRecordKey(deckId)`, in `window.sessionStorage`, value
`JSON.stringify(TimerState)`, holding elapsed time only. On read, the value
goes through `JSON.parse` and `parseTimerState`, and anything unparsable
counts as no record. `sessionStorage` is right for this, and `localStorage`
would be wrong:

| Event | Record | Result |
| --- | --- | --- |
| Reload or crash of the owning tab | survives (no unmount) | restored; a running timer counts the gap |
| Escape, end screen, leaving fullscreen, Back, a link | removed on unmount | the next Present starts fresh |
| A deck opened in a new tab or window | absent | fresh |
| A different deck in the same tab | different key | fresh |
| `sessionStorage` blocked or throwing | treated as absent; writes are skipped | the timer works, with no persistence |

A browser may copy the opener's `sessionStorage` into the audience popup.
That copy is harmless, because a follower never reads or writes the record
and a promoted follower adopts the channel's `lastTimer` first.

### Surfaces
`ShowTimer.vue` lays out one row: the readout, the toggle and Stop. With
`editable` set it adds a second row holding the mode group and, in Countdown
mode only, the length field.

- **`presenter` prefix:** renders with `UiButton size="sm"`, matching
  today's row. Overtime colours the readout `--text-danger`.
- **`play` prefix:** renders with the HUD's pill styling: `.hud-pill`'s
  colours, `color-mix(in srgb, #000 55%, transparent)`, `--radius-sm` and
  `tabular-nums`. Overtime turns the pill's background to `--danger`
  behind white text. In the play view it is placed in `.play-hud` after
  `play-next`.

Visible text never depends on the HUD's opacity. The HUD fades with
`opacity` and not `visibility`, so the status live region still announces
"Time is up" while the HUD is faded.

### What this never writes (both formats)
No model field, reader, writer, serialiser or preserved slot changes, in
`.pptx` or `.odp`. No `p:transition@advTm`, `presentation:duration`,
`p:timing` or `anim:*` is touched. The timer state exists in memory, in the
channel and in the tab's `sessionStorage`. The mode and length exist in the
device's preferences. Neither is stored anywhere else.

### Tests
- `src/lib/presentation/presenterTimer.spec.ts`: new cases for
  - `timerPhase` in all three phases, including a 0 ms pause.
  - `toggleTimer` through stopped → running → paused → running, and
    `stopTimer` from each phase.
  - `timerReadout`: every row of the boundary table above, plus stopwatch
    (never overtime) and the 240-minute `4:00:00`.
  - `isCountdownMinutes`: 1 and 240 pass; 0, 241, 2.5, `NaN` and `"20"`
    fail.
  - `parseCountdownMinutes`: `"25"` → 25, `" 7 "` → 7, `"2.6"` → 3, `"0"` →
    1, `"999"` → 240, and `""`, `"abc"` → `null`.
  - `parseTimerState`: accepts a `null` start; rejects `NaN`, `Infinity`,
    negative numbers, strings, missing fields and non-objects; drops extra
    fields.
  - `showTimerRecordKey`.

  Every existing case stays unchanged.
- `src/lib/presentation/presenterSession.spec.ts`: `timer` messages parse,
  and malformed ones, ones from another session and ones with a non-finite
  number are dropped. `timerToggle` and `timerStop` intents parse.
- `src/views/presentation/usePresenterSession.spec.ts`: an owner's
  `perform("timerToggle")` calls `onTimerIntent` and posts nothing with a
  new `seq`. So a follower that is mid-build when its own timer intent comes
  back gets **no** `onRemoteAck`; count the calls. A follower's `perform`
  forwards the intent. `timer` sets `lastTimer` and re-arms the lost timer.
  `hello` is answered with `state` and then `timer`. `postTimer` does nothing
  as a follower and with no channel.
- New `src/views/presentation/useShowTimer.spec.ts`:
  - `begin()` with auto-start on and with it off.
  - `begin()` is idempotent per mount.
  - A running record is restored and counts the gap. Paused and stopped
    records are restored without auto-start overriding them. A corrupt
    record falls back to fresh.
  - The record is written on every change and removed on unmount. A
    throwing storage is tolerated, and `null` storage persists nothing.
  - `adopt`, and `status` announcements for each transition (none on
    `begin`).
  - "Time is up" is announced exactly once when a running countdown crosses
    zero on a tick, and once when `setLengthMinutes` drops the length below
    the elapsed time. It is not announced when a record restores into
    overtime, or while the timer is paused.
  - `setMode` and `setLengthMinutes` write the preferences and leave
    `timer` untouched (the elapsed time is kept).
  - The tick test moves here from `usePresenterView.spec.ts` with its three
    assertions intact: "0:00" before the tick, "1:01" after one
    `TIMER_TICK_MS`, "1:01:01" an hour later.
- `src/views/presentation/usePresenterView.spec.ts`: the test "starts the
  timer with the show; t/the toggle pause and resume it (aria-pressed);
  r/reset zero it" is rewritten for this contract, and every check keeps a
  successor:
  - 0:00 running at show start, with the label "Pause (T)" in place of
    `aria-pressed="false"`.
  - The click pauses at 1:15, with "Resume (T)" in place of `aria-pressed`.
  - `t` resumes and the paused stretch is not counted.
  - `r` stops at 0:00, then `presenter-timer-stop` is disabled while
    stopped, and `s` from running stops it.
  - New: auto-start off begins stopped with "Start (T)".
  - New: an audience `timerToggle` intent toggles the presenter's timer.
  - New: the mode buttons and the length field. Their `aria-pressed`
    values are checked. ArrowRight while a mode button is focused still
    advances the slide exactly once and does not change the mode. Typing
    `t` into the length field neither toggles the timer nor advances.
- `src/views/PresentationPlayView.spec.ts`:
  - The HUD timer exists when the window owns the show and is absent while
    following. It renders no mode or length controls.
  - It appears on promotion with the adopted value.
  - `t`, `s` and `r` work, and a timer key reveals the HUD.
  - A toggle click does not advance.
  - On the end screen, `t` exits.
  - A follower's `t` sends `timerToggle`.
  - Countdown and overtime render from the preferences.
- `src/components/PreferencesDialog.spec.ts`: the three new fields reflect
  and set their keys, and the length field commits through
  `parseCountdownMinutes`. `src/lib/preferences.spec.ts`: the defaults are
  `true`, `"stopwatch"` and `20`, and an invalid stored value of each falls
  back to its default.
- `src/views/PresenterView.size.spec.ts` and
  `src/views/PresentationPlayView.size.spec.ts` keep passing at their
  current budgets. `PresentationPlayView.transitions.spec.ts` and
  `PresentationPlayView.follow.spec.ts` keep passing unchanged after the
  `useAutoAdvance` extraction.
- `e2e/SPEC-146.spec.ts`. All tests open `e2e/fixtures/mini-deck.pptx` (4
  slides) the way `e2e/SPEC-100.spec.ts`'s `openFixture` does, and pin the
  clock with `page.clock.setFixedTime(t0)` before `page.goto`. The clock is
  per context, so it covers the popup too. Every later "time passes" step is
  another `setFixedTime`, and the auto-waiting `toHaveText` covers the next
  one-second tick. The minus sign in expected texts is U+2212.
  1. **Stopwatch, full-screen show:** `slides-present` → move the mouse →
     `play-timer` reads `0:00`, `data-state="running"`,
     `data-mode="stopwatch"`, and the toggle reads "Pause (T)". Advance the
     clock 75 s → `1:15`. Press `t` → `paused`, "Resume (T)". Advance 10 min
     → still `1:15`. ArrowRight → `play-slide-counter` reads `2 / 4` and the
     timer is still `1:15`. Click `play-timer-stop` → `0:00` `stopped`,
     "Start (T)", stop disabled, the counter still `2 / 4`. Press `t` →
     `running` from `0:00`. Advance 30 s and `page.reload()` → `0:30`
     `running`. Escape → `slides-present` again → `0:00` `running`.
  2. **Countdown, full-screen show:** open Preferences → set
     `preferences-presentation-timer-mode` to Countdown and
     `preferences-presentation-timer-length` to `2` → close →
     `slides-present` → `play-timer` reads `2:00` `running`
     `data-mode="countdown"`. Advance 75 s → `0:45`, `data-overtime="false"`.
     Advance to 162 s → `−0:42`, `data-overtime="true"`, and
     `play-timer-status` reads "Time is up". Press `s` → `2:00` `stopped`.
  3. **Presenter view:** use `startPresenterView` (copied from
     `SPEC-100.spec.ts`) → `presenter-timer` `running` in Stopwatch mode.
     Advance 75 s → `1:15`. Click `presenter-timer-mode-countdown` → `18:45`
     (the default 20 minutes), still `running`. Fill `presenter-timer-length`
     with `1` and press Enter → `−0:15`, `data-overtime="true"`. Click
     `presenter-timer-mode-stopwatch` → `1:15`. Press `t` **in the audience
     window** → `presenter-timer` `paused` at `1:15`, and the audience has
     `play-timer` count 0. Close the presenter page → within
     `PRESENTER_LOST_MS` plus a margin, the audience (now on `/present/n`
     with no `?follow=`) shows `play-timer` at `1:15` `paused` once the mouse
     moves.
  4. **Auto-start off:** open Preferences → untick
     `preferences-presentation-timer-autostart` → close → `slides-present` →
     the timer reads `0:00` `stopped` with "Start (T)". Press `t` →
     `running`.

  No new fixture is needed, so `scripts/build-presentation-fixtures.ts` does
  not change.

### Specs this changes
- **SPEC-100:** the timer's controls change from a Pause toggle with
  `aria-pressed` plus Reset to Start, Pause or Resume plus Stop.
  `presenter-timer-reset` becomes `presenter-timer-stop`. `r` stops the timer
  instead of resetting it while it keeps running. The timer gains a
  Countdown mode, with mode and length controls on the presenter surface,
  and it survives a reload. The protocol gains the `timer` message and two
  intent actions, and `usePresenterTools` loses its timer half. Its notes,
  grid, blanking, window pair and failure modes are unchanged, and its e2e
  (`presenter-timer` matching `/^\d+:\d\d$/` with the default Stopwatch
  mode) passes as written.
- **SPEC-084:** the HUD gains the timer while the window owns the show.
  `t`, `s` and `r` are taken by the view through `extraKey` everywhere except
  the end screen. The auto-advance block moves to `useAutoAdvance.ts` with
  its behaviour unchanged. Navigation semantics and `playbackKeys.ts` are
  untouched.
- **SPEC-066:** three new keys (`presentationTimerAutoStart`,
  `presentationTimerMode`, `presentationTimerCountdownMinutes`) and a
  "Presentations" section in the dialog.
- `specs/000-roadmap.md` gains this spec's phase row. Nothing else there
  changes.
