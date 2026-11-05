<script setup lang="ts">
// The slideshow playback surface (SPEC-084): a full-screen, chrome-free page
// at /slides/:id/present/:n that walks the deck slide by slide. No AppShell,
// toolbar or filmstrip — just the SPEC-079 stage on black, letterboxed and
// scaled to fill the viewport, driven by the usePlayback step machine and the
// shared playbackKeys map. This spec is the normative source for playback
// navigation; SPEC-085/086 layer transitions/builds onto the same pipeline and
// SPEC-088 reimplements it for the HTML export. Runtime declarations, not
// defineProps<{...}>() generics — espree lints .vue scripts, so typed logic
// (the step machine, the key map) lives in the imported .ts modules. SPEC-100:
// with ?follow= this is the presenter view's audience window (usePresenterSession).
// SPEC-146: a window that owns the show runs the show timer (useShowTimer) in its HUD.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import SlideStage from "../components/presentation/SlideStage.vue";
import { stageFields } from "../components/presentation/slideRender";
import { now } from "../lib/clock";
import SlideTransitionLayer from "../components/presentation/SlideTransitionLayer.vue";
import { emptyTwoStageView, transitionAnimates } from "../components/presentation/transitionStyles";
import UiButton from "../components/ui/UiButton.vue";
import ShowTimer from "../components/presentation/ShowTimer.vue";
import { documentKindOf, getPresentation, initLibrary, refreshDocumentFromStorage, storePresentation } from "../lib/documentStore";
import { isLegacyPresentationModel } from "../lib/presentation/model";
import { upgradePresentationModel } from "../lib/presentation/open";
import { cmToPx } from "../lib/core/units";
import { slideOutline } from "../lib/presentation/slideModel";
import { slideBuildCount } from "../lib/presentation/buildPlayback";
import { usePlayback } from "./presentation/usePlayback";
import { useAutoAdvance } from "./presentation/useAutoAdvance";
import { showTimerKeys, useShowTimer } from "./presentation/useShowTimer";
import { fullyBuiltStyles, useBuildPreview, watchBuildSteps } from "./presentation/useBuildPreview";
import { useShowingMedia } from "./presentation/useShowingMedia";
import { usePlayInput } from "./presentation/usePlayInput";
import { usePresenterSession } from "./presentation/usePresenterSession";
import { useReducedMotion } from "./presentation/useReducedMotion";

const route = useRoute();
const router = useRouter();

const presentationId = computed(() => String(route.params.id ?? ""));

function initialModel() {
  const model = getPresentation(presentationId.value) ?? null;
  return model && isLegacyPresentationModel(model) ? null : model;
}

const presentation = shallowRef(initialModel());
const loaded = ref(false);
const slideCount = computed(() => presentation.value?.slides.length ?? 0);

// --- reduced motion (read from this window's media query, reactively) ------
const reducedMotion = useReducedMotion();

// --- the step machine -------------------------------------------------------
function routeStartIndex() {
  const n = Number(route.params.n ?? 1);
  return Number.isFinite(n) ? n - 1 : 0; // usePlayback clamps out-of-range
}
function buildCountAt(index = 0) {
  const slide = presentation.value?.slides[index];
  return slide ? slideBuildCount(slide) : 0;
}
// SPEC-104: hidden slides are stepped over by advance/back/goTo — never by
// the deep link itself (see loadPresentation).
function isHidden(index = 0) {
  return presentation.value?.slides[index]?.hidden === true;
}
const playback = usePlayback({ slideCount, reducedMotion, start: routeStartIndex(), buildCountAt, isHidden });
const currentSlide = computed(() => presentation.value?.slides[playback.current.value.slideIndex] ?? null);
// SPEC-093: the playing slide's resolved deck field band.
const currentFields = computed(() => {
  const model = presentation.value;
  return model && !playback.atEnd.value ? stageFields(model, playback.current.value.slideIndex, now()) : [];
});

// --- build playback (SPEC-086): frames hidden until their build plays --------
const buildPreview = useBuildPreview(() => presentation.value && { widthCm: presentation.value.slideWidthCm, heightCm: presentation.value.slideHeightCm });
watchBuildSteps(buildPreview, () => playback.current.value, () => currentSlide.value, () => reducedMotion.value);

// --- presenter session (SPEC-100): ?follow= makes this a follower -----------
const showTimer = useShowTimer({ deckId: presentationId }); // begun only while owning (SPEC-146)
const presenter = usePresenterSession({
  window: "audience",
  deckId: presentationId,
  session: String(route.query.follow ?? ""),
  playback,
  slideCount,
  buildCountAt,
  isHidden,
  navigation: { advance: advanceStep, back: backStep, goTo: jumpTo, exit },
  onRemoteStep: (fromIndex, forward) => {
    clearTransition();
    if (forward) playIncomingTransition(presentation.value?.slides[fromIndex] ?? null);
  },
  onRemoteAck: () => {
    clearTransition();
    buildPreview.complete();
  },
  // A follower never upgrades or stores: it re-reads the store once, after
  // the owner's first state (the owner may have just stored an upgrade).
  onFirstState: () => void refreshDocumentFromStorage(presentationId.value).then(loadPresentation),
  onPromote: () => {
    if (presenter.lastTimer.value) showTimer.adopt(presenter.lastTimer.value); else showTimer.begin();
    if (!presentation.value) return void loadPresentation(); // a legacy deck waited on the owner
    syncUrl(); // drops ?follow=
    autoAdvance.schedule();
  },
  onTimerIntent: (action) => (action === "timerToggle" ? showTimer.toggle() : showTimer.stop()),
});
// --- per-showing GIF URLs: a GIF restarts each time its slide is shown -------
const showingMedia = useShowingMedia();
watch(
  [() => playback.current.value.slideIndex, () => playback.atEnd.value, presentation],
  () => showingMedia.enter(presentation.value?.media ?? []),
  { immediate: true },
);
const deckTitle = computed(() => presentation.value?.title ?? "Slide show");
const counterLabel = computed(() => `${playback.current.value.slideIndex + 1} / ${slideCount.value}`);
const announce = computed(() => {
  if (presenter.blank.value !== "none") return "Screen blanked";
  if (!presentation.value || playback.atEnd.value) return "";
  const outline = currentSlide.value ? slideOutline(currentSlide.value) : { title: "" };
  const title = outline.title && outline.title.length > 0 ? outline.title : "untitled";
  return `Slide ${playback.current.value.slideIndex + 1} of ${slideCount.value} — ${title}`;
});

// --- auto-advance (AC-17): a slide with advanceAfterMs schedules a forward
// step; any navigation (which changes the index) or exit reschedules/clears it.
// Declared before the load invocation below so the setup-time load can call it.
const autoAdvance = useAutoAdvance({ currentSlide, atEnd: playback.atEnd, enabled: () => !presenter.following.value, advance: advanceStep });
watch([() => playback.current.value.slideIndex, () => playback.atEnd.value], autoAdvance.schedule);

// --- load (mirrors PresentationView: text id redirects, legacy upgrades) ----
async function loadPresentation() {
  if (documentKindOf(presentationId.value) === "text") {
    void router.replace(`/doc/${presentationId.value}`);
    return;
  }
  let model = getPresentation(presentationId.value) ?? null;
  if (model && isLegacyPresentationModel(model)) {
    if (presenter.following.value) return; // the owner upgrades; see onFirstState
    try {
      const upgraded = await upgradePresentationModel(model);
      storePresentation(upgraded);
      model = getPresentation(presentationId.value) ?? upgraded;
    } catch {
      // Reading stays tolerant — keep the record as-is rather than blanking.
    }
  }
  presentation.value = model;
  loaded.value = true;
  if (model) {
    // Clamp the deep-link n into range and land exactly on it: a deep link to
    // a hidden slide still shows it (SPEC-104), which goTo would step over.
    playback.setStep({ slideIndex: routeStartIndex(), buildIndex: 0 });
    presenter.reapply(); // a follower re-lands on the owner's last state
    syncUrl(); // reflect the clamp even when the index did not change (AC-5)
    if (currentSlide.value) {
      buildPreview.settle(currentSlide.value, playback.current.value.buildIndex); // hide unbuilt frames (AC-20)
    }
    autoAdvance.schedule(); // the starting slide may itself auto-advance (AC-17)
    void nextTick(() => {
      recomputeScale();
      input.enterFullscreen();
      if (!presenter.following.value) showTimer.begin(); // the show starts (once per mount)
    });
  }
}

// --- URL tracking (replace, never push: one history entry per session) ------
let exiting = false;
function syncUrl() {
  if (exiting || !presentation.value) return;
  const target = `/slides/${presentationId.value}/present/${playback.current.value.slideIndex + 1}`;
  const follow = presenter.following.value ? presenter.session : undefined;
  if (route.path !== target || route.query.follow !== follow) {
    void router.replace(follow ? { path: target, query: route.query } : target); // promoted: ?follow= drops
  }
}
watch(() => playback.current.value.slideIndex, syncUrl);

void loadPresentation();
void initLibrary().then(loadPresentation);

// --- fit-to-viewport scaling ------------------------------------------------
const stageContainer = useTemplateRef("stageContainer");
const stageScale = ref(1);
function recomputeScale() {
  const model = presentation.value;
  if (!model || typeof window === "undefined") return;
  const availW = window.innerWidth;
  const availH = window.innerHeight;
  const surfaceW = cmToPx(model.slideWidthCm);
  const surfaceH = cmToPx(model.slideHeightCm);
  if (surfaceW <= 0 || surfaceH <= 0 || availW <= 0 || availH <= 0) return;
  stageScale.value = Math.min(availW / surfaceW, availH / surfaceH);
}

// --- fullscreen (requested from the Present gesture; denial tolerated — the
// request and release live in usePlayInput beside SPEC-100's prompt) ---------
function onFullscreenChange() {
  // The user leaving OS fullscreen (Esc) must exit playback, not strand a black
  // page (AC-10). Our own exit()'s exitFullscreen also lands here — the exiting
  // guard stops the recursion. A follower stays put: its prompt returns.
  if (typeof document !== "undefined" && !document.fullscreenElement && input.fullscreenRequested() && !exiting && !presenter.following.value) {
    exit();
  }
}

// --- exit (Esc / end screen / fullscreen-leave / browser back) --------------
function exit() {
  if (exiting) return;
  exiting = true;
  autoAdvance.clear();
  clearTransition();
  buildPreview.stop();
  input.releaseFullscreen();
  const index = playback.atEnd.value ? Math.max(0, slideCount.value - 1) : playback.current.value.slideIndex;
  void router.replace({ path: `/slides/${presentationId.value}`, query: { return: String(index + 1) } });
}

// --- transitions (SPEC-085) -------------------------------------------------
// A forward step between slides plays the INCOMING slide's transition as two
// stacked stages; backward steps and jumps are instant (AC-13/14). An advance
// input arriving mid-transition completes the in-flight transition instantly
// (snap to the plain incoming stage) AND then applies the advance itself — the
// animation is collapsed, not left running and not queued behind a second one
// (AC-15). `advanceOnClick` is never consulted here: in-editor playback always
// advances (the presenter escape hatch, AC-18).
const transitionView = shallowRef(emptyTwoStageView());
// The outgoing stage shows its slide fully built (a forward step only leaves a
// slide once every group has played), so exited frames stay gone; the incoming
// stage shares the live build state, which the step watcher has already settled
// to 0 groups — unbuilt frames stay hidden through the transition (SPEC-086).
const transitionOutgoing = computed(() => fullyBuiltStyles(transitionView.value?.outgoing, presentation.value));
function clearTransition() {
  transitionView.value = null;
}
function onTransitionDone() {
  clearTransition();
}

// --- advance / navigation ---------------------------------------------------
function advanceStep() {
  if (transitionView.value) {
    // Complete the in-flight transition instantly, then fall through to apply
    // this input as the next advance — the input is honoured (never dropped),
    // but only one animation is ever live, so bursts collapse rather than
    // queue a second advance (AC-15).
    clearTransition();
  }
  // A build group animating on the current slide (SPEC-086): the first advance
  // snaps it to its completed state and holds — unlike a transition it does NOT
  // then walk on to the next group or slide (AC-24). The next advance proceeds.
  if (buildPreview.complete()) {
    return;
  }
  if (playback.atEnd.value) {
    exit();
    return;
  }
  const fromIndex = playback.current.value.slideIndex;
  const fromSlide = currentSlide.value;
  playback.advance();
  if (playback.atEnd.value) {
    return; // advanced onto the end screen — no transition
  }
  if (playback.current.value.slideIndex === fromIndex) {
    return; // a build advanced on the same slide (SPEC-086) — the build watcher animates it
  }
  playIncomingTransition(fromSlide);
}
/** Plays the incoming slide's transition from `fromSlide` — after an advance
 * here, or (SPEC-100) a follower's one-step advance arriving from the owner. */
function playIncomingTransition(fromSlide = currentSlide.value) {
  const toSlide = currentSlide.value;
  if (toSlide && toSlide.transition && transitionAnimates(toSlide.transition, reducedMotion.value, !!fromSlide)) {
    transitionView.value = { outgoing: fromSlide, incoming: toSlide, transition: toSlide.transition };
  }
}
function backStep() {
  clearTransition();
  playback.back();
}
function jumpTo(index = 0) {
  clearTransition();
  playback.goTo(index);
}

// --- input (SPEC-084's keyboard, pointer and idle HUD; SPEC-100 routes it) --
// SPEC-105: a pointer-down on a trigger frame plays its interactive sequence.
const input = usePlayInput({ atEnd: playback.atEnd, isExiting: () => exiting, perform: presenter.perform, fullscreenTarget: () => stageContainer.value,
  pointerTrigger: (event) => !playback.atEnd.value && buildPreview.playTriggerAt(currentSlide.value, event, reducedMotion.value),
  extraKey: showTimerKeys({ atEnd: playback.atEnd, perform: presenter.perform, onKey: () => !presenter.following.value && input.revealControls() }) });

// --- lifecycle --------------------------------------------------------------
onMounted(() => {
  window.addEventListener("resize", recomputeScale);
  document.addEventListener("fullscreenchange", onFullscreenChange);
  void nextTick(recomputeScale);
});
onBeforeUnmount(() => {
  window.removeEventListener("resize", recomputeScale);
  document.removeEventListener("fullscreenchange", onFullscreenChange);
  autoAdvance.clear();
  buildPreview.stop();
  showingMedia.releaseAll();
});
</script>

<template>
  <div
    v-if="loaded && !presentation"
    class="play-not-found"
    data-testid="presentation-play-not-found"
  >
    <h2>Presentation not found</h2>
    <p>This link doesn't match a presentation in your library.</p>
    <UiButton
      data-testid="presentation-play-back"
      @click="router.push('/')"
    >
      Back to library
    </UiButton>
  </div>

  <div
    v-else
    ref="stageContainer"
    class="play-stage"
    role="region"
    aria-roledescription="slide show"
    :aria-label="deckTitle"
    data-testid="presentation-play"
    :data-building="buildPreview.playing.value ? 'true' : 'false'"
    @click="input.onStageClick"
    @pointerdown="input.onPointerDown"
    @pointerup="input.onPointerUp"
    @pointermove="input.onPointerMove"
  >
    <SlideTransitionLayer
      v-if="presentation && transitionView && !playback.atEnd.value"
      :outgoing="transitionView.outgoing ?? undefined"
      :incoming="transitionView.incoming"
      :transition="transitionView.transition"
      :width-cm="presentation.slideWidthCm"
      :height-cm="presentation.slideHeightCm"
      :media="presentation.media ?? []"
      :scale="stageScale"
      :reduced-motion="reducedMotion"
      :theme="presentation.theme"
      :outgoing-frame-styles="transitionOutgoing.frameStyles"
      :outgoing-paragraph-styles="transitionOutgoing.paragraphStyles"
      :incoming-frame-styles="buildPreview.frameStyles.value"
      :incoming-paragraph-styles="buildPreview.paragraphStyles.value"
      :outgoing-media-url-override="showingMedia.previousUrl()"
      :incoming-media-url-override="showingMedia.currentUrl()"
      @done="onTransitionDone"
    />
    <SlideStage
      v-else-if="presentation && currentSlide && !playback.atEnd.value"
      :slide="currentSlide"
      :width-cm="presentation.slideWidthCm"
      :height-cm="presentation.slideHeightCm"
      :media="presentation.media ?? []"
      :scale="stageScale"
      :interactive="false"
      :playback-mode="true"
      :frame-styles="buildPreview.frameStyles.value"
      :paragraph-styles="buildPreview.paragraphStyles.value"
      :media-url-override="showingMedia.currentUrl()"
      :fields="currentFields"
      :theme="presentation.theme"
    />

    <div
      v-if="playback.atEnd.value"
      class="end-screen"
      role="alert"
      data-testid="play-end-screen"
    >
      End of slide show — click or press any key to exit
    </div>

    <div
      v-if="presenter.blank.value !== 'none'"
      class="play-blank"
      :class="`is-${presenter.blank.value}`"
      data-testid="play-blank"
      :data-blank="presenter.blank.value"
      aria-hidden="true"
    />
    <div
      v-if="presenter.isAudience && !input.fullscreenActive.value"
      class="play-fullscreen"
    >
      <button
        type="button"
        class="hud-pill"
        data-testid="play-fullscreen-prompt"
        @click.stop="input.requestFullscreen"
      >
        Full screen (F)
      </button>
      <span
        v-if="route.query.placed !== '1'"
        class="hud-pill"
        data-testid="play-screen-hint"
      >Drag this window to your projector, then press F for full screen</span>
    </div>

    <div
      class="play-hud"
      :class="{ 'is-visible': input.controlsVisible.value && !playback.atEnd.value }"
    >
      <button
        type="button"
        class="hud-affordance"
        data-testid="play-prev"
        aria-label="Previous slide"
        @click.stop="presenter.perform('back')"
      >
        ‹
      </button>
      <div
        class="hud-counter"
        data-testid="play-slide-counter"
      >
        {{ counterLabel }}
      </div>
      <button
        type="button"
        class="hud-affordance"
        data-testid="play-next"
        aria-label="Next slide"
        @click.stop="presenter.perform('advance')"
      >
        ›
      </button>
      <ShowTimer
        v-if="!presenter.following.value"
        testid-prefix="play"
        v-bind="showTimer.display.value"
        @toggle="presenter.perform('timerToggle')"
        @stop="presenter.perform('timerStop')"
      />
    </div>

    <div
      class="sr-only"
      aria-live="polite"
      data-testid="play-live"
    >
      {{ announce }}
    </div>
  </div>
</template>

<style scoped>
.play-not-found {
  height: 100vh;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-4);
  background: var(--surface-default);
  color: var(--text-primary);
  text-align: center;
  padding: var(--space-6);
}

.play-not-found p {
  margin: 0;
  color: var(--text-secondary);
}

/* Black is explicit and theme-independent — the show is never light/dark. */
.play-stage {
  position: fixed;
  inset: 0;
  width: 100vw;
  height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #000;
  overflow: hidden;
  cursor: default;
  touch-action: none;
  /* SPEC-146: the HUD pill's ink, for ShowTimer (whose styles take tokens only). */
  --hud-ink: #fff;
  --hud-pill: color-mix(in srgb, #000 55%, transparent);
  --hud-shade: #000;
}

/* SPEC-097: rendered links (a linked run, a linked frame's overlay) are live
   in the show. */
.play-stage :deep([data-testid="slide-link"]) {
  cursor: pointer;
}

.end-screen {
  color: #fff;
  background: #000;
  font-size: var(--text-xl);
  font-weight: var(--weight-medium);
  text-align: center;
  padding: var(--space-6);
}

.play-hud {
  position: absolute;
  inset-block-end: var(--space-5);
  inset-inline: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-5);
  opacity: 0;
  transition: opacity var(--motion-fast) var(--easing);
  pointer-events: none;
}

.play-hud.is-visible {
  opacity: 1;
  pointer-events: auto;
}

/* SPEC-100 blank screen: solid over the stage, the HUD still above it. */
.play-blank {
  position: absolute;
  inset: 0;
  background: #000;
}

.play-blank.is-white {
  background: #fff;
}

.play-fullscreen {
  position: absolute;
  inset-block-start: var(--space-4);
  inset-inline: 0;
  display: flex;
  flex-wrap: wrap;
  justify-content: center;
  gap: var(--space-3);
}

.hud-pill,
.hud-counter {
  color: #fff;
  background: color-mix(in srgb, #000 55%, transparent);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-md);
  font-variant-numeric: tabular-nums;
}

.hud-pill {
  border: none;
  font: inherit;
}

button.hud-pill {
  cursor: pointer;
}

.hud-affordance {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: var(--space-7);
  height: var(--space-7);
  border: none;
  border-radius: var(--radius-full, 999px);
  background: color-mix(in srgb, #000 55%, transparent);
  color: #fff;
  font-size: var(--text-xl);
  line-height: 1;
  cursor: pointer;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
