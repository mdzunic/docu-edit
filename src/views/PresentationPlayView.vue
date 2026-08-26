<script setup lang="ts">
// The slideshow playback surface (SPEC-084): a full-screen, chrome-free page
// at /slides/:id/present/:n that walks the deck slide by slide. No AppShell,
// toolbar or filmstrip — just the SPEC-079 stage on black, letterboxed and
// scaled to fill the viewport, driven by the usePlayback step machine and the
// shared playbackKeys map. This spec is the normative source for playback
// navigation; SPEC-085/086 layer transitions/builds onto the same pipeline and
// SPEC-088 reimplements it for the HTML export. Runtime declarations, not
// defineProps<{...}>() generics — espree lints .vue scripts, so typed logic
// (the step machine, the key map) lives in the imported .ts modules.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import SlideStage from "../components/presentation/SlideStage.vue";
import SlideTransitionLayer from "../components/presentation/SlideTransitionLayer.vue";
import { emptyTwoStageView, transitionAnimates } from "../components/presentation/transitionStyles";
import UiButton from "../components/ui/UiButton.vue";
import { documentKindOf, getPresentation, initLibrary, storePresentation } from "../lib/documentStore";
import { isLegacyPresentationModel } from "../lib/presentation/model";
import { upgradePresentationModel } from "../lib/presentation/open";
import { cmToPx } from "../lib/core/units";
import { slideOutline } from "../lib/presentation/slideModel";
import { slideBuildCount } from "../lib/presentation/buildPlayback";
import { playbackActionForKey } from "../lib/presentation/playbackKeys";
import { usePlayback } from "./presentation/usePlayback";
import { useBuildPreview } from "./presentation/useBuildPreview";

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

// --- reduced motion (read from the media query, exposed reactively) ---------
const reducedMotion = ref(false);
const motionQuery =
  typeof window !== "undefined" && typeof window.matchMedia === "function"
    ? window.matchMedia("(prefers-reduced-motion: reduce)")
    : null;
if (motionQuery) {
  reducedMotion.value = motionQuery.matches;
}
function onMotionChange(event = new MediaQueryListEvent("change")) {
  reducedMotion.value = event.matches;
}

// --- the step machine -------------------------------------------------------
function routeStartIndex() {
  const n = Number(route.params.n ?? 1);
  return Number.isFinite(n) ? n - 1 : 0; // usePlayback clamps out-of-range
}
function buildCountAt(index = 0) {
  const slide = presentation.value?.slides[index];
  return slide ? slideBuildCount(slide) : 0;
}
const playback = usePlayback({ slideCount, reducedMotion, start: routeStartIndex(), buildCountAt });
const currentSlide = computed(() => presentation.value?.slides[playback.current.value.slideIndex] ?? null);

// --- build playback (SPEC-086): frames hidden until their build plays --------
const buildPreview = useBuildPreview();
watch(
  () => playback.current.value,
  (cur, prev) => {
    const slide = currentSlide.value;
    if (!slide) return;
    buildPreview.stop();
    if (!prev || cur.slideIndex !== prev.slideIndex || cur.buildIndex <= prev.buildIndex) {
      // A new slide, or a backward/jump landing: settle to the target instantly
      // (back lands fully-built — AC-22).
      buildPreview.settle(slide, cur.buildIndex);
      return;
    }
    // A forward build step on the same slide: animate the group that just began.
    buildPreview.playGroup(slide, cur.buildIndex - 1, reducedMotion.value);
  },
);
const deckTitle = computed(() => presentation.value?.title ?? "Slide show");
const counterLabel = computed(() => `${playback.current.value.slideIndex + 1} / ${slideCount.value}`);
const announce = computed(() => {
  if (!presentation.value || playback.atEnd.value) return "";
  const outline = currentSlide.value ? slideOutline(currentSlide.value) : { title: "" };
  const title = outline.title && outline.title.length > 0 ? outline.title : "untitled";
  return `Slide ${playback.current.value.slideIndex + 1} of ${slideCount.value} — ${title}`;
});

// --- auto-advance (AC-17): a slide with advanceAfterMs schedules a forward
// step; any navigation (which changes the index) or exit reschedules/clears it.
// Declared before the load invocation below so the setup-time load can call it.
let autoAdvanceTimer = 0;
function clearAutoAdvance() {
  if (autoAdvanceTimer) {
    clearTimeout(autoAdvanceTimer);
    autoAdvanceTimer = 0;
  }
}
function scheduleAutoAdvance() {
  clearAutoAdvance();
  const ms = currentSlide.value?.transition?.advanceAfterMs;
  if (typeof ms === "number" && ms > 0 && !playback.atEnd.value) {
    autoAdvanceTimer = window.setTimeout(() => {
      autoAdvanceTimer = 0;
      advanceStep();
    }, ms);
  }
}
watch([() => playback.current.value.slideIndex, () => playback.atEnd.value], scheduleAutoAdvance);

// --- load (mirrors PresentationView: text id redirects, legacy upgrades) ----
async function loadPresentation() {
  if (documentKindOf(presentationId.value) === "text") {
    void router.replace(`/doc/${presentationId.value}`);
    return;
  }
  let model = getPresentation(presentationId.value) ?? null;
  if (model && isLegacyPresentationModel(model)) {
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
    playback.goTo(routeStartIndex()); // clamp the deep-link n into range
    syncUrl(); // reflect the clamp even when the index did not change (AC-5)
    if (currentSlide.value) {
      buildPreview.settle(currentSlide.value, playback.current.value.buildIndex); // hide unbuilt frames (AC-20)
    }
    scheduleAutoAdvance(); // the starting slide may itself auto-advance (AC-17)
    void nextTick(() => {
      recomputeScale();
      enterFullscreen();
    });
  }
}

// --- URL tracking (replace, never push: one history entry per session) ------
let exiting = false;
function syncUrl() {
  if (exiting || !presentation.value) return;
  const target = `/slides/${presentationId.value}/present/${playback.current.value.slideIndex + 1}`;
  if (route.path !== target) {
    void router.replace(target);
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

// --- fullscreen (requested from the Present gesture; denial tolerated) -------
let fullscreenRequested = false;
function enterFullscreen() {
  const el = stageContainer.value;
  if (!el || fullscreenRequested || typeof el.requestFullscreen !== "function") return;
  fullscreenRequested = true;
  try {
    void Promise.resolve(el.requestFullscreen()).catch(() => {});
  } catch {
    // Unavailable — playback just fills the viewport (AC-9).
  }
}
function releaseFullscreen() {
  if (typeof document !== "undefined" && document.fullscreenElement && typeof document.exitFullscreen === "function") {
    try {
      void Promise.resolve(document.exitFullscreen()).catch(() => {});
    } catch {
      // ignore
    }
  }
}
function onFullscreenChange() {
  // The user leaving OS fullscreen (Esc) must exit playback, not strand a black
  // page (AC-10). Our own exit()'s exitFullscreen also lands here — the exiting
  // guard stops the recursion.
  if (typeof document !== "undefined" && !document.fullscreenElement && fullscreenRequested && !exiting) {
    exit();
  }
}

// --- exit (Esc / end screen / fullscreen-leave / browser back) --------------
function exit() {
  if (exiting) return;
  exiting = true;
  clearAutoAdvance();
  clearTransition();
  buildPreview.stop();
  releaseFullscreen();
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

// --- keyboard (window-level while mounted) ----------------------------------
let digitBuffer = "";
function isTextInput(target = new EventTarget()) {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
  );
}
function onKeydown(event = new KeyboardEvent("keydown")) {
  if (exiting) return;
  // Ignore combos we don't own (Cmd/Ctrl+Z etc. do nothing — inert; AC-45) and
  // anything typed into an input.
  if (event.metaKey || event.ctrlKey || event.altKey) return;
  if (event.target && isTextInput(event.target)) return;

  if (event.key === "Escape") {
    event.preventDefault();
    exit();
    return;
  }
  if (playback.atEnd.value) {
    // The end screen: any key exits (AC-31/32).
    event.preventDefault();
    exit();
    return;
  }
  if (event.key >= "0" && event.key <= "9") {
    digitBuffer += event.key;
    event.preventDefault();
    return;
  }
  if (event.key === "Enter" && digitBuffer.length > 0) {
    event.preventDefault();
    const target = Number(digitBuffer);
    digitBuffer = "";
    jumpTo(target - 1);
    return;
  }
  const action = playbackActionForKey(event.key);
  if (!action) return;
  event.preventDefault();
  digitBuffer = "";
  if (action === "advance") advanceStep();
  else if (action === "back") backStep();
  else if (action === "first") jumpTo(0);
  else if (action === "last") jumpTo(slideCount.value - 1);
}

// --- pointer / touch: click advances, horizontal swipe navigates ------------
const SWIPE_THRESHOLD_PX = 48;
let pointerActive = false;
let pointerStartX = 0;
let pointerStartY = 0;
let swiped = false;
function onPointerDown(event = new PointerEvent("pointerdown")) {
  pointerActive = true;
  swiped = false;
  pointerStartX = event.clientX;
  pointerStartY = event.clientY;
}
function onPointerUp(event = new PointerEvent("pointerup")) {
  if (!pointerActive) return;
  pointerActive = false;
  const dx = event.clientX - pointerStartX;
  const dy = event.clientY - pointerStartY;
  if (Math.abs(dx) >= SWIPE_THRESHOLD_PX && Math.abs(dx) > Math.abs(dy)) {
    swiped = true; // suppress the click that follows
    if (playback.atEnd.value) {
      exit();
    } else if (dx < 0) {
      advanceStep(); // swipe left → next (SPEC-068 convention)
    } else {
      backStep();
    }
  }
}
function onStageClick() {
  if (swiped) {
    swiped = false;
    return;
  }
  advanceStep();
}

// --- idle-fade HUD (counter + touch prev/next), 2s after pointer idle -------
const controlsVisible = ref(false);
let idleTimer = 0;
function revealControls() {
  controlsVisible.value = true;
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = window.setTimeout(() => {
    controlsVisible.value = false;
  }, 2000);
}
function onPointerMove() {
  revealControls();
}

// --- lifecycle --------------------------------------------------------------
onMounted(() => {
  window.addEventListener("keydown", onKeydown);
  window.addEventListener("resize", recomputeScale);
  document.addEventListener("fullscreenchange", onFullscreenChange);
  motionQuery?.addEventListener("change", onMotionChange);
  void nextTick(recomputeScale);
});
onBeforeUnmount(() => {
  window.removeEventListener("keydown", onKeydown);
  window.removeEventListener("resize", recomputeScale);
  document.removeEventListener("fullscreenchange", onFullscreenChange);
  motionQuery?.removeEventListener("change", onMotionChange);
  if (idleTimer) clearTimeout(idleTimer);
  clearAutoAdvance();
  buildPreview.stop();
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
    @click="onStageClick"
    @pointerdown="onPointerDown"
    @pointerup="onPointerUp"
    @pointermove="onPointerMove"
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
      :frame-styles="buildPreview.frameStyles.value"
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
      class="play-hud"
      :class="{ 'is-visible': controlsVisible && !playback.atEnd.value }"
    >
      <button
        type="button"
        class="hud-affordance"
        data-testid="play-prev"
        aria-label="Previous slide"
        @click.stop="backStep"
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
        @click.stop="advanceStep"
      >
        ›
      </button>
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

.hud-counter {
  color: #fff;
  background: color-mix(in srgb, #000 55%, transparent);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-md);
  font-variant-numeric: tabular-nums;
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
