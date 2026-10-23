// The presenter window's controller (SPEC-100) — everything PresenterView
// wires. The window at /slides/:id/presenter/:n loads the deck exactly as the
// play view does (initLibrary().then(load), the legacy upgrade, the deep link
// clamped into range), OWNS the show's single usePlayback machine, mints the
// session and opens the audience window (useAudienceWindow), and drives it
// through usePresenterSession. Its own panes never transition — only the
// audience plays SPEC-085's transitions — but the current-slide pane runs its
// own useBuildPreview through the same watchBuildSteps the show uses, so it
// mirrors the audience's build state exactly; and an advance that lands while
// that pane is mid-build completes it and holds, as in the show (AC-24).
//
// The surface: counters, the next slide, the speaker notes and the timer (both
// in usePresenterTools), the slide grid and the keys. Keys go
// through usePlayInput's playKeyHandler — the show's own classifier — with the
// presenter's extras (g grid, +/- notes, t/r timer; Escape closes the grid
// first) layered in front. Presenting writes nothing: no dirty flag, no
// markSlideEdited, no edit lock; only a legacy deck's upgrade stores, as
// opening one anywhere does.
import { computed, nextTick, onBeforeUnmount, ref, shallowRef, watch } from "vue";
import type { RouteLocationNormalizedLoaded, Router } from "vue-router";
import { filmstripLabel, stageFields } from "../../components/presentation/slideRender";
import { now } from "../../lib/core/clock";
import type { WindowChannel, WindowChannelTransport } from "../../lib/core/windowChannel";
import { documentKindOf, getPresentation, initLibrary, storePresentation } from "../../lib/documentStore";
import { slideBuildCount } from "../../lib/presentation/buildPlayback";
import { isLegacyPresentationModel } from "../../lib/presentation/model";
import { upgradePresentationModel } from "../../lib/presentation/open";
import { playbackActionForKey } from "../../lib/presentation/playbackKeys";
import { mintPresenterSession } from "../../lib/presentation/presenterSession";
import { useAudienceWindow, type AudienceWindowDeps } from "./useAudienceWindow";
import { useBuildPreview, watchBuildSteps } from "./useBuildPreview";
import { playKeyHandler } from "./usePlayInput";
import { usePlayback } from "./usePlayback";
import { usePresenterSession } from "./usePresenterSession";
import { usePresenterTools } from "./usePresenterTools";
import { useReducedMotion } from "./useReducedMotion";
import { useWindowKeys } from "./useWindowKeys";

export interface PresenterViewDeps {
  route: RouteLocationNormalizedLoaded;
  router: Router;
  /** Test seams. */
  openWindow?: AudienceWindowDeps["openWindow"];
  screenAccess?: AudienceWindowDeps["screenAccess"];
  openChannel?: (name: string, onMessage: (payload: unknown) => void) => WindowChannel;
  transport?: () => WindowChannelTransport;
  nowMs?: () => number;
}

function isControl(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && (target.tagName === "BUTTON" || target.getAttribute("role") === "option");
}

export function usePresenterView(deps: PresenterViewDeps) {
  const { route, router } = deps;
  const nowMs = deps.nowMs ?? (() => now().getTime());
  const presentationId = computed(() => String(route.params.id ?? ""));

  // --- the deck (loaded exactly as PresentationPlayView loads it) -------------
  function initialModel() {
    const model = getPresentation(presentationId.value) ?? null;
    return model && isLegacyPresentationModel(model) ? null : model;
  }
  const presentation = shallowRef(initialModel());
  const loaded = ref(false);
  const slideCount = computed(() => presentation.value?.slides.length ?? 0);
  const reducedMotion = useReducedMotion();

  // --- the step machine (the one the show runs on) ----------------------------
  function routeStartIndex(): number {
    const n = Number(route.params.n ?? 1);
    return Number.isFinite(n) ? n - 1 : 0;
  }
  function buildCountAt(index: number): number {
    const slide = presentation.value?.slides[index];
    return slide ? slideBuildCount(slide) : 0;
  }
  function isHidden(index: number): boolean {
    return presentation.value?.slides[index]?.hidden === true;
  }
  const playback = usePlayback({ slideCount, reducedMotion, start: routeStartIndex(), buildCountAt, isHidden });
  const currentIndex = computed(() => playback.current.value.slideIndex);
  const currentSlide = computed(() => presentation.value?.slides[currentIndex.value] ?? null);
  /** The slide an advance past this one shows — the next visible one. */
  const nextIndex = computed(() => {
    for (let index = currentIndex.value + 1; index < slideCount.value; index += 1) {
      if (!isHidden(index)) return index;
    }
    return -1;
  });
  /** What the two panes show (undefined: the end of the show). */
  const shownSlide = computed(() => (playback.atEnd.value ? undefined : (currentSlide.value ?? undefined)));
  const nextSlide = computed(() => (playback.atEnd.value ? undefined : presentation.value?.slides[nextIndex.value]));
  const currentFields = computed(() => {
    const model = presentation.value;
    return model && !playback.atEnd.value ? stageFields(model, currentIndex.value, now()) : [];
  });
  const nextFields = computed(() => {
    const model = presentation.value;
    return model && nextSlide.value ? stageFields(model, nextIndex.value, now()) : [];
  });

  // The current pane's builds, driven exactly as the show drives its own.
  const buildPreview = useBuildPreview();
  watchBuildSteps(buildPreview, () => playback.current.value, () => currentSlide.value, () => reducedMotion.value);

  // --- step functions (the owner's; the audience's intents land here too) -----
  let exiting = false;
  function advanceStep(): void {
    if (buildPreview.complete()) return; // a build in flight completes and holds
    if (playback.atEnd.value) {
      exit();
      return;
    }
    playback.advance();
  }
  function jumpTo(index: number): void {
    playback.goTo(index);
  }

  // --- the session and the audience window ------------------------------------
  const session = usePresenterSession({
    window: "presenter",
    deckId: presentationId,
    session: mintPresenterSession(),
    playback,
    slideCount,
    buildCountAt,
    isHidden,
    navigation: { advance: advanceStep, back: () => playback.back(), goTo: jumpTo, exit },
    onAudience: (event) => audience.onAudience(event),
    openChannel: deps.openChannel,
    transport: deps.transport,
  });
  const audience = useAudienceWindow({
    deckId: presentationId,
    session: session.session,
    slideNumber: () => currentIndex.value + 1,
    transport: session.transport,
    openWindow: deps.openWindow,
    screenAccess: deps.screenAccess,
  });
  /** Single-window mode (AC-15): no audience window could open or follow, so
   * the current-slide pane is the show. */
  const solo = computed(() => audience.status.value === "blocked");
  const blockedBannerOpen = ref(true);
  /** "Present here": this window is the show — take the current-slide pane
   * full screen from the click (denial tolerated: the pane is the show either
   * way). */
  function presentHere(pane: unknown): void {
    blockedBannerOpen.value = false;
    if (pane instanceof HTMLElement && typeof pane.requestFullscreen === "function") {
      try {
        void Promise.resolve(pane.requestFullscreen()).catch(() => {});
      } catch {
        // Unavailable — the pane stays the show in place.
      }
    }
  }
  function reopenAudience(): void {
    blockedBannerOpen.value = true;
    void audience.open();
  }

  // --- labels ------------------------------------------------------------------
  const counterLabel = computed(() => `Slide ${currentIndex.value + 1} of ${slideCount.value}`);
  const buildTotal = computed(() => (currentSlide.value ? slideBuildCount(currentSlide.value) : 0));
  const buildLabel = computed(() => `Build ${playback.current.value.buildIndex} of ${buildTotal.value}`);
  const notes = computed(() => currentSlide.value?.notes ?? []);
  const grid = computed(() =>
    (presentation.value?.slides ?? []).map((slide, index, all) => ({ slide, label: filmstripLabel(slide, index, all.length) })),
  );

  // --- the live region (slide changes and session lifecycle) -------------------
  const liveMessage = ref("");
  watch(currentIndex, (index) => {
    const slide = presentation.value?.slides[index];
    if (slide) liveMessage.value = filmstripLabel(slide, index, slideCount.value);
  });
  watch(session.blank, (blank) => {
    liveMessage.value = blank === "none" ? "Screen restored" : "Screen blanked";
  });
  watch(audience.status, (status) => {
    const messages: Record<string, string> = {
      connected: "Audience window connected",
      lost: "Audience window closed",
      blocked: "Single-window mode — the current slide is the show",
    };
    if (messages[status]) liveMessage.value = messages[status];
  });

  // --- auto-advance (SPEC-085 AC-17): the owner schedules it -------------------
  let autoAdvanceTimer = 0;
  function clearAutoAdvance(): void {
    if (autoAdvanceTimer) clearTimeout(autoAdvanceTimer);
    autoAdvanceTimer = 0;
  }
  function scheduleAutoAdvance(): void {
    clearAutoAdvance();
    const ms = currentSlide.value?.transition?.advanceAfterMs;
    if (typeof ms === "number" && ms > 0 && !playback.atEnd.value) {
      autoAdvanceTimer = window.setTimeout(() => {
        autoAdvanceTimer = 0;
        advanceStep();
      }, ms);
    }
  }
  watch([currentIndex, () => playback.atEnd.value], scheduleAutoAdvance);

  // --- the timer and the notes scale --------------------------------------------
  const tools = usePresenterTools(nowMs);

  // --- the grid ----------------------------------------------------------------
  const gridOpen = ref(false);
  function toggleGrid(): void {
    gridOpen.value = !gridOpen.value;
  }
  function chooseSlide(index: number): void {
    gridOpen.value = false;
    session.perform("goTo", index);
  }

  // --- keys --------------------------------------------------------------------
  /** The presenter's own keys, in front of the show's classifier. */
  function extraKey(event: KeyboardEvent): boolean {
    const key = event.key.length === 1 ? event.key.toLowerCase() : event.key;
    if (gridOpen.value) {
      // The open grid owns its keys; Escape (or g) closes it before anything exits.
      if (key === "Escape" || key === "g") {
        event.preventDefault();
        gridOpen.value = false;
      }
      return true;
    }
    // A focused button or option takes its own Enter/Space activation.
    if ((key === "Enter" || key === " ") && isControl(event.target)) return true;
    if (playbackActionForKey(event.key)) return false; // the show's map first
    const extras: Record<string, () => void> = {
      g: toggleGrid,
      "+": () => tools.stepNotesScale(1),
      "=": () => tools.stepNotesScale(1),
      "-": () => tools.stepNotesScale(-1),
      t: tools.toggleTimer,
      r: tools.resetTimer,
    };
    const run = extras[key];
    if (!run) return false;
    event.preventDefault();
    run();
    return true;
  }
  useWindowKeys([
    playKeyHandler({ atEnd: playback.atEnd, isExiting: () => exiting, perform: session.perform, extraKey }),
  ]);

  // --- load --------------------------------------------------------------------
  let audienceRequested = false;
  async function loadPresentation(): Promise<void> {
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
    if (!model) return;
    // Clamp the deep-link n and land exactly on it (a hidden one still shows).
    playback.setStep({ slideIndex: routeStartIndex(), buildIndex: 0 });
    syncUrl();
    if (currentSlide.value) buildPreview.settle(currentSlide.value, playback.current.value.buildIndex);
    scheduleAutoAdvance();
    if (!audienceRequested) {
      audienceRequested = true;
      tools.startTimer(); // the timer starts with the show
      // The gesture that opened this route is still live; open the audience once.
      void nextTick(() => void audience.open());
    }
  }

  // --- URL tracking (replace, never push) --------------------------------------
  function syncUrl(): void {
    if (exiting || !presentation.value) return;
    const target = `/slides/${presentationId.value}/presenter/${currentIndex.value + 1}`;
    if (route.path !== target) void router.replace(target);
  }
  watch(currentIndex, syncUrl);

  // --- exit (Escape, presenter-exit, advance past the end) ---------------------
  function exit(): void {
    if (exiting) return;
    exiting = true;
    clearAutoAdvance();
    buildPreview.stop();
    session.leave(); // bye — the audience it opened is closing anyway
    audience.close();
    if (typeof document !== "undefined" && document.fullscreenElement) {
      void Promise.resolve(document.exitFullscreen?.()).catch(() => {});
    }
    const index = playback.atEnd.value ? Math.max(0, slideCount.value - 1) : currentIndex.value;
    void router.replace({ path: `/slides/${presentationId.value}`, query: { return: String(index + 1) } });
  }

  void loadPresentation();
  void initLibrary().then(loadPresentation);

  onBeforeUnmount(() => {
    clearAutoAdvance();
    buildPreview.stop();
  });

  return {
    presentation,
    loaded,
    playback,
    blank: session.blank,
    perform: session.perform,
    transport: session.transport,
    currentSlide,
    shownSlide,
    currentFields,
    nextSlide,
    nextFields,
    buildPreview,
    counterLabel,
    buildTotal,
    buildLabel,
    notes,
    grid,
    liveMessage,
    audienceStatus: audience.status,
    solo,
    blockedBannerOpen,
    presentHere,
    reopenAudience,
    ...tools,
    gridOpen,
    toggleGrid,
    chooseSlide,
    exit,
    goHome: () => router.push("/"),
  };
}
