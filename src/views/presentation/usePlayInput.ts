// The slideshow's input handling (SPEC-084), moved out of PresentationPlayView
// for SPEC-100 with its behaviour unchanged: the window-level keyboard (the
// shared playbackKeys map, Escape, the end screen's any-key exit, number-then-
// Enter jumps), the stage's pointer (click advances, a horizontal swipe
// navigates, a click on a rendered link follows it — SPEC-097) and the idle-
// fade HUD. What changed is only where an input goes: every one is handed to
// `perform` — the presenter session's entry point — which runs it in a window
// that owns the show and forwards it as an intent from an audience window
// that follows one (SPEC-100), so the input map exists exactly once.
//
// `playKeyHandler` is the keyboard half on its own, so the presenter window
// classifies keys through the very same function (its extra keys — grid,
// notes scale, timer — come in through `extraKey`, which runs first).
// The view's load-time full-screen request and its release on exit moved here
// too, unchanged, beside what SPEC-100 adds: `f`, and the audience's persistent
// full-screen prompt — a browser may refuse the load-time request in a window
// that never had a gesture, so a real key press or click grants it.
import { onBeforeUnmount, onMounted, ref, type Ref } from "vue";
import { slideLinkElement, slideLinkTargetOf } from "../../components/presentation/slideRender";
import { playbackActionForKey } from "../../lib/presentation/playbackKeys";
import type { PresenterIntentAction } from "../../lib/presentation/presenterSession";

export type PlayPerform = (action: PresenterIntentAction, slideIndex?: number) => void;

export interface PlayKeyDeps {
  atEnd: Ref<boolean>;
  isExiting: () => boolean;
  perform: PlayPerform;
  /** Runs before the shared map; returns true when it consumed the key. */
  extraKey?: (event: KeyboardEvent) => boolean;
}

export interface PlayInputDeps extends PlayKeyDeps {
  /** The element `f` and the prompt take full screen (the stage). */
  fullscreenTarget: () => HTMLElement | null;
  /** SPEC-105: offered every stage pointer-down first; returns true when it
   * consumed it (a click on a trigger frame playing its interactive
   * sequence), and the press then neither advances nor swipes. */
  pointerTrigger?: (event: PointerEvent) => boolean;
}

function isTextInput(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
  );
}

/** The slideshow keyboard: one window-level keydown handler. */
export function playKeyHandler(deps: PlayKeyDeps): (event: KeyboardEvent) => void {
  let digitBuffer = "";
  return (event) => {
    if (deps.isExiting()) return;
    // Ignore combos we don't own (Cmd/Ctrl+Z etc. do nothing — inert; AC-45) and
    // anything typed into an input.
    if (event.metaKey || event.ctrlKey || event.altKey) return;
    if (event.target && isTextInput(event.target)) return;
    // Enter on a focused slide link activates it (its click follows the link,
    // SPEC-097); it must not also advance.
    if (event.key === "Enter" && digitBuffer.length === 0 && slideLinkElement(event.target)) return;
    if (deps.extraKey?.(event)) return;

    if (event.key === "Escape") {
      event.preventDefault();
      deps.perform("exit");
      return;
    }
    if (deps.atEnd.value) {
      // The end screen: any key exits (AC-31/32).
      event.preventDefault();
      deps.perform("exit");
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
      deps.perform("goTo", target - 1);
      return;
    }
    const action = playbackActionForKey(event.key);
    if (!action) return;
    event.preventDefault();
    digitBuffer = "";
    deps.perform(action);
  };
}

const SWIPE_THRESHOLD_PX = 48;

function releaseFullscreen(): void {
  if (typeof document !== "undefined" && document.fullscreenElement && typeof document.exitFullscreen === "function") {
    try {
      void Promise.resolve(document.exitFullscreen()).catch(() => {});
    } catch {
      // ignore
    }
  }
}

export function usePlayInput(deps: PlayInputDeps) {
  // --- fullscreen (requested from the Present gesture; denial tolerated) -------
  let fullscreenRequested = false;
  function enterFullscreen(): void {
    const el = deps.fullscreenTarget();
    if (!el || fullscreenRequested || typeof el.requestFullscreen !== "function") return;
    fullscreenRequested = true;
    try {
      void Promise.resolve(el.requestFullscreen()).catch(() => {});
    } catch {
      // Unavailable — playback just fills the viewport (AC-9).
    }
  }
  // SPEC-100: the same request from a real gesture (`f`, the prompt).
  const fullscreenActive = ref(typeof document !== "undefined" && !!document.fullscreenElement);
  function requestFullscreen(): void {
    const el = deps.fullscreenTarget();
    if (!el || typeof el.requestFullscreen !== "function" || document.fullscreenElement) return;
    try {
      void Promise.resolve(el.requestFullscreen()).catch(() => {});
    } catch {
      // Refused — the prompt stays up.
    }
  }
  function onFullscreenChange(): void {
    fullscreenActive.value = !!document.fullscreenElement;
  }

  // --- keyboard (window-level while mounted) ----------------------------------
  const onKeydown = playKeyHandler({
    ...deps,
    extraKey: (event) => {
      if (deps.extraKey?.(event)) return true;
      if (event.key === "f" || event.key === "F") {
        event.preventDefault();
        requestFullscreen();
        return true;
      }
      return false;
    },
  });

  // --- pointer / touch: click advances, horizontal swipe navigates ------------
  let pointerActive = false;
  let pointerStartX = 0;
  let pointerStartY = 0;
  let swiped = false;
  let triggered = false;
  function onPointerDown(event: PointerEvent): void {
    triggered = deps.pointerTrigger?.(event) === true;
    if (triggered) {
      pointerActive = false;
      return;
    }
    pointerActive = true;
    swiped = false;
    pointerStartX = event.clientX;
    pointerStartY = event.clientY;
  }
  function onPointerUp(event: PointerEvent): void {
    if (!pointerActive) return;
    pointerActive = false;
    const dx = event.clientX - pointerStartX;
    const dy = event.clientY - pointerStartY;
    if (Math.abs(dx) >= SWIPE_THRESHOLD_PX && Math.abs(dx) > Math.abs(dy)) {
      swiped = true; // suppress the click that follows
      if (deps.atEnd.value) {
        deps.perform("exit");
      } else if (dx < 0) {
        deps.perform("advance"); // swipe left → next (SPEC-068 convention)
      } else {
        deps.perform("back");
      }
    }
  }
  function onStageClick(event: MouseEvent): void {
    if (swiped || triggered) {
      swiped = false;
      triggered = false;
      return;
    }
    // SPEC-097 AC-49/50: a click on a rendered link follows it and NEVER also
    // advances (or completes) a build — return before the advance.
    const link = slideLinkElement(event.target);
    if (link) {
      followLink(slideLinkTargetOf(link));
      return;
    }
    deps.perform("advance");
  }

  /** Follows a slide link: an in-deck target jumps (SPEC-084's jump, 0-based)
   * and an external one opens in a new tab with no opener/referrer. */
  function followLink(target: ReturnType<typeof slideLinkTargetOf>): void {
    if (!target) return;
    if (target.slide !== undefined) {
      deps.perform("goTo", target.slide - 1);
    } else if (target.href) {
      window.open(target.href, "_blank", "noopener,noreferrer");
    }
  }

  // --- idle-fade HUD (counter + touch prev/next), 2s after pointer idle -------
  const controlsVisible = ref(false);
  let idleTimer = 0;
  function revealControls(): void {
    controlsVisible.value = true;
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = window.setTimeout(() => {
      controlsVisible.value = false;
    }, 2000);
  }
  function onPointerMove(): void {
    revealControls();
  }

  // --- lifecycle --------------------------------------------------------------
  onMounted(() => {
    window.addEventListener("keydown", onKeydown);
    document.addEventListener("fullscreenchange", onFullscreenChange);
  });
  onBeforeUnmount(() => {
    window.removeEventListener("keydown", onKeydown);
    document.removeEventListener("fullscreenchange", onFullscreenChange);
    if (idleTimer) clearTimeout(idleTimer);
  });

  return {
    onKeydown,
    onPointerDown,
    onPointerUp,
    onPointerMove,
    onStageClick,
    controlsVisible,
    enterFullscreen,
    releaseFullscreen,
    fullscreenRequested: () => fullscreenRequested,
    fullscreenActive,
    requestFullscreen,
  };
}
