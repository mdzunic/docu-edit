// Auto-advance (SPEC-085 AC-17), shared by the two windows that can own a show
// — PresentationPlayView and the presenter window — since SPEC-146, with the
// behaviour both had: a slide whose transition carries `advanceAfterMs`
// schedules one forward step; scheduling again (any navigation, which changes
// the index) replaces it, and the end screen or `enabled() === false` (a
// following audience window: the owner advances) schedules nothing. The views
// call `schedule` on load and promotion and from their step watch, and
// `clear` on exit and unmount.
import type { Ref } from "vue";
import type { Slide } from "../../lib/presentation/model";

export interface AutoAdvanceDeps {
  currentSlide: Readonly<Ref<Slide | null | undefined>>;
  atEnd: Readonly<Ref<boolean>>;
  enabled: () => boolean;
  advance: () => void;
}

export function useAutoAdvance(deps: AutoAdvanceDeps) {
  let timer = 0;
  function clear(): void {
    if (timer) clearTimeout(timer);
    timer = 0;
  }
  function schedule(): void {
    clear();
    const ms = deps.currentSlide.value?.transition?.advanceAfterMs;
    if (typeof ms === "number" && ms > 0 && !deps.atEnd.value && deps.enabled()) {
      timer = window.setTimeout(() => {
        timer = 0;
        deps.advance();
      }, ms);
    }
  }
  return { schedule, clear };
}
