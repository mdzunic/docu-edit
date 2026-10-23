// `prefers-reduced-motion`, read from THIS window's display and kept live
// while the calling component is mounted. Lifted out of PresentationPlayView
// (SPEC-084) for SPEC-100, whose presenter and audience windows may sit on
// different displays: each reads its own and honours it locally — no motion
// flag ever travels between the two windows.
import { onBeforeUnmount, onMounted, ref, type Ref } from "vue";

export function useReducedMotion(): Ref<boolean> {
  const reducedMotion = ref(false);
  const motionQuery =
    typeof window !== "undefined" && typeof window.matchMedia === "function"
      ? window.matchMedia("(prefers-reduced-motion: reduce)")
      : null;
  if (motionQuery) {
    reducedMotion.value = motionQuery.matches;
  }
  function onMotionChange(event: MediaQueryListEvent): void {
    reducedMotion.value = event.matches;
  }
  onMounted(() => motionQuery?.addEventListener("change", onMotionChange));
  onBeforeUnmount(() => motionQuery?.removeEventListener("change", onMotionChange));
  return reducedMotion;
}
