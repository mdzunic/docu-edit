import { computed, ref, type ComputedRef } from "vue";

// SPEC-068 — the three width-based layout tiers. Boundaries key off the app's
// available width (measured by a ResizeObserver on the shell root), never off
// device sniffing (AC-1): a narrow window on a desktop behaves identically to
// a tablet of the same width. Pointer coarseness is a SEPARATE axis handled in
// CSS (`@media (pointer: coarse)`) and never conflated with these tiers.
export type LayoutTier = "compact" | "medium" | "full";

/** Below this the layout is compact (<640px). */
export const COMPACT_MAX_WIDTH = 640;
/** At or above this the layout is full (≥1024px); the range in between is
 * the medium tier (640–1023px). */
export const FULL_MIN_WIDTH = 1024;

/** Maps an available width in CSS px to its layout tier. The single source of
 * the tier boundaries, so the composable and its consumers never re-spell the
 * numbers. */
export function tierForWidth(width: number): LayoutTier {
  if (width < COMPACT_MAX_WIDTH) {
    return "compact";
  }
  if (width < FULL_MIN_WIDTH) {
    return "medium";
  }
  return "full";
}

// Module-level shared state, like useToasts: the tier is a single app-wide
// fact read by chrome deep inside DocumentView's slots (menu bar, toolbar,
// status bar, side panel) as well as by canvas descendants (link chip, context
// menus). provide/inject would not reach slotted content — a slot's injection
// ancestor is the component that DECLARES it (DocumentView), not the one that
// renders the slot (AppShell) — so a shared ref sidesteps the tree entirely and
// AppShell's observer feeds it. Seeded from window width so the first paint is
// already close before the observer measures the real shell root.
const initialWidth =
  typeof window !== "undefined" && typeof window.innerWidth === "number"
    ? window.innerWidth
    : FULL_MIN_WIDTH;
const shellWidth = ref(initialWidth);
const tier = computed<LayoutTier>(() => tierForWidth(shellWidth.value));

/** Installs a ResizeObserver on the shell root so the tier tracks the app's
 * available width live (AC-1, AC-2 — moves between tiers as the window resizes,
 * no reload). Called once by AppShell; returns a stop function AppShell wires
 * to onBeforeUnmount. Falls back to a window `resize` listener where
 * ResizeObserver is unavailable, and no-ops in jsdom (zero-width root). */
export function installTierObserver(root: () => Element | null): () => void {
  function measure() {
    const el = root();
    if (el && el.clientWidth > 0) {
      shellWidth.value = el.clientWidth;
    } else if (typeof window !== "undefined" && typeof window.innerWidth === "number") {
      shellWidth.value = window.innerWidth;
    }
  }

  measure();

  if (typeof ResizeObserver !== "undefined") {
    const observer = new ResizeObserver(() => measure());
    const el = root();
    if (el) {
      observer.observe(el);
    }
    return () => observer.disconnect();
  }

  if (typeof window !== "undefined") {
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }

  return () => {};
}

/** Reads the live layout tier. Any component may call this; it shares the one
 * app-wide tier the shell observer maintains, so a component mounted in
 * isolation (unit tests, no observer) simply reports the seeded width's tier. */
export function useTier(): ComputedRef<LayoutTier> {
  return tier;
}

/** Test-only: force the shared width to a known value so a component's tier
 * behaviour can be exercised without a live layout. */
export function setTierWidthForTest(width: number): void {
  shellWidth.value = width;
}
