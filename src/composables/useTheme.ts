import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import {
  loadDimPage,
  loadThemePreference,
  saveDimPage,
  saveThemePreference,
  type ThemePreference,
} from "../lib/themePrefs";

// The app-shell theme controller (SPEC-059). A single module-level singleton
// so every view — the library (HomeView), the open-from-URL loader and the
// editor — reads the *same* live state (AC-52/53) and one media-query
// listener drives them all. The Vue app owns every theme change after the
// no-flash script in index.html sets the initial data-theme (AC-60): this
// module is that owner.
//
// It deliberately holds no CSS-var reading of its own beyond resolveThemeColor
// (the canvas helper, AC-61); chrome themes purely through the token
// redefinitions in tokens.css. Persistence lives in lib/themePrefs.ts behind
// the SPEC-003 storage guard (AC-56).

const theme: Ref<ThemePreference> = ref(loadThemePreference());
const dimPage: Ref<boolean> = ref(loadDimPage());
// Tracks the OS setting so `system` follows prefers-color-scheme live (AC-3).
const systemDark = ref(false);

let installed = false;

/** The theme actually in effect — `system` collapses to the OS value. */
const resolvedTheme: ComputedRef<"light" | "dark"> = computed(() =>
  theme.value === "system" ? (systemDark.value ? "dark" : "light") : theme.value,
);

/** Sets data-theme on <html> for an explicit choice, or removes it for
 * `system` so the media query alone decides (AC-4/6). Mirrors the no-flash
 * script in index.html so the app's first application never differs from it. */
function applyDataTheme(): void {
  if (typeof document === "undefined") {
    return;
  }
  const el = document.documentElement;
  if (theme.value === "system") {
    el.removeAttribute("data-theme");
  } else {
    el.setAttribute("data-theme", theme.value);
  }
}

/** Reflects the dim-page flag onto <html>; the CSS treatment (App.vue) is
 * gated on both this and the resolved dark theme, and is print-excluded. */
function applyDimPage(): void {
  if (typeof document === "undefined") {
    return;
  }
  document.documentElement.setAttribute("data-dim-page", dimPage.value ? "1" : "0");
}

/** Installs the singleton's side effects once: the OS-appearance listener
 * (AC-3), the persistence + DOM-application watchers, and the initial apply. */
function install(): void {
  if (installed || typeof window === "undefined") {
    return;
  }
  installed = true;

  // matchMedia is present in every browser the app targets; guard only so the
  // unit-test environment (which omits it) still installs the watchers below.
  if (typeof window.matchMedia === "function") {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    systemDark.value = mq.matches;
    mq.addEventListener("change", (event) => {
      systemDark.value = event.matches;
    });
  }

  watch(theme, () => {
    saveThemePreference(theme.value);
    applyDataTheme();
  });
  watch(dimPage, () => {
    saveDimPage(dimPage.value);
    applyDimPage();
  });

  applyDataTheme();
  applyDimPage();
}

/** The shared theme state and setters. Calling it installs the singleton on
 * first use, so any top-level view that mounts it keeps the OS listener and
 * the persistence watchers alive (AC-52/53). */
export function useTheme() {
  install();
  return {
    theme,
    resolvedTheme,
    dimPage,
    setTheme(next: ThemePreference) {
      theme.value = next;
    },
    setDimPage(on: boolean) {
      dimPage.value = on;
    },
    toggleDimPage() {
      dimPage.value = !dimPage.value;
    },
  };
}

/** Resolves a design token to its computed value on the current theme, for
 * canvas painters that cannot use `var()` (AC-61). Reads
 * getComputedStyle(document.documentElement) so the value already reflects
 * whichever palette data-theme / the media query selected. Returns "" for an
 * unknown token so a caller's fallback stays explicit (AC-62). Accepts a name
 * with or without the leading `--`. */
export function resolveThemeColor(token: string): string {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return "";
  }
  const name = token.startsWith("--") ? token : `--${token}`;
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}
