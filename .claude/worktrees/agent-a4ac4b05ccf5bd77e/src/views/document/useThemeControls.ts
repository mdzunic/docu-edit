import { computed, watch } from "vue";
import { useTheme } from "../../composables/useTheme";
import type { ThemePreference } from "../../lib/themePrefs";
import type { EditorContext } from "./context";

// DocumentView's thin wiring for the app-shell theme feature (SPEC-059). The
// theme state itself is a module singleton in composables/useTheme.ts (shared
// with HomeView and the open-from-URL view, AC-52/53); this composable only
// adapts it for the editor: the menu-prop bag the DocumentMenuBar renders, the
// four editorCommands the menu dispatches (AC-50), and the canvas repaint on a
// resolved-theme change (AC-63). No keyboard shortcut is claimed for any of
// the four (AC-51) — they exist only as menu commands.

export interface ThemeControlsDeps extends EditorContext {
  /** Invalidate the paginated render cache so a theme switch repaints the
   * canvas — the same reflow a page-setup change drives (AC-63). Wired to
   * pagination.invalidateAll(), an existing path, so SPEC-047's scheduler
   * contract is unchanged (AC-64). */
  invalidate: () => void;
}

export function useThemeControls(deps: ThemeControlsDeps) {
  const { theme, resolvedTheme, dimPage, setTheme, toggleDimPage } = useTheme();

  // A theme switch changes what resolveThemeColor() returns for the canvas
  // painters; invalidating repaints them. Only fires on an actual change (the
  // watcher is lazy), never during mount, so it stays clear of the open-phase
  // scheduler ceiling.
  watch(resolvedTheme, () => deps.invalidate());

  /** The props the DocumentMenuBar's Appearance submenu renders from. */
  const themeMenu = computed(() => ({
    theme: theme.value,
    resolvedTheme: resolvedTheme.value,
    dimPage: dimPage.value,
  }));

  const pick = (next: ThemePreference) => () => setTheme(next);

  /** The four command ids the menu emits, for spreading into DocumentView's
   * editorCommands Map (AC-50). */
  const themeCommands: [string, () => void][] = [
    ["view.theme-system", pick("system")],
    ["view.theme-light", pick("light")],
    ["view.theme-dark", pick("dark")],
    ["view.dim-page", () => toggleDimPage()],
  ];

  return { themeMenu, themeCommands };
}
