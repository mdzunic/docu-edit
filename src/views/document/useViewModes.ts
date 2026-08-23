// The document editor's view modes (SPEC-063): Print layout (the default and
// the only mode used for print/export), Web layout (the same paginated model
// as one continuous column), and the transient Focus mode (Print layout with
// the chrome hidden). Print vs Web persists per document; Focus is a view
// state that is never written to storage and always opens off (AC-48, AC-50).
//
// Focus mode hides the menu bar, toolbar and open panes purely with a root
// CSS class — it never closes a pane's own state, so the exact set of open
// panes is restored on exit for free (AC-47).
import { computed, nextTick, ref, watch, type Ref, type ShallowRef } from "vue";
import { cmToPx } from "../../lib/core/units";
import { type DocumentModel, sectionContentWidthCm } from "../../lib/documentModel";
import { usePreferences } from "../../composables/usePreferences";

export type ViewMode = "print" | "web";

export interface ViewModesDeps {
  documentId: Readonly<Ref<string>>;
  /** The open model — Web layout reads the first section's content width. */
  doc: ShallowRef<DocumentModel | null>;
  /** Run after a Print/Web/Focus change has settled in the DOM: the presenter
   * re-tiles the page grid and re-fits an active fit mode (SPEC-063 AC-36). */
  onLayoutSettled?: () => void;
}

/** The per-document persistence key for the print/web view mode (AC-49). */
export function viewModeStorageKey(documentId: string): string {
  return `docu-edit:view-mode:${documentId}`;
}

function readStoredViewMode(documentId: string): ViewMode | null {
  if (!documentId) {
    return null;
  }
  try {
    const raw = localStorage.getItem(viewModeStorageKey(documentId));
    return raw === "web" ? "web" : raw === "print" ? "print" : null;
  } catch {
    return null;
  }
}

function writeStoredViewMode(documentId: string, mode: ViewMode): void {
  if (!documentId) {
    return;
  }
  try {
    localStorage.setItem(viewModeStorageKey(documentId), mode);
  } catch {
    // A storage failure must never break editing.
  }
}

export function useViewModes(deps: ViewModesDeps) {
  const { documentId, doc } = deps;

  const viewMode = ref<ViewMode>("print");
  // Focus is deliberately NOT loaded from or written to storage (AC-48, 50).
  const focusMode = ref(false);
  // The default a new document opens in comes from preferences (SPEC-066).
  const { preferences } = usePreferences();

  // Load the stored print/web mode when the document changes; a document with
  // no stored value opens in the "Default view mode" preference (SPEC-066),
  // Print by default, always with Focus off (AC-48, 51).
  watch(
    documentId,
    (id) => {
      viewMode.value = readStoredViewMode(id) ?? preferences.defaultViewMode;
      focusMode.value = false;
    },
    { immediate: true },
  );

  function setViewMode(mode: ViewMode): void {
    viewMode.value = mode;
    writeStoredViewMode(documentId.value, mode);
  }

  function toggleFocus(): void {
    focusMode.value = !focusMode.value;
  }

  function exitFocus(): void {
    focusMode.value = false;
  }

  // The Print/Web menu radio treats Focus as a third selectable layout even
  // though it is not a persisted view mode (AC-32): entering Focus keeps the
  // underlying print/web mode intact so exiting returns to it.
  const isWebLayout = computed(() => viewMode.value === "web" && !focusMode.value);

  // SPEC-063 AC-40: Web layout constrains the column to the FIRST section's
  // content width (its page width minus its left/right margins), exposed as a
  // CSS var the .is-web-layout rules read. No sections → the page's own width
  // (`auto`); Print layout ignores the var entirely.
  const webContentWidthCss = computed(() => {
    const model = doc.value;
    if (!model || model.sections.length === 0) {
      return "auto";
    }
    return `${cmToPx(sectionContentWidthCm(model, 0))}px`;
  });

  // Toggling Print/Web (or Focus, which is Print) changes grid eligibility and
  // the available width, so re-tile/re-fit once the layout has settled (AC-36).
  watch(
    [viewMode, focusMode],
    () => {
      void nextTick().then(() => deps.onLayoutSettled?.());
    },
    { flush: "post" },
  );

  /** The class list for the editor root (SPEC-063). */
  const rootClass = computed(() => ({
    "is-web-layout": isWebLayout.value,
    "is-focus-mode": focusMode.value,
  }));

  const commands: Array<[string, () => void]> = [
    ["view.layout-print", () => setViewMode("print")],
    ["view.layout-web", () => setViewMode("web")],
    ["view.focus", toggleFocus],
  ];

  return {
    viewMode: viewMode as Readonly<Ref<ViewMode>>,
    focusMode: focusMode as Readonly<Ref<boolean>>,
    isWebLayout,
    /** The `--web-content-width` value for the editor root (AC-40). */
    webContentWidthCss,
    rootClass,
    setViewMode,
    toggleFocus,
    exitFocus,
    commands,
  };
}
