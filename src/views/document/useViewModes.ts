// The document editor's view modes (SPEC-063): Print layout (the default and
// the only mode used for print/export), Web layout (the same paginated model
// as one continuous column), and the transient Focus mode (Print layout with
// the chrome hidden). Print vs Web persists per document; Focus is a view
// state that is never written to storage and always opens off (AC-48, AC-50).
//
// Focus mode hides the menu bar, toolbar and open panes purely with a root
// CSS class — it never closes a pane's own state, so the exact set of open
// panes is restored on exit for free (AC-47).
import { computed, ref, watch, type Ref } from "vue";

export type ViewMode = "print" | "web";

export interface ViewModesDeps {
  documentId: Readonly<Ref<string>>;
}

/** The per-document persistence key for the print/web view mode (AC-49). */
export function viewModeStorageKey(documentId: string): string {
  return `docu-edit:view-mode:${documentId}`;
}

function readStoredViewMode(documentId: string): ViewMode {
  if (!documentId) {
    return "print";
  }
  try {
    return localStorage.getItem(viewModeStorageKey(documentId)) === "web" ? "web" : "print";
  } catch {
    return "print";
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
  const { documentId } = deps;

  const viewMode = ref<ViewMode>("print");
  // Focus is deliberately NOT loaded from or written to storage (AC-48, 50).
  const focusMode = ref(false);

  // Load the stored print/web mode when the document changes; a document with
  // no stored value opens in Print layout, always with Focus off (AC-48, 51).
  watch(
    documentId,
    (id) => {
      viewMode.value = readStoredViewMode(id);
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
    rootClass,
    setViewMode,
    toggleFocus,
    exitFocus,
    commands,
  };
}
